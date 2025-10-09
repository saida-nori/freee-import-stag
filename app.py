# app.py
"""
楽楽精算→freee仕訳CSV 生成ツール（FastAPI 単一ファイル版）

v2.4 / 最終更新: 2025-10-08
■ 今回の修正
- 貸方摘要：②データ貼付の H列（ﾏﾈ以外申請者氏名）を使用。Hが空なら G列（申請者）を使用。
- 借方税区分：10→「課対仕入10%」／8→「課対仕入8%（軽）」／0→「対象外」に正規化（数値・文字どちらでも可）。
- 出力の重複日付列は作成しない。
- 出力の「日付」は元データの日付の「当月末日」を記載。
- 複合仕訳仕様（借方＝明細行／貸方＝合計1行、残りの貸方金額0）は従来通り。
"""

from fastapi import FastAPI, UploadFile, Request
from fastapi.responses import HTMLResponse, StreamingResponse, RedirectResponse, PlainTextResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

import pandas as pd
import io
import json
from pathlib import Path
from datetime import datetime
import zipfile
import re
import traceback

# ─────────────────────────────────────
# 基本
# ─────────────────────────────────────
BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"
STATIC_DIR = BASE_DIR / "static"
STATIC_DIR.mkdir(exist_ok=True)

APP_VERSION = "v2.4"
APP_DATE = datetime.now().strftime("%Y-%m-%d")

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# ─────────────────────────────────────
# 設定（列名・貸方ルール）
# ─────────────────────────────────────
DEFAULT_CONFIG = {
    "INPUT_SHEET": "②データ貼付",
    # 出力列（重複する日付列は作らない）
    "OUTPUT_COLUMNS": [
        "伝票番号",
        "日付",  # ← 元データの当月末日
        "借方勘定科目", "借方補助科目", "借方部門", "借方税区分", "借方金額", "借方摘要",
        "貸方勘定科目", "貸方補助科目", "貸方部門", "貸方税区分", "貸方金額", "貸方摘要",
        "備考",
    ],
    # 楽楽精算 ②データ貼付 の列名（完全一致）
    "SRC_HEADERS": {
        "date": "日付",                     # 元データの日付（起点）→ 当月末日に変換して出力
        "account": "勘定科目名",
        "subaccount": "補助科目名",
        "dept": "負担部門(選択必須)",
        "tax": "税率",                      # 10/8/0（数値 or 文字）→ 規定文字列へ変換
        "amount": "小計",
        "memo": "自由記入欄",
        "pay_method": "支払方法",
        "card_brand": "カード",
        "ticket_type": "伝票種別",
        # 変更1 対応：貸方摘要用
        "applicant": "申請者",              # G列
        "applicant_not_manager": "ﾏﾈ以外申請者氏名",  # H列
    },
    # 税区分の表記（変更2）
    "TAX_MAP_EXPLICIT": {
        10: "課対仕入10%",
        8:  "課対仕入8%（軽）",
        0:  "対象外",
    },
    # 貸方ルール（カテゴリ別）
    "CREDIT_RULES": {
        "amex":    {"貸方勘定科目": "未払金", "貸方補助科目": "AMEX",     "貸方部門": "本社", "貸方税区分": "対象外"},
        "keihi":   {"貸方勘定科目": "未払金", "貸方補助科目": "従業員立替", "貸方部門": "本社", "貸方税区分": "対象外"},
        "kotsuhi": {"貸方勘定科目": "未払金", "貸方補助科目": "従業員立替", "貸方部門": "本社", "貸方税区分": "対象外"},
    },
}

def load_config() -> dict:
    if CONFIG_PATH.exists():
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_CONFIG

def save_config(data: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

CONFIG = load_config()

# ─────────────────────────────────────
# 小物
# ─────────────────────────────────────
def map_tax_value(val):
    """変更2：10/8/0 を指定の文字列に。それ以外は値をそのまま返す（必要なら追記で拡張）。"""
    explicit = CONFIG["TAX_MAP_EXPLICIT"]
    if pd.isna(val):
        return "対象外"
    s = str(val).strip()
    # 数値化を試みる
    try:
        iv = int(float(s))
        if iv in explicit:
            return explicit[iv]
    except Exception:
        pass
    # "10%" のような表記が来た時でも拾える程度に
    if "10" == s or s == "10.0" or s.startswith("10%"):
        return explicit[10]
    if "8" == s or s == "8.0" or "8" in s:
        return explicit[8]
    if s in {"0", "0.0", "対象外", "非対象"}:
        return explicit[0]
    # デフォルト
    return s

def _coerce_str(s: pd.Series) -> pd.Series:
    return s.astype(str).replace({"nan": "", "NaT": ""}).fillna("")

def _read_csv_safely(file_bytes: bytes) -> pd.DataFrame:
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
        except Exception:
            continue
    return pd.read_csv(io.BytesIO(file_bytes), encoding_errors="ignore")

def _read_base(file_bytes: bytes, filename: str) -> pd.DataFrame:
    name = (filename or "").lower()
    if name.endswith(".csv"):
        return _read_csv_safely(file_bytes)
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = CONFIG.get("INPUT_SHEET", "②データ貼付")
    if sheet not in xls.sheet_names:
        cand = [s for s in xls.sheet_names if "データ貼" in s]
        sheet = cand[0] if cand else xls.sheet_names[0]
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)

def _col_letter_to_idx(col: str) -> int:
    col = col.strip().upper()
    val = 0
    for ch in col:
        if "A" <= ch <= "Z":
            val = val * 26 + (ord(ch) - ord("A") + 1)
    return val - 1

def _pick_by_letter(row: pd.Series, letters: list[str]) -> list[str]:
    values = []
    for col in letters:
        idx = _col_letter_to_idx(col)
        v = row.iloc[idx] if 0 <= idx < len(row) else ""
        values.append("" if pd.isna(v) else str(v))
    return values

def _join_clean(parts: list[str], sep: str = " ") -> str:
    parts = [p for p in [p.strip() for p in parts] if p]
    return sep.join(parts)

def _format_mmdd(val) -> str:
    """C列の値などを 月/日(MM/DD) に変換。日付でなければそのまま。"""
    try:
        dt = pd.to_datetime(val, errors="coerce")
        if pd.isna(dt):
            m = re.search(r"(\d{1,2})[/-](\d{1,2})", str(val))
            if m:
                return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}"
            return str(val)
        return dt.strftime("%m/%d")
    except Exception:
        return str(val)

def _end_of_month_str(val) -> str:
    """変更4：元データの日付から当月末日を YYYY-MM-DD で返す。"""
    dt = pd.to_datetime(val, errors="coerce")
    if pd.isna(dt):
        return ""
    # pandas で月末日に置換
    return (dt + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")

# ─────────────────────────────────────
# ②データ貼付 → カテゴリ振り分け
# ─────────────────────────────────────
def split_categories(df: pd.DataFrame):
    h = CONFIG["SRC_HEADERS"]
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    amex = pd.Series(False, index=df.index)
    if h.get("pay_method") in df.columns:
        amex = amex | _coerce_str(df[h["pay_method"]]).str.contains("AMEX|アメックス", case=False, na=False)
    if h.get("card_brand") in df.columns:
        amex = amex | _coerce_str(df[h["card_brand"]]).str.contains("AMEX|アメックス", case=False, na=False)

    kotsu = pd.Series(False, index=df.index)
    if h.get("ticket_type") in df.columns:
        kotsu = _coerce_str(df[h["ticket_type"]]).str.contains("交通費", na=False)

    keihi = (~amex) & (~kotsu)

    return df[amex].copy(), df[keihi].copy(), df[kotsu & (~amex)].copy()

# ─────────────────────────────────────
# 摘要（借方・貸方）
# ─────────────────────────────────────
# 借方摘要：従来の「列の組合せ」ルール（必要に応じて調整可）
_DEBIT_MEMO_PARTS = {
    "amex":    ["G", "C", "P"],
    "keihi":   ["G", "C", "AM", "P"],
    "kotsuhi": ["G", "C", "M", "K", "P"],
}
def build_debit_memo_series(df: pd.DataFrame, kind: str) -> pd.Series:
    parts = _DEBIT_MEMO_PARTS[kind]
    memos = []
    for _, row in df.reset_index(drop=True).iterrows():
        vals = _pick_by_letter(row, parts)
        for i, letter in enumerate(parts):
            if letter.upper() == "C":
                vals[i] = _format_mmdd(vals[i])
        memos.append(_join_clean(vals, " "))
    return pd.Series(memos, index=df.index)

# 変更1：貸方摘要＝H列（ﾏﾈ以外申請者氏名）、Hが空ならG列（申請者）
def build_credit_memo_series(df: pd.DataFrame) -> pd.Series:
    memos = []
    for _, row in df.reset_index(drop=True).iterrows():
        # H列 → G列（フォールバック）
        h_val = _pick_by_letter(row, ["H"])[0]
        g_val = _pick_by_letter(row, ["G"])[0]
        memo = h_val.strip() if h_val.strip() else g_val.strip()
        memos.append(memo)
    return pd.Series(memos, index=df.index)

# ─────────────────────────────────────
# 複合仕訳の生成（借方：明細／貸方：合計1行、残り0）
# ─────────────────────────────────────
def build_compound_voucher(df: pd.DataFrame, kind: str, voucher_id: str) -> pd.DataFrame:
    h = CONFIG["SRC_HEADERS"]
    # 借方摘要（従来ルール）／貸方摘要（変更1ルール）
    memo_debit = build_debit_memo_series(df, kind)
    memo_credit = build_credit_memo_series(df)

    # 変更4：当月末日に変換した日付
    eom = df[h["date"]].apply(_end_of_month_str)

    deb = pd.DataFrame({
        "伝票番号": voucher_id,
        "日付": eom,  # ← 当月末日
        "借方勘定科目": df[h["account"]],
        "借方補助科目": df[h["subaccount"]] if h["subaccount"] in df.columns else "",
        "借方部門": df[h["dept"]] if h["dept"] in df.columns else "",
        "借方税区分": df[h["tax"]].apply(map_tax_value) if h["tax"] in df.columns else "対象外",  # 変更2
        "借方金額": pd.to_numeric(df[h["amount"]], errors="coerce"),
        "借方摘要": memo_debit,
    })

    credit = CONFIG["CREDIT_RULES"][kind]
    total = deb["借方金額"].sum(skipna=True)
    credit_amounts = [total] + [0] * (len(deb) - 1)

    cred = pd.DataFrame({
        "貸方勘定科目": credit["貸方勘定科目"],
        "貸方補助科目": credit.get("貸方補助科目", ""),
        "貸方部門": credit.get("貸方部門", ""),
        "貸方税区分": credit.get("貸方税区分", "対象外"),
        "貸方金額": credit_amounts,
        "貸方摘要": memo_credit,  # 変更1
    }, index=deb.index)

    out = pd.concat([deb, cred], axis=1)
    out["備考"] = ""
    out = out.reindex(columns=CONFIG["OUTPUT_COLUMNS"], fill_value="")
    # 金額NaNは落とす
    out = out[pd.notna(out["借方金額"])]
    return out

# ─────────────────────────────────────
# 画面（最小限：トップ／設定／マニュアルは省略可）
# ─────────────────────────────────────
INDEX_HTML = f"""<!DOCTYPE html>
<html lang="ja"><head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>楽楽精算→freee仕訳CSV 生成ツール</title>
<style>
  body {{ font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Hiragino Kaku Gothic ProN', 'Noto Sans JP', sans-serif; margin: 40px; }}
  header {{ margin-bottom: 24px; display:flex; align-items:center; gap:16px; }}
  .title {{ display:flex; flex-direction:column; }}
  .card {{ border: 1px solid #e5e7eb; border-radius: 12px; padding: 20px; margin-bottom: 20px; }}
  .row {{ display: flex; gap: 16px; align-items: center; flex-wrap: wrap; }}
  input[type=file] {{ padding: 8px; }}
  button {{ padding: 10px 16px; border: 1px solid #111827; background: #111827; color: white; border-radius: 8px; cursor: pointer; }}
  button:hover {{ opacity: 0.9; }}
  .hint {{ color: #6b7280; font-size: 13px; }}
  .logo {{ height:60px; }}
  .version {{ font-size: 12px; color:#555; margin-top:4px; }}
</style></head>
<body>
  <header>
    <img src="/static/logo.png" alt="Company Logo" class="logo" />
    <div class="title">
      <h2>楽楽精算→freee仕訳CSV 生成ツール</h2>
      <div class="version">{APP_VERSION} - {APP_DATE}</div>
    </div>
  </header>

  <div class="card">
    <form action="/convert" method="post" enctype="multipart/form-data">
      <div class="row">
        <input type="file" name="file" accept=".csv,.xlsx,.xlsm,.xls" required />
        <button type="submit">変換してダウンロード</button>
      </div>
      <p class="hint">入力は <b>②データ貼付</b>（CSV/Excel）。出力は複合仕訳（借方＝明細／貸方＝合計1行・他0）。日付は当月末日。</p>
    </form>
  </div>
</body></html>
"""

@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(INDEX_HTML)

# ─────────────────────────────────────
# 変換API：②データ貼付のみ解析 → カテゴリ別の複合仕訳CSVをZIP出力
# ─────────────────────────────────────
@app.post("/convert")
async def convert(file: UploadFile):
    try:
        raw = await file.read()
        df = _read_base(raw, file.filename or "uploaded")

        amex_df, keihi_df, kotsu_df = split_categories(df)

        outputs = {}
        seq = 1
        today_key = datetime.now().strftime("%Y%m%d")

        if not amex_df.empty:
            outputs["amex"] = build_compound_voucher(amex_df, "amex",  f"AMEX-{today_key}-{seq:03d}"); seq += 1
        if not keihi_df.empty:
            outputs["keihi"] = build_compound_voucher(keihi_df, "keihi", f"KEIHI-{today_key}-{seq:03d}"); seq += 1
        if not kotsu_df.empty:
            outputs["kotsuhi"] = build_compound_voucher(kotsu_df, "kotsuhi", f"KOTSU-{today_key}-{seq:03d}"); seq += 1

        mem = io.BytesIO()
        with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            wrote = 0
            for kind, df_out in outputs.items():
                if df_out.empty: continue
                buf = io.StringIO(); df_out.to_csv(buf, index=False)
                zf.writestr(f"{kind}_journal_freee.csv", buf.getvalue().encode("utf-8-sig"))
                wrote += 1
            if wrote:
                merged = pd.concat([d for d in outputs.values() if not d.empty], ignore_index=True)
                buf = io.StringIO(); merged.to_csv(buf, index=False)
                zf.writestr("merged_all_freee.csv", buf.getvalue().encode("utf-8-sig"))
            else:
                zf.writestr("README.txt", "②データ貼付から振り分けできませんでした。列名や判定列をご確認ください。".encode("utf-8"))

        mem.seek(0)
        filename = f"freee_journals_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        return StreamingResponse(mem, media_type="application/zip",
                                 headers={"Content-Disposition": f"attachment; filename={filename}"})
    except Exception:
        return PlainTextResponse("Convert error:\n" + traceback.format_exc(), status_code=500)

# ─────────────────────────────────────
# Entrypoint
# ─────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
