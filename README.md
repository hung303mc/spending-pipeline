# Spending Pipeline

Tự động cào email VCB + TPBank → parse → build Google Sheet → gửi email summary hàng tháng.

## Run local

```bash
# 1. Setup
pip install -r requirements.txt
python auth.py                 # one-time OAuth (combined scopes)

# 2. Run
python spending_pipeline.py                          # auto previous month
python spending_pipeline.py --month 2026-04          # specific month
python spending_pipeline.py --month 2026-04 --include-quarter
python spending_pipeline.py --no-email               # skip email
```

## Run in Claude Code cloud routine

Env vars required:
- `TOKEN_JSON` — content of `token.json` (combined scopes)
- `SHEET_ID` — Google Sheet master (default baked in)
- `NOTIFY_EMAIL` — email destination
- `OVERRIDES_JSON` — (optional) per-txn overrides

Cron: `0 2 5 * *` (mùng 5 hàng tháng 9:00 VN time = 02:00 UTC).

Auto-includes quarterly report when target month ∈ {Mar, Jun, Sep, Dec}.

## Sources parsed

| Source | Pattern | Format |
|---|---|---|
| VCB transfers | `from:VCBDigibank@info.vietcombank.com.vn` | Plaintext body, regex |
| VCB cards | `from:info@info.vietcombank.com.vn subject:"giao dịch thẻ"` | Plaintext body, regex |
| TPBank credit cards (Visa 7606 + JCB 8272) | `from:tpbank@tpb.com.vn subject:"giao dịch thẻ tín dụng"` | HTML→text, fields with `:\n` separator |
| TPBank TKTT (chuyển khoản) | `subject:"sao kê tài khoản"` (gửi đầu tháng kế tiếp) | PDF table via pdfplumber |

## Double-count protection

- `TT QUA TPBANK EBANKING` rows in TKTT statement = trả dư nợ thẻ tín dụng → auto-exclude.
- Custom overrides in `overrides.json` for internal transfers between accounts.

## Categorization

Rule-based in `RULES` array — pattern → category. Adjust by editing source.

Per-transaction overrides via `overrides.json`:
```json
{
  "203V009260972957": {
    "category": "Chuyển nội bộ (loại)",
    "exclude": true,
    "user_note": "CK nội bộ sang VCB"
  }
}
```
