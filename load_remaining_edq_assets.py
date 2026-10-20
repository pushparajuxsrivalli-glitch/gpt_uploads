#!/usr/bin/env python3
"""Load the 18 remaining selected EDQ workbook sheets into PostgreSQL.

The already-loaded 03_ISS_account_status asset is excluded. Existing tables are
never overwritten. Set DB_PASSWORD before running.
"""
import os
import re
import sys
import pandas as pd
from sqlalchemy import create_engine, inspect, text
from sqlalchemy.types import Boolean, Date, DateTime, BigInteger, Numeric, Text

EXCEL_FILE = os.path.expanduser('~/edq-project/PCN_DQ_TestDataset (1).xlsx')
DB_HOST = os.getenv('DB_HOST', 'localhost')
DB_PORT = os.getenv('DB_PORT', '5432')
DB_NAME = os.getenv('DB_NAME', 'edq_test_data')
DB_USER = os.getenv('DB_USER', 'edq_app')
DB_PASSWORD = os.getenv('DB_PASSWORD','edq_app')

# 03_ISS_account_status is already loaded and intentionally omitted.
SHEETS = [
    '03_ISS_issuer_account',
    '01_TXN_transaction_reversal',
    '01_TXN_transaction_lifecycle',
    '03_ISS_issuer_relationship',
    '03_ISS_issuer',
    '02_CRD_card_bin',
    '02_CRD_card_expiration',
    '02_CRD_card_replacement',
    '02_CRD_card_issuance',
    '02_CRD_card_activation',
    '02_CRD_card_program',
    '02_CRD_card_product',
    '02_CRD_card_number',
    '02_CRD_card_status',
    '02_CRD_card_account',
    '02_CRD_card',
    '01_TXN_transaction_cancellation',
    '01_TXN_transaction_status',
]

# Keep identifier-like values as text, including leading zeros.
ID_COLUMNS = {x.upper() for x in '''STATUS_ID ACCOUNT_ID ISSUER_ACCOUNT_ID ISSUER_ID
RELATIONSHIP_ID PARTNER_ID REVERSAL_ID ORIGINAL_TXN_ID TRANSACTION_ID
CANCELLATION_ID EXPIRATION_ID CARD_ID REPLACEMENT_ID OLD_CARD_ID NEW_CARD_ID
ISSUANCE_ID DELIVERY_ADDRESS_ID BATCH_ID ACTIVATION_ID PROGRAM_ID PRODUCT_ID
CARD_NUMBER_ID STATUS_RECORD_ID CARD_ACCOUNT_ID CARDHOLDER_ID RENEWAL_CARD_ID
ACCOUNT_NUMBER PAN_FULL PAN_MASKED NETWORK_BIN BIN ROUTING_NUMBER SWIFT_CODE
TOKENIZED_PAN PAN_HASH'''.split()}
MONEY_COLUMNS = {x.upper() for x in '''BALANCE ORIGINAL_AMOUNT REVERSAL_AMOUNT
AMOUNT REFUND_AMOUNT ANNUAL_FEE CREDIT_LIMIT AVAILABLE_BALANCE'''.split()}
BOOLEAN_COLUMNS = {x.upper() for x in '''IS_ACTIVE IS_CROSS_BORDER NOTIFICATION_SENT
AUTO_RENEWAL OLD_CARD_BLOCKED IS_SAME_CARD IS_INSTANT_ISSUE IS_VALID_LENGTH
VALID_TRANSITION IS_OVERLIMIT IS_TOKENIZED CVV_PRESENT REFUND_TRIGGERED
FINAL_STATUS'''.split()}
DATE_COLUMNS = {x.upper() for x in '''EFFECTIVE_DATE START_DATE END_DATE
REGISTERED_DATE NOTIFICATION_DATE REPLACEMENT_DATE ISSUED_DATE ACTIVATION_DATE
LAUNCH_DATE DISCONTINUE_DATE OPEN_DATE'''.split()}
TIMESTAMP_COLUMNS = {x.upper() for x in '''REVERSAL_TIMESTAMP INITIATED_AT
AUTHORIZED_AT CLEARED_AT SETTLED_AT CANCELLED_AT STATUS_TIMESTAMP CHANGED_AT'''.split()}
# These remain text, not dates or numbers.
TEXT_COLUMNS = {x.upper() for x in '''DQ_FLAG EXPIRY_DATE PAN_LAST4 STATUS
PREVIOUS_STATUS STATUS_CODE REASON_CODE'''.split()}
TRUE_VALUES = {'Y', 'YES', 'TRUE', 'T', '1'}
FALSE_VALUES = {'N', 'NO', 'FALSE', 'F', '0', ''}


def to_bool(value, sheet, column):
    if pd.isna(value):
        return None
    if isinstance(value, bool):
        return value
    val = str(value).strip().upper()
    if val in TRUE_VALUES:
        return True
    if val in FALSE_VALUES:
        return False
    raise ValueError(f'{sheet}.{column}: unrecognized Boolean value {value!r}; no table loaded for this sheet')


def prepare(df, sheet):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # Empty/whitespace-only Excel cells become SQL NULL.
    df = df.replace(r'^\s*$', None, regex=True)
    for col in df.columns:
        key = col.upper()
        if key in BOOLEAN_COLUMNS:
            df[col] = df[col].map(lambda v: to_bool(v, sheet, col)).astype('boolean')
        elif key in MONEY_COLUMNS:
            df[col] = pd.to_numeric(df[col], errors='raise')
        elif key in DATE_COLUMNS:
            df[col] = pd.to_datetime(df[col], errors='raise').dt.date
        elif key in TIMESTAMP_COLUMNS:
            parsed = pd.to_datetime(df[col], errors='raise', utc=True)
            # Explicit default assumption for timezone-naive source timestamps: UTC.
            # Override with SOURCE_TIMEZONE if the workbook uses another known zone.
            source_tz = os.getenv('SOURCE_TIMEZONE', 'UTC')
            original = pd.to_datetime(df[col], errors='raise')
            if original.dt.tz is None:
                parsed = original.dt.tz_localize(source_tz).dt.tz_convert('UTC')
            df[col] = parsed
        elif key in ID_COLUMNS or key in TEXT_COLUMNS or key in {'COUNTRY_CODE', 'CURRENCY_CODE', 'NETWORK', 'CARD_TYPE', 'CARD_STATUS', 'ACCOUNT_TYPE', 'ACCOUNT_STATUS', 'PROGRAM_TYPE', 'PRODUCT_TYPE', 'RELATIONSHIP_TYPE', 'LIFECYCLE_STAGE', 'TRANSACTION_TYPE', 'CHANNEL', 'FUNDING_TYPE', 'DELIVERY_METHOD', 'DELIVERY_STATUS', 'ACTIVATION_CHANNEL', 'ACTIVATION_STATUS', 'REVERSAL_TYPE', 'REVERSAL_REASON', 'CANCELLATION_REASON', 'STATUS_DESCRIPTION', 'STATUS_CATEGORY', 'RESPONSE_SOURCE', 'INITIATED_BY', 'CANCELLED_BY', 'CHANGED_BY', 'CHANGE_REASON'}:
            df[col] = df[col].map(lambda v: None if pd.isna(v) else str(v))
    return df


def sql_types(df):
    mapping = {}
    for col in df.columns:
        key = col.upper()
        if key == 'DQ_FLAG':
            mapping[col] = Text()             # preserve values as text
        elif key in MONEY_COLUMNS:
            mapping[col] = Numeric(18, 4)     # agreed precision
        elif key in BOOLEAN_COLUMNS:
            mapping[col] = Boolean()          # agreed Boolean mapping
        elif key in TIMESTAMP_COLUMNS:
            mapping[col] = DateTime(timezone=True)  # TIMESTAMPTZ
        elif key in DATE_COLUMNS:
            mapping[col] = Date()
        elif key in ID_COLUMNS or key in TEXT_COLUMNS:
            mapping[col] = Text()
        elif pd.api.types.is_integer_dtype(df[col].dtype):
            mapping[col] = BigInteger()
        elif pd.api.types.is_float_dtype(df[col].dtype):
            mapping[col] = Numeric()
        else:
            # Unknown/object fields remain text rather than guessing semantics.
            mapping[col] = Text()
    return mapping


def main():
    if not DB_PASSWORD:
        sys.exit('ERROR: Set DB_PASSWORD in the shell before running.')
    if not os.path.isfile(EXCEL_FILE):
        sys.exit(f'ERROR: Workbook not found: {EXCEL_FILE}')

    engine = create_engine(
        f'postgresql+psycopg://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}'
    )
    existing = set(inspect(engine).get_table_names(schema='public'))
    loaded, skipped, failed = [], [], []

    with pd.ExcelFile(EXCEL_FILE, engine='openpyxl') as workbook:
        available = set(workbook.sheet_names)
        for sheet in SHEETS:
            table = re.sub(r'[^a-zA-Z0-9_]', '_', sheet).lower()
            if sheet not in available:
                failed.append((sheet, 'Sheet not found'))
                print(f'[FAIL] {sheet}: sheet not found')
                continue
            if table in existing:
                skipped.append((sheet, table))
                print(f'[SKIP] public.{table} already exists; not overwritten')
                continue
            try:
                df = prepare(pd.read_excel(workbook, sheet_name=sheet), sheet)
                df.to_sql(table, engine, schema='public', if_exists='fail', index=False,
                          dtype=sql_types(df), method='multi', chunksize=500)
                with engine.connect() as conn:
                    rows = conn.execute(text(f'SELECT COUNT(*) FROM public."{table}"')).scalar_one()
                if rows != len(df):
                    raise RuntimeError(f'row count mismatch: Excel={len(df)}, PostgreSQL={rows}')
                loaded.append((sheet, table, rows, len(df.columns)))
                existing.add(table)
                print(f'[OK] {sheet} -> public.{table}: {rows} rows, {len(df.columns)} columns')
            except Exception as exc:
                failed.append((sheet, str(exc)))
                print(f'[FAIL] {sheet}: {exc}')

    print('\n===== SUMMARY =====')
    print(f'Loaded: {len(loaded)} | Skipped existing: {len(skipped)} | Failed: {len(failed)}')
    for row in skipped:
        print(f'SKIPPED: {row[0]} -> {row[1]}')
    for row in failed:
        print(f'FAILED: {row[0]} -> {row[1]}')
    engine.dispose()
    if failed:
        sys.exit(1)

if __name__ == '__main__':
    main()
