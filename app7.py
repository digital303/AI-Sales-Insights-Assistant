import os

import re

import hashlib

import traceback
import time
import random

from io import BytesIO

import xlsxwriter

from pathlib import Path

import pandas as pd

import streamlit as st

from dotenv import load_dotenv

from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import ChatPromptTemplate

load_dotenv()

st.set_page_config(page_title='AI Sales Insights Assistant', layout='wide')

st.title('AI Sales Insights Assistant')

st.caption('Laptop reports | Orders • Ads • FBA • Returns • Traffic | Gemini AI')

TYPES = ['Orders', 'Ads', 'FBA Inventory', 'Customer Returns', 'Traffic', 'Other Data']

ALIASES = {

    'sku': ['sku', 'seller sku', 'seller-sku', 'msku', 'merchant sku', 'advertised sku',

            'seller-sku-id', 'merchant-sku-id', 'child asin', '(child) asin', 'asin'],

    'category': ['category', 'product category', 'sub category', 'subcategory', 'category name'],

    'zone': ['zone', 'warehouse', 'fulfillment center', 'fulfilment center', 'fulfillment-center-id', 'fc', 'location'],

    'fnsku': ['fnsku'],

    'disposition': ['disposition'],

    'fba_ending': ['ending-warehouse-balance'],

    'fba_starting': ['starting-warehouse-balance'],

    'fba_receipts': ['receipts'],

    'fba_shipments': ['customer-shipments'],

    'fba_returns': ['customer-returns'],

    'not_in_report': [],

    'date': ['purchase date', 'purchase-date', 'date', 'report date', 'snapshot date',

             'snapshot-date', 'return date', 'return request date', 'refund date',

             'order date', 'start date', 'end date', 'day'],

    'order_id': ['amazon order id', 'amazon-order-id', 'order id'],

    'item_id': ['order item id', 'order-item-id'],

    'status': ['order status', 'order-status'],

    'revenue': ['item price', 'item-price', 'item sales', 'ordered product sales',

                'ordered-product-sales', 'ordered product sales amount'],

    'units': ['quantity', 'quantity ordered', 'units ordered', 'units-ordered', 'units sold'],

    'spend': ['spend', 'cost', 'total cost', 'ad spend'],

    'ad_sales': ['14 day total sales', '7 day total sales', '14 day advertised sku sales',

                 '7 day advertised sku sales', 'attributed sales', 'sales', 'total sales'],

    'ad_orders': ['14 day total orders', '7 day total orders', '14 day total purchases',

                  '7 day total purchases', 'purchases', 'orders'],

    'impressions': ['impressions'],

    'clicks': ['clicks'],

    'campaign': ['campaign name', 'campaign', 'campaign name (informational only)'],

    'available': ['available', 'available-inventory', 'available-quantity',

                  'afn-fulfillable-quantity', 'fulfillable-quantity', 'fulfillable',

                  'available-stock', 'available-inventory-quantity'],

    'on_hand': ['on-hand', 'on-hand-quantity', 'total-inventory',

                'total-inventory-quantity', 'inventory', 'inventory-quantity',

                'total-stock', 'afn-total-quantity', 'total-quantity'],

    'inbound': ['inbound', 'inbound-quantity', 'afn-inbound-working-quantity',

                'afn-inbound-shipped-quantity', 'afn-inbound-receiving-quantity',

                'inbound-working', 'inbound-shipped', 'inbound-receiving'],

    'returned': ['return quantity', 'quantity returned', 'returned quantity', 'return units', 'quantity'],

    'refund': ['refund amount', 'refunded amount', 'refund value', 'refund amount inr', 'refund total'],

    'sessions': ['sessions total', 'sessions - total', 'sessions', 'total sessions'],

    'page_views': ['page views total', 'page views - total', 'page views', 'pageviews'],

    'traffic_units': ['units ordered', 'units ordered - total', 'units ordered total', 'ordered units'],

    'traffic_sales': ['ordered product sales', 'ordered product sales - total',

                      'ordered product sales total', 'ordered product sales amount'],

}

METRICS = {

    'Orders': ['Revenue', 'Units'],

    'Ads': ['Ad Sales', 'Spend', 'Ad Orders', 'Clicks', 'Impressions'],

    'FBA Inventory': ['Ending Balance', 'Starting Balance', 'Receipts', 'Customer Shipments', 'Customer Returns', 'Sellable Ending Balance', 'Inbound Quantity'],

    'Customer Returns': ['Refund Value', 'Returned Units'],

    'Traffic': ['Ordered Sales', 'Units Ordered', 'Sessions', 'Page Views'],

}

MONEY = {'Revenue', 'Spend', 'Ad Sales', 'Refund Value', 'Ordered Sales'}

def norm(value):

    return re.sub(r'[^a-z0-9]+', '-', str(value).strip().lower()).strip('-')

def col(df, name):

    return next((norm(a) for a in ALIASES[name] if norm(a) in df.columns), None)

def texts(df, name, default='Unknown'):

    c = col(df, name)

    if c is None:

        return pd.Series(default, index=df.index, dtype='string')

    return df[c].astype('string').fillna(default).str.strip().replace('', default)

def nums(df, name):

    c = col(df, name)

    if c is None:

        return pd.Series(0.0, index=df.index)

    s = df[c].astype('string').fillna('').str.replace(',', '', regex=False)

    s = s.str.replace(r'[^0-9.\-]', '', regex=True)

    return pd.to_numeric(s, errors='coerce').fillna(0.0)

def filename_date(filename):

    found = re.findall(r'(?<!\d)(20\d{6})(?!\d)', filename)

    if not found:

        return pd.NaT, False

    return pd.to_datetime(found[-1], format='%Y%m%d', errors='coerce'), len(found) > 1

def report_dates(df, filename):
    """Read individual report dates; date-only DD/MM/YYYY is supported."""
    c = col(df, 'date')
    if c is None:
        local = pd.Series(pd.NaT, index=df.index, dtype='datetime64[ns]')
    else:
        def parse_one(raw):
            if pd.isna(raw) or not str(raw).strip():
                return pd.NaT
            value = str(raw).strip()
            try:
                # Prefer day-first for Indian slash-formatted dates, but never
                # reinterpret ISO YYYY-MM-DD as DD-MM-YYYY.
                if re.fullmatch(r'\d{1,2}[/]\d{1,2}[/]\d{4}', value):
                    parsed = pd.to_datetime(value, dayfirst=True, errors='coerce')
                elif re.fullmatch(r'\d{1,2}-\d{1,2}-\d{4}', value):
                    parsed = pd.to_datetime(value, dayfirst=True, errors='coerce')
                else:
                    parsed = pd.to_datetime(value, errors='coerce')
                if pd.isna(parsed):
                    return pd.NaT
                parsed = pd.Timestamp(parsed)
                if parsed.tzinfo is not None:
                    parsed = parsed.tz_convert('Asia/Kolkata').tz_localize(None)
                return parsed.normalize()
            except (ValueError, TypeError, OverflowError):
                return pd.NaT

        local = pd.to_datetime(df[c].map(parse_one), errors='coerce')

    fallback, is_range = filename_date(filename)
    used_filename = pd.notna(fallback) and local.isna().any()
    if pd.notna(fallback):
        local = local.fillna(fallback)
    return local, bool(is_range and used_filename)


@st.cache_data(show_spinner=False, max_entries=40)

def read_file(blob, filename):

    suffix = Path(filename).suffix.lower()

    if suffix in ('.xlsx', '.xls'):

        frame = pd.read_excel(BytesIO(blob), sheet_name=0, dtype=str)

    else:

        frame = None

        for encoding in ('utf-8-sig', 'cp1252'):

            try:

                frame = pd.read_csv(BytesIO(blob), sep=None, engine='python',

                                    encoding=encoding, dtype=str)

                break

            except UnicodeDecodeError:

                pass

        if frame is None:

            raise ValueError('Unable to decode the text file')

    frame.columns = [norm(x) for x in frame.columns]

    frame = frame.loc[:, ~pd.Index(frame.columns).duplicated()].copy()

    return frame

def prepare(raw, kind, filename):

    if kind == 'Other Data':

        return raw.copy(), []

    minimum = {

        'Orders': [('sku',), ('units',), ('revenue',)],

        'Ads': [('spend',)],

        'FBA Inventory': [('fba_ending', 'available', 'on_hand')],

        'Customer Returns': [('returned', 'refund')],

        'Traffic': [('sessions', 'page_views', 'traffic_units', 'traffic_sales')],

    }

    missing = [group for group in minimum[kind] if not any(col(raw, x) for x in group)]

    if missing:

        raise ValueError('Missing fields: ' + ', '.join('/'.join(g) for g in missing))

    df = raw.copy()

    if kind == 'Orders':

        active = ~texts(df, 'status', 'active').str.lower().isin(

            ['cancelled', 'canceled', 'pending'])

        df = df.loc[active].copy()

        item = col(df, 'item_id')

        if item:

            present = df[item].notna() & df[item].str.strip().ne('')

            keys = [item] + ([col(df, 'order_id')] if col(df, 'order_id') else [])

            df = pd.concat([df.loc[present].drop_duplicates(keys), df.loc[~present]])

    out = pd.DataFrame(index=df.index)

    out['SKU'] = texts(df, 'sku')

    out['Category'] = texts(df, 'category')

    unknown = out['Category'].eq('Unknown') & out['SKU'].ne('Unknown')

    out.loc[unknown, 'Category'] = out.loc[unknown, 'SKU'].str[:7]

    out['Zone'] = texts(df, 'zone')

    out['FNSKU'] = texts(df, 'fnsku')

    out['Disposition'] = texts(df, 'disposition')

    out['Campaign'] = texts(df, 'campaign')

    out['Date'], range_date = report_dates(df, filename)

    out['Source'] = filename

    fields = {

        'Orders': {'Revenue': 'revenue', 'Units': 'units'},

        'Ads': {'Spend': 'spend', 'Ad Sales': 'ad_sales', 'Ad Orders': 'ad_orders',

                'Clicks': 'clicks', 'Impressions': 'impressions'},

        'FBA Inventory': ({'Ending Balance': 'fba_ending', 'Starting Balance': 'fba_starting',

                           'Receipts': 'fba_receipts', 'Customer Shipments': 'fba_shipments',

                           'Customer Returns': 'fba_returns'} if col(df, 'fba_ending')

                          else {'Ending Balance': 'on_hand', 'Starting Balance': 'not_in_report',

                                'Receipts': 'not_in_report', 'Customer Shipments': 'not_in_report',

                                'Customer Returns': 'not_in_report'}),

        'Customer Returns': {'Returned Units': 'returned', 'Refund Value': 'refund'},

        'Traffic': {'Sessions': 'sessions', 'Page Views': 'page_views',

                    'Units Ordered': 'traffic_units', 'Ordered Sales': 'traffic_sales'},

    }

    for label, key in fields[kind].items():

        out[label] = nums(df, key)

    if kind == 'FBA Inventory':

        # Preserve whether receipts were present in the source movement report.
        # A missing receipts column is not the same as zero receipts.
        out['Receipts Reported'] = col(df, 'fba_receipts') is not None

        if col(df, 'fba_ending'):

            # This movement report does not contain an Amazon 'available inventory' field.

            # SELLABLE ending balance is a separate, clearly labeled subset.

            out['Sellable Ending Balance'] = out['Ending Balance'].where(

                out['Disposition'].str.upper().isin(['SELLABLE', 'SELLABLE-INVENTORY']), 0.0)

        else:

            out['Sellable Ending Balance'] = nums(df, 'available')

        # Inbound units are snapshot values, never sum across historical dates.
        # Prefer an explicit overall inbound quantity; otherwise sum status buckets.
        total_inbound_column = next((c for c in (
            'inbound', 'inbound-quantity', 'afn-inbound-quantity'
        ) if c in df.columns), None)
        inbound_status_columns = [c for c in (
            'afn-inbound-working-quantity', 'afn-inbound-shipped-quantity',
            'afn-inbound-receiving-quantity', 'inbound-working',
            'inbound-shipped', 'inbound-receiving'
        ) if c in df.columns]
        if total_inbound_column:
            out['Inbound Quantity'] = pd.to_numeric(
                df[total_inbound_column].astype('string').str.replace(',', '', regex=False),
                errors='coerce')
        elif inbound_status_columns:
            out['Inbound Quantity'] = sum(
                (pd.to_numeric(df[c].astype('string').str.replace(',', '', regex=False),
                               errors='coerce').fillna(0) for c in inbound_status_columns),
                start=pd.Series(0.0, index=df.index)
            )
        else:
            out['Inbound Quantity'] = float('nan')

    if kind == 'Orders':

        out['Order ID'] = texts(df, 'order_id', '')

        out['Order Item ID'] = texts(df, 'item_id', '')

    warnings = []

    if range_date:

        warnings.append(f'{filename}: reporting-range end date used where individual dates are missing')

    if out['Date'].isna().all():

        warnings.append(f'{filename}: no date column or date in filename; period comparisons unavailable')

    if out['SKU'].eq('Unknown').all():

        warnings.append(f'{filename}: no SKU/ASIN column; SKU analysis unavailable')

    for label, key in fields[kind].items():

        if kind == 'FBA Inventory' and col(df, 'fba_ending') and key in ('fba_starting','fba_receipts','fba_shipments','fba_returns') and col(df,key) is None:

            continue

        if col(df, key) is None:

            warnings.append(f'{filename}: {label} column absent (not included in KPI totals)')

    return out.reset_index(drop=True), warnings

# Snapshot balances need last observation per FNSKU, location and disposition.

# Movement fields (receipts and shipments) must instead be summed across days.

FBA_BALANCES = {'Ending Balance', 'Starting Balance', 'Sellable Ending Balance'}

def deduplicate_fba(df):

    keys = ['Date', 'SKU', 'FNSKU', 'Zone', 'Disposition']

    # Repeated upload of identical CSV files should not double stock or movements.

    return (df.sort_values('Source')

              .drop_duplicates(keys, keep='last')

              .copy())

def latest_inventory(df, period=None):

    dated = deduplicate_fba(df).dropna(subset=['Date']).copy()

    if dated.empty:

        return deduplicate_fba(df)

    keys = ['SKU', 'FNSKU', 'Zone', 'Disposition']

    if period:

        keys.insert(0, period)

    return (dated.sort_values(['Date', 'Source'])

                 .drop_duplicates(keys, keep='last'))

def show_periods(df, metric, kind):

    frame = df.dropna(subset=['Date']).copy()

    summaries = {}

    if frame.empty:

        st.warning('No readable dates. Use a date column or filenames like 20260205.csv.')

        return summaries

    frame['Week'] = (frame['Date'] - pd.to_timedelta(

        (frame['Date'].dt.dayofweek + 1) % 7, unit='D'))

    frame['Month'] = frame['Date'].dt.to_period('M').astype(str)

    for period in ['Week', 'Month']:

        st.header('Weekly Analysis' if period == 'Week' else 'Monthly Analysis')

        f = (latest_inventory(frame, period) if metric in FBA_BALANCES

             else deduplicate_fba(frame)) if kind == 'FBA Inventory' else frame

        summary = f.groupby(period, as_index=False)[metric].sum().sort_values(period)

        summary['Previous'] = summary[metric].shift(1)

        summary['Change %'] = (summary[metric] / summary['Previous'].replace(0, float('nan')) - 1) * 100

        st.subheader(f'{metric} — Period Comparison')

        st.dataframe(summary.iloc[::-1], hide_index=True, use_container_width=True)

        st.line_chart(summary.set_index(period)[metric])

        summaries[period] = summary

        for dimension in ['Category', 'SKU']:

            known = f[f[dimension].ne('Unknown')]

            if known.empty:

                st.info(f'{dimension} comparison unavailable: no {dimension.lower()} column found.')

                continue

            pivot = known.pivot_table(index=dimension, columns=period, values=metric,

                                      aggfunc='sum', fill_value=0)

            st.subheader(f'{dimension}-wise {metric}')

            st.dataframe(pivot, use_container_width=True)

    return summaries

# =====================================

# UNIFIED MULTI-REPORT DASHBOARD

# =====================================

# OneDrive synced-folder integration (local Windows Streamlit server).
# Folder values are read from .env; no folders are created or modified.
# Existing *_FOLDER_URL environment names are retained for compatibility.
ONEDRIVE_FOLDERS = {
    'Orders': os.getenv('ONEDRIVE_ORDER_FOLDER_URL', '').strip(),
    'Ads': os.getenv('ONEDRIVE_ADS_FOLDER_URL', '').strip(),
    'FBA Inventory': os.getenv('ONEDRIVE_FBA_FOLDER_URL', '').strip(),
    'Customer Returns': os.getenv('ONEDRIVE_RETURNS_FOLDER_URL', '').strip(),
    'Traffic': os.getenv('ONEDRIVE_TRAFFIC_FOLDER_URL', '').strip(),
}

SUPPORTED_REPORT_EXTENSIONS = {'.csv', '.txt', '.xlsx', '.xls'}
DIRECT_ONEDRIVE_TYPES = ('Orders', 'Ads', 'FBA Inventory', 'Customer Returns', 'Traffic')


def local_onedrive_folder(report_type):
    configured = ONEDRIVE_FOLDERS.get(report_type, '')
    if not configured:
        return None, f'OneDrive folder path missing in .env for {report_type}.'
    # HTTP links cannot be read as local directories. A synced filesystem path is required.
    if configured.lower().startswith(('https://', 'http://')):
        return None, f'{report_type}: use the local synced OneDrive directory path, not a web URL.'
    folder = Path(os.path.expandvars(os.path.expanduser(configured)))
    if not folder.is_dir():
        return None, f'{report_type}: OneDrive directory not found: {folder}'
    return folder, None


def load_report_blobs(report_type, blobs):
    # Reuse the application's original reading, preparation, caching and error handling.
    signature = [(name, hashlib.sha256(blob).hexdigest()) for name, blob in blobs]
    saved = st.session_state.unified_reports.get(report_type, {})
    if saved.get('signature') == signature:
        return
    frames, errors, warnings = [], [], []
    for filename, blob in blobs:
        try:
            raw = read_file(blob, filename)
            prepared, notes = prepare(raw, report_type, filename)
            frames.append(prepared)
            warnings.extend(notes)
        except Exception as exc:
            errors.append(f'{filename}: {exc}')
    st.session_state.unified_reports[report_type] = {
        'signature': signature, 'frames': frames,
        'errors': errors, 'warnings': warnings,
    }


st.sidebar.subheader('Upload Reports for Analysis')
st.sidebar.caption('Select reports from your locally synced OneDrive folders.')

if 'unified_reports' not in st.session_state:
    st.session_state.unified_reports = {}

for report_type in DIRECT_ONEDRIVE_TYPES:
    folder, folder_error = local_onedrive_folder(report_type)
    if folder_error:
        st.sidebar.warning(folder_error)
        continue

    # All five report types, including Traffic, use the same local OneDrive dropdown.
    # Listing is read-only; no duplicate folders or report files are created.
    try:
        candidates = sorted(
            (item for item in folder.iterdir()
             if item.is_file() and item.suffix.lower() in SUPPORTED_REPORT_EXTENSIONS
             and not item.name.startswith('~$')),
            key=lambda item: item.name.lower(),
        )
    except OSError as exc:
        st.sidebar.error(f'{report_type}: unable to list folder: {exc}')
        continue

    file_by_name = {item.name: item for item in candidates}
    st.sidebar.caption(f'{report_type}: {len(candidates)} supported file(s) in OneDrive')
    chosen_names = st.sidebar.multiselect(
        f'Select {report_type} files',
        options=list(file_by_name),
        default=[],
        key=f'onedrive_select_{report_type}',
        placeholder='Choose files (leave empty for all)',
    )
    if st.sidebar.button(
        f'Upload {report_type} Files',
        key=f'onedrive_upload_{report_type}',
        use_container_width=True,
    ):
        selected = [file_by_name[name] for name in (chosen_names or list(file_by_name))]
        if not selected:
            st.sidebar.warning(f'No supported {report_type} files found in the configured folder.')
        else:
            blobs = []
            read_errors = []
            for item in selected:
                try:
                    blobs.append((item.name, item.read_bytes()))
                except OSError as exc:
                    read_errors.append(f'{item.name}: {exc}')
            load_report_blobs(report_type, blobs)
            if read_errors:
                st.session_state.unified_reports.setdefault(report_type, {
                    'signature': [], 'frames': [], 'errors': [], 'warnings': []
                })['errors'].extend(read_errors)
            st.sidebar.success(f'{len(blobs)} {report_type} file(s) read from OneDrive.')

if st.sidebar.button('Clear all report data'):

    st.session_state.unified_reports = {}

    for report_type in TYPES[:-1]:

        st.session_state.pop('unified_' + report_type, None)
        st.session_state.pop('onedrive_select_' + report_type, None)

    st.rerun()

reports = {}

for report_type, state in st.session_state.unified_reports.items():

    for error in state.get('errors', []):

        st.error(f'{report_type}: {error}')

    with st.expander(f'{report_type} upload notes', expanded=False):

        for warning in state.get('warnings', []):

            st.warning(warning)

        st.write(f"Valid files: {len(state.get('frames', []))}")

    if state.get('frames'):

        data = pd.concat(state['frames'], ignore_index=True)

        if report_type == 'FBA Inventory':

            data = deduplicate_fba(data)

        elif report_type == 'Orders' and 'Order Item ID' in data:

            with_id = data['Order Item ID'].astype('string').fillna('').str.strip().ne('')

            keys = ['Order Item ID'] + (['Order ID'] if 'Order ID' in data else [])

            data = pd.concat([

                data.loc[with_id].drop_duplicates(subset=keys),

                data.loc[~with_id].drop_duplicates()

            ], ignore_index=True)

        reports[report_type] = data

if not reports:

    st.info('Upload Orders, Ads, Traffic, FBA Inventory or Returns from the sidebar.')

    st.stop()

st.header('Amazon Unified Dashboard')

st.caption('Sources loaded: ' + ', '.join(reports))

# =====================================
# GLOBAL DATE SLICER (ALL FIVE REPORTS)
# =====================================
# Every loaded report is filtered with the same selected range.
# All Dates uses the entire range of dates available in uploaded reports.
# The preset dates are based on India time, not the computer's UTC clock.

valid_dates = pd.concat(
    [pd.to_datetime(d['Date'], errors='coerce').dropna()
     for d in reports.values()],
    ignore_index=True,
)

start = end = None

if not valid_dates.empty:
    minimum_date = valid_dates.min().date()
    maximum_date = valid_dates.max().date()
    today_india = pd.Timestamp.now(tz='Asia/Kolkata').date()

    st.subheader('Dashboard Date Slicer')
    preset = st.selectbox(
        'Select reporting period',
        [
            'All Dates',
            'Today',
            'Yesterday',
            'Last 7 Days',
            'Last 30 Days',
            'This Month',
            'Last Month',
            'Custom Date Range',
        ],
        key='global_date_preset',
    )

    if preset == 'Today':
        start = end = today_india
    elif preset == 'Yesterday':
        start = end = today_india - pd.Timedelta(days=1)
    elif preset == 'Last 7 Days':
        start = today_india - pd.Timedelta(days=6)
        end = today_india
    elif preset == 'Last 30 Days':
        start = today_india - pd.Timedelta(days=29)
        end = today_india
    elif preset == 'This Month':
        start = today_india.replace(day=1)
        end = today_india
    elif preset == 'Last Month':
        this_month_start = today_india.replace(day=1)
        end = this_month_start - pd.Timedelta(days=1)
        start = end.replace(day=1)
    elif preset == 'Custom Date Range':
        chosen_dates = st.date_input(
            'Choose start and end dates',
            value=(max(minimum_date, maximum_date - pd.Timedelta(days=2)), maximum_date),
            min_value=minimum_date,
            max_value=maximum_date,
            format='DD/MM/YYYY',
            key='global_custom_date_range',
        )
        if not isinstance(chosen_dates, (tuple, list)) or len(chosen_dates) != 2:
            st.info('Select both start and end dates to apply your filter.')
            st.stop()
        start, end = chosen_dates
    else:
        start, end = minimum_date, maximum_date

    st.caption(
        f'Selected period: {start:%d %b %Y} – {end:%d %b %Y} | '
        f'Available data: {minimum_date:%d %b %Y} – {maximum_date:%d %b %Y}'
    )
    if end < minimum_date or start > maximum_date:
        st.warning('The selected period is outside the dates of the uploaded reports.')
else:
    st.warning('No valid report dates found. Date filtering is unavailable; all rows are shown.')

filtered = {}
for kind, df in reports.items():
    if start is not None and end is not None:
        report_date = pd.to_datetime(df['Date'], errors='coerce').dt.normalize()
        date_mask = report_date.between(pd.Timestamp(start), pd.Timestamp(end))
        filtered[kind] = df.loc[date_mask].copy()
    else:
        filtered[kind] = df.copy()

if valid_dates.size:
    with st.expander('Rows in selected period', expanded=False):
        st.dataframe(
            pd.DataFrame([
                {
                    'Report': kind,
                    'Loaded Rows': len(df),
                    'Selected Rows': len(filtered[kind]),
                }
                for kind, df in reports.items()
            ]),
            hide_index=True,
            use_container_width=True,
        )

orders = filtered.get('Orders', pd.DataFrame())

ads = filtered.get('Ads', pd.DataFrame())

traffic = filtered.get('Traffic', pd.DataFrame())

returns = filtered.get('Customer Returns', pd.DataFrame())

fba = filtered.get('FBA Inventory', pd.DataFrame())

# Sales/Ads/Traffic/Returns are date-window events. FBA inventory is a snapshot:
# use all FBA history up to the chosen end date for ending stock, even if there
# was no new inventory report within the date window.
full_fba = reports.get('FBA Inventory', pd.DataFrame())
if not full_fba.empty and end is not None:
    full_fba_dates = pd.to_datetime(full_fba['Date'], errors='coerce')
    fba_asof_end = full_fba.loc[
        full_fba_dates.notna() & full_fba_dates.le(pd.Timestamp(end))
    ].copy()
else:
    fba_asof_end = full_fba.copy()
stock = latest_inventory(fba_asof_end) if not fba_asof_end.empty else pd.DataFrame()

def total(df, field):

    return float(df[field].sum()) if not df.empty and field in df else 0.0

def ratio(numerator, denominator):

    return f'{numerator / denominator * 100:.2f}%' if denominator else 'N/A'

revenue = total(orders, 'Revenue')

units = total(orders, 'Units')

spend = total(ads, 'Spend')

ad_sales = total(ads, 'Ad Sales')

sessions = total(traffic, 'Sessions')

returned_units = total(returns, 'Returned Units')

# ==========================================
# DASHBOARD KPI CALCULATIONS
# ==========================================

def indian_format(value, decimals=0):
    """Display counts and INR values with Indian lakh/crore comma grouping."""
    if value is None or pd.isna(value):
        return 'N/A'
    number = float(value)
    negative = number < 0
    whole, _, fraction = f'{abs(number):.{decimals}f}'.partition('.')
    if len(whole) > 3:
        left, last_three = whole[:-3], whole[-3:]
        groups = []
        while left:
            groups.insert(0, left[-2:])
            left = left[:-2]
        whole = ','.join(groups + [last_three])
    return ('-' if negative else '') + whole + ('.' + fraction if decimals else '')

def inr(value, decimals=0):
    formatted = indian_format(value, decimals)
    return 'N/A' if formatted == 'N/A' else '₹' + formatted

clicks = total(ads, 'Clicks')
impressions = total(ads, 'Impressions')
cpc = spend / clicks if clicks > 0 else None
cpm = spend * 1000 / impressions if impressions > 0 else None
roas = ad_sales / spend if spend > 0 else None

# Stock balances are snapshots. The *end-date* snapshot defines current stock;
# the last snapshot *before the start date* defines previous stock. If the
# selected window begins with the first inventory observation, fall back to
# the last observation before the current snapshot, if present.
current_inventory = None
previous_inventory = None
inventory_change = None
inventory_change_pct = None
inbound_stock = None
sellable_inventory = None
inventory_snapshot_date = None
previous_snapshot_date = None

if not stock.empty:
    current_inventory = total(stock, 'Ending Balance')
    sellable_inventory = total(stock, 'Sellable Ending Balance')
    if 'Inbound Quantity' in stock and stock['Inbound Quantity'].notna().any():
        inbound_stock = float(stock['Inbound Quantity'].sum(min_count=1))
    inventory_snapshot_date = pd.to_datetime(stock['Date'], errors='coerce').max()

if not full_fba.empty and current_inventory is not None:
    observed = pd.to_datetime(full_fba['Date'], errors='coerce')
    if start is not None:
        earlier = full_fba.loc[observed.notna() & observed.lt(pd.Timestamp(start))].copy()
    else:
        earlier = pd.DataFrame()
    # Prior observation before start is preferable for window change.
    if earlier.empty and pd.notna(inventory_snapshot_date):
        earlier = full_fba.loc[
            observed.notna() & observed.lt(inventory_snapshot_date)
        ].copy()
    if not earlier.empty:
        previous_snapshot = latest_inventory(earlier)
        if not previous_snapshot.empty:
            previous_inventory = total(previous_snapshot, 'Ending Balance')
            previous_snapshot_date = pd.to_datetime(
                previous_snapshot['Date'], errors='coerce'
            ).max()
            inventory_change = current_inventory - previous_inventory
            if previous_inventory:
                inventory_change_pct = inventory_change / previous_inventory * 100

# Total FBA stock is the current ending balance. Inbound is shown separately:
# movement-ledger ending stock is not guaranteed to be the same as sellable stock.
total_inventory = current_inventory

# FBA RECEIVED STOCK: sum receipts for the selected reporting period.
# These are units already received by Amazon, NOT pending inbound shipments.
# The source data must explicitly include a Receipts field.
received_stock = None
if not fba.empty and 'Receipts' in fba.columns:
    if 'Receipts Reported' in fba.columns:
        receipt_rows = fba.loc[fba['Receipts Reported'].fillna(False).astype(bool)]
    else:
        # Compatibility with previously prepared reports.
        receipt_rows = fba
    if not receipt_rows.empty:
        receipt_values = pd.to_numeric(receipt_rows['Receipts'], errors='coerce')
        if receipt_values.notna().any():
            received_stock = float(receipt_values.sum(min_count=1))

metrics = [
    ('Total Sales', inr(revenue) if not orders.empty else 'N/A'),
    ('Total Units', indian_format(units) if not orders.empty else 'N/A'),
    ('Ad Spend', inr(spend) if not ads.empty else 'N/A'),
    ('Ad Sales', inr(ad_sales) if not ads.empty else 'N/A'),
    ('ACOS', ratio(spend, ad_sales) if not ads.empty else 'N/A'),
    ('TACOS', ratio(spend, revenue) if not ads.empty and not orders.empty else 'N/A'),
    ('ROAS', f'{roas:.2f}x' if roas is not None else 'N/A'),
    ('CPC', inr(cpc, 2) if cpc is not None else 'N/A'),
    ('CPM', inr(cpm, 2) if cpm is not None else 'N/A'),
    ('Total Impressions', indian_format(impressions) if not ads.empty else 'N/A'),
    ('Total Clicks', indian_format(clicks) if not ads.empty else 'N/A'),
    ('Sessions', indian_format(sessions) if not traffic.empty else 'N/A'),
    ('Traffic Unit Session %', ratio(total(traffic, 'Units Ordered'), sessions)
     if not traffic.empty else 'N/A'),
    ('Total Inventory', indian_format(total_inventory)),
    ('Current Inventory', indian_format(current_inventory)),
    ('Previous Inventory', indian_format(previous_inventory)),
    ('Sellable Inventory', indian_format(sellable_inventory)),
    ('Inbound Stock', indian_format(inbound_stock)),
    ('FBA Received Stock', indian_format(received_stock)),
    ('Inventory Change', indian_format(inventory_change)),
    ('Inventory Change %', f'{inventory_change_pct:+.2f}%'
     if inventory_change_pct is not None else 'N/A'),
    ('Returned Units', indian_format(returned_units) if not returns.empty else 'N/A'),
    ('Refund Value', inr(total(returns, 'Refund Value')) if not returns.empty else 'N/A'),
]

for offset in range(0, len(metrics), 4):
    cols = st.columns(4)
    for column, (label, value) in zip(cols, metrics[offset:offset + 4]):
        column.metric(label, value)

if not fba.empty and inbound_stock is None:
    st.info('Inbound Stock is N/A: the latest available FBA snapshot has no inbound quantity or inbound status columns. Received Stock (receipts) is a separate measure.')
if not fba.empty and received_stock is None:
    st.info('FBA Received Stock is N/A: no valid Receipts column was found in the selected FBA movement reports.')

st.caption('Order revenue, attributed ad sales and traffic ordered sales are separate measures. '
           'FBA stock is taken from the latest available observations on or before the selected end date; '
           'receipts are summed only within the selected date range.')
if inventory_snapshot_date is not None and pd.notna(inventory_snapshot_date):
    st.caption(f'Current FBA snapshot: {inventory_snapshot_date:%d %b %Y}' +
               (f' | Previous FBA snapshot: {previous_snapshot_date:%d %b %Y}'
                if previous_snapshot_date is not None and pd.notna(previous_snapshot_date)
                else ' | Previous FBA snapshot unavailable'))

# =====================================

# SALES PERFORMANCE: ORDERS ONLY

# =====================================

st.divider()

st.header('Sales Performance')

category_performance = pd.DataFrame()

top20_skus = pd.DataFrame()

weekly_sales = pd.DataFrame()

weekly_category = pd.DataFrame()

weekly_category_units = pd.DataFrame()

week_comparison = pd.DataFrame()

category_week_comparison = pd.DataFrame()

monthly_sales = pd.DataFrame()

monthly_category = pd.DataFrame()

month_comparison = pd.DataFrame()

category_month_comparison = pd.DataFrame()

def percent_change(new_value, old_value):

    if pd.isna(old_value) or old_value == 0:

        return float('nan')

    return (new_value - old_value) / old_value * 100

def period_category_comparison(data, period_col, first, second, title_first, title_second):

    # Missing categories in a selected period are true zeroes only within the loaded period.

    selected = data.loc[data[period_col].isin([first, second])].copy()

    grouped = (selected.groupby(['Category', period_col], as_index=False)[['Revenue', 'Units']].sum())

    names = sorted(grouped['Category'].dropna().unique())

    result = pd.DataFrame({'Category': names})

    for label, period_value in [(title_first, first), (title_second, second)]:

        part = grouped.loc[grouped[period_col].eq(period_value), ['Category', 'Revenue', 'Units']]

        result = result.merge(part.rename(columns={'Revenue': f'{label} Revenue', 'Units': f'{label} Units'}),

                              on='Category', how='left')

    for c in result.columns:

        if c != 'Category':

            result[c] = result[c].fillna(0)

    result['Revenue Difference'] = result[f'{title_second} Revenue'] - result[f'{title_first} Revenue']

    result['Revenue Change %'] = result.apply(

        lambda r: percent_change(r[f'{title_second} Revenue'], r[f'{title_first} Revenue']), axis=1)

    result['Units Difference'] = result[f'{title_second} Units'] - result[f'{title_first} Units']

    result['Units Change %'] = result.apply(

        lambda r: percent_change(r[f'{title_second} Units'], r[f'{title_first} Units']), axis=1)

    return result.sort_values('Revenue Difference', ascending=False)

if orders.empty:

    st.info('Upload an Orders report to view Category Performance, Top 20 SKUs, and Weekly Sales Performance.')

else:

    # Use order revenue only, not advertising-attributed or traffic sales.

    st.subheader('Category Performance')

    category_rows = orders.loc[orders['Category'].notna() & orders['Category'].ne('Unknown')].copy()

    if category_rows.empty:

        st.info('No category information found in the uploaded Orders report.')

    else:

        category_performance = (category_rows.groupby('Category', as_index=False)

            .agg(Revenue=('Revenue', 'sum'), Units=('Units', 'sum'))

            .sort_values('Revenue', ascending=False))

        category_performance['ASP'] = (category_performance['Revenue'] /

            category_performance['Units'].replace(0, float('nan')))

        category_performance['Revenue Share %'] = (category_performance['Revenue'] /

            revenue * 100) if revenue else float('nan')

        st.dataframe(category_performance, hide_index=True, use_container_width=True,

            column_config={

                'Revenue': st.column_config.NumberColumn(format='₹%.2f'),

                'ASP': st.column_config.NumberColumn(format='₹%.2f'),

                'Revenue Share %': st.column_config.NumberColumn(format='%.2f%%'),

            })

        st.bar_chart(category_performance.set_index('Category')['Revenue'])

    st.subheader('Top 20 SKUs')

    sku_rows = orders.loc[orders['SKU'].notna() & orders['SKU'].ne('Unknown')].copy()

    if sku_rows.empty:

        st.info('No SKU identifiers found in the Orders report.')

    else:

        top20_skus = (
            sku_rows.groupby("SKU", as_index=False)
            .agg(
                Revenue=("Revenue", "sum"),
                Units=("Units", "sum")
            )
            .sort_values(
                ["Revenue", "Units", "SKU"],
                ascending=[False, False, True]
            )
            .head(20)
            .reset_index(drop=True)
        )

        top20_skus.insert(
            0, "Revenue Rank",
            range(1, len(top20_skus) + 1)
        )

        top20_skus["ASP"] = (
            top20_skus["Revenue"]
            / top20_skus["Units"].replace(0, float("nan"))
        )

        st.dataframe(top20_skus, hide_index=True, use_container_width=True,

            column_config={

                'Revenue': st.column_config.NumberColumn(format='₹%.2f'),

                'ASP': st.column_config.NumberColumn(format='₹%.2f'),

            })

        st.bar_chart(top20_skus.set_index('SKU')['Revenue'])

        st.caption('Top 20 SKUs ranked by revenue; ties are ordered by units sold and then SKU.')

    st.header('Weekly Sales Performance')

    dated_orders = orders.dropna(subset=['Date']).copy()

    if dated_orders.empty:

        st.warning('Weekly analysis requires valid individual order dates.')

    else:

        # Sunday–Saturday weeks; Week Start is the Sunday of each reporting week.

        dated_orders['Week Start'] = (dated_orders['Date'].dt.normalize() -

            pd.to_timedelta((dated_orders['Date'].dt.dayofweek + 1) % 7, unit='D'))

        weekly_sales = (dated_orders.groupby('Week Start', as_index=False)

            .agg(Revenue=('Revenue', 'sum'), Units=('Units', 'sum'))

            .sort_values('Week Start'))

        weekly_sales['Week End'] = weekly_sales['Week Start'] + pd.Timedelta(days=6)

        # Missing intermediate weeks must not be mistaken for adjacent weeks.

        weekly_sales['Previous Week Revenue'] = weekly_sales['Revenue'].shift(1)

        weekly_sales['Previous Week Units'] = weekly_sales['Units'].shift(1)

        consecutive = weekly_sales['Week Start'].diff().eq(pd.Timedelta(days=7))

        weekly_sales.loc[~consecutive, ['Previous Week Revenue', 'Previous Week Units']] = float('nan')

        weekly_sales['Revenue WoW %'] = (

            (weekly_sales['Revenue'] / weekly_sales['Previous Week Revenue'].replace(0, float('nan'))) - 1

        ) * 100

        weekly_sales['Units WoW %'] = (

            (weekly_sales['Units'] / weekly_sales['Previous Week Units'].replace(0, float('nan'))) - 1

        ) * 100

        weekly_sales['ASP'] = weekly_sales['Revenue'] / weekly_sales['Units'].replace(0, float('nan'))

        weekly_sales['Week'] = (weekly_sales['Week Start'].dt.strftime('%d %b %Y') +

            ' – ' + weekly_sales['Week End'].dt.strftime('%d %b %Y'))

        st.subheader('Week-over-Week Comparison')

        compare_cols = ['Week', 'Revenue', 'Previous Week Revenue', 'Revenue WoW %',

                        'Units', 'Previous Week Units', 'Units WoW %', 'ASP']

        st.dataframe(weekly_sales[compare_cols].iloc[::-1], hide_index=True,

            use_container_width=True,

            column_config={

                'Revenue': st.column_config.NumberColumn(format='₹%.2f'),

                'Previous Week Revenue': st.column_config.NumberColumn(format='₹%.2f'),

                'Revenue WoW %': st.column_config.NumberColumn(format='%.2f%%'),

                'Units WoW %': st.column_config.NumberColumn(format='%.2f%%'),

                'ASP': st.column_config.NumberColumn(format='₹%.2f'),

            })

        st.caption('Weeks run Sunday–Saturday. WoW is shown only when the previous consecutive week '

                   'is present and its value is nonzero. Weeks partly covered by uploaded reports '

                   'or the selected date filter may not be directly comparable.')

        st.subheader('Weekly Revenue Trend')

        st.line_chart(weekly_sales.set_index('Week Start')['Revenue'], y_label='Revenue (₹)')

        st.subheader('Weekly Units Sold')

        st.bar_chart(weekly_sales.set_index('Week Start')['Units'], y_label='Units')

        st.download_button('Download weekly sales comparison (CSV)',

            data=weekly_sales[compare_cols].to_csv(index=False).encode('utf-8-sig'),

            file_name='weekly_sales_performance.csv', mime='text/csv')

        st.subheader('Category-wise Weekly Comparison')

        category_dated = dated_orders[dated_orders['Category'].notna() &

                                      dated_orders['Category'].ne('Unknown')].copy()

        if category_dated.empty:

            st.info('No recognized category values in the Orders report.')

        else:

            weekly_category = (category_dated.groupby(['Week Start', 'Category'], as_index=False)

                               .agg(Revenue=('Revenue', 'sum'), Units=('Units', 'sum')))

            weekly_category['Week End'] = weekly_category['Week Start'] + pd.Timedelta(days=6)

            weekly_category['Week'] = (weekly_category['Week Start'].dt.strftime('%d %b %Y') +

                                       ' – ' + weekly_category['Week End'].dt.strftime('%d %b %Y'))

            weekly_category['ASP'] = weekly_category['Revenue'] / weekly_category['Units'].replace(0, float('nan'))

            st.dataframe(weekly_category[['Week', 'Category', 'Revenue', 'Units', 'ASP']]

                         .sort_values(['Week', 'Revenue'], ascending=[False, False]),

                         use_container_width=True, hide_index=True)

            weekly_category_revenue = weekly_category.pivot_table(

                index='Week Start', columns='Category', values='Revenue', aggfunc='sum', fill_value=0)

            st.bar_chart(weekly_category_revenue, stack=True, y_label='Revenue (₹)')

            st.subheader('Weekly Category Units Sold')

            weekly_category_units = weekly_category.pivot_table(

                index='Week Start', columns='Category', values='Units', aggfunc='sum', fill_value=0)

            st.dataframe(weekly_category_units.reset_index(), use_container_width=True, hide_index=True)

            st.bar_chart(weekly_category_units, stack=True, y_label='Units')

        st.subheader('Compare Two Weeks')

        week_options = list(weekly_sales['Week Start'].sort_values())

        week_labels = {w: f"{w:%d %b %Y} – {(w + pd.Timedelta(days=6)):%d %b %Y}"

                       for w in week_options}

        if len(week_options) < 2:

            st.info('Upload data from at least two distinct weeks to compare.')

        else:

            wa, wb = st.columns(2)

            w_a = wa.selectbox('First week', week_options, index=len(week_options)-2,

                               format_func=lambda x: week_labels[x], key='week_first')

            w_b = wb.selectbox('Second week', week_options, index=len(week_options)-1,

                               format_func=lambda x: week_labels[x], key='week_second')

            if w_a == w_b:

                st.warning('Select two different weeks.')

            else:

                first_week = weekly_sales.loc[weekly_sales['Week Start'].eq(w_a)].iloc[0]

                second_week = weekly_sales.loc[weekly_sales['Week Start'].eq(w_b)].iloc[0]

                week_comparison = pd.DataFrame([

                    {'Metric': m, week_labels[w_a]: first_week[m], week_labels[w_b]: second_week[m],

                     'Difference': second_week[m] - first_week[m],

                     'Change %': percent_change(second_week[m], first_week[m])}

                    for m in ['Revenue', 'Units']])

                st.dataframe(week_comparison, use_container_width=True, hide_index=True)

                st.caption('Comparisons reflect uploaded data and date filters; partial weeks may be unequal.')

                if not weekly_category.empty:

                    st.subheader('Category Revenue Comparison')

                    category_week_comparison = period_category_comparison(

                        weekly_category, 'Week Start', w_a, w_b,

                        week_labels[w_a], week_labels[w_b])

                    st.dataframe(category_week_comparison, use_container_width=True, hide_index=True)

                    chart_cols = [f'{week_labels[w_a]} Revenue', f'{week_labels[w_b]} Revenue']

                    st.bar_chart(category_week_comparison.set_index('Category')[chart_cols])

    st.header('Monthly Sales Performance')

    if dated_orders.empty:

        st.info('Monthly analysis requires valid dates.')

    else:

        dated_orders['Month Start'] = dated_orders['Date'].dt.to_period('M').dt.to_timestamp()

        monthly_sales = (dated_orders.groupby('Month Start', as_index=False)

                         .agg(Revenue=('Revenue', 'sum'), Units=('Units', 'sum'))

                         .sort_values('Month Start'))

        monthly_sales['Month'] = monthly_sales['Month Start'].dt.strftime('%B %Y')

        monthly_sales['ASP'] = monthly_sales['Revenue'] / monthly_sales['Units'].replace(0, float('nan'))

        monthly_sales['Previous Month Revenue'] = monthly_sales['Revenue'].shift(1)

        monthly_sales['Previous Month Units'] = monthly_sales['Units'].shift(1)

        consecutive_month = ((monthly_sales['Month Start'].dt.year * 12 + monthly_sales['Month Start'].dt.month)

                             .diff().eq(1))

        monthly_sales.loc[~consecutive_month, ['Previous Month Revenue', 'Previous Month Units']] = float('nan')

        monthly_sales['Revenue MoM %'] = (monthly_sales['Revenue'] /

            monthly_sales['Previous Month Revenue'].replace(0, float('nan')) - 1) * 100

        monthly_sales['Units MoM %'] = (monthly_sales['Units'] /

            monthly_sales['Previous Month Units'].replace(0, float('nan')) - 1) * 100

        st.subheader('Month-over-Month Comparison')

        st.dataframe(monthly_sales[['Month', 'Revenue', 'Previous Month Revenue', 'Revenue MoM %',

                                    'Units', 'Previous Month Units', 'Units MoM %', 'ASP']].iloc[::-1],

                     hide_index=True, use_container_width=True)

        st.caption('MoM is shown only for consecutive calendar months with a nonzero prior value.')

        st.subheader('Monthly Revenue Trend')

        st.line_chart(monthly_sales.set_index('Month Start')['Revenue'], y_label='Revenue (₹)')

        st.subheader('Monthly Units Sold')

        st.bar_chart(monthly_sales.set_index('Month Start')['Units'], y_label='Units')

        st.subheader('Category-wise Monthly Performance')

        category_month_rows = dated_orders.loc[dated_orders['Category'].notna() &

                                               dated_orders['Category'].ne('Unknown')].copy()

        if category_month_rows.empty:

            st.info('No category identifiers available.')

        else:

            monthly_category = (category_month_rows.groupby(['Month Start', 'Category'], as_index=False)

                                .agg(Revenue=('Revenue', 'sum'), Units=('Units', 'sum')))

            monthly_category['Month'] = monthly_category['Month Start'].dt.strftime('%B %Y')

            monthly_category['ASP'] = (monthly_category['Revenue'] /

                                       monthly_category['Units'].replace(0, float('nan')))

            st.dataframe(monthly_category[['Month', 'Category', 'Revenue', 'Units', 'ASP']]

                         .sort_values(['Month', 'Revenue'], ascending=[False, False]),

                         use_container_width=True, hide_index=True)

            st.bar_chart(monthly_category.pivot_table(

                index='Month Start', columns='Category', values='Revenue', fill_value=0),

                stack=True, y_label='Revenue (₹)')

        st.subheader('Compare Two Months')

        month_options = list(monthly_sales['Month Start'].sort_values())

        month_labels = {m: m.strftime('%B %Y') for m in month_options}

        if len(month_options) < 2:

            st.info('Upload data from at least two different calendar months to compare.')

        else:

            ma, mb = st.columns(2)

            m_a = ma.selectbox('First month', month_options, index=len(month_options)-2,

                               format_func=lambda x: month_labels[x], key='month_first')

            m_b = mb.selectbox('Second month', month_options, index=len(month_options)-1,

                               format_func=lambda x: month_labels[x], key='month_second')

            if m_a == m_b:

                st.warning('Select two different months.')

            else:

                first_month = monthly_sales.loc[monthly_sales['Month Start'].eq(m_a)].iloc[0]

                second_month = monthly_sales.loc[monthly_sales['Month Start'].eq(m_b)].iloc[0]

                month_comparison = pd.DataFrame([

                    {'Metric': m, month_labels[m_a]: first_month[m], month_labels[m_b]: second_month[m],

                     'Difference': second_month[m] - first_month[m],

                     'Change %': percent_change(second_month[m], first_month[m])}

                    for m in ['Revenue', 'Units']])

                st.dataframe(month_comparison, use_container_width=True, hide_index=True)

                if not monthly_category.empty:

                    category_month_comparison = period_category_comparison(

                        monthly_category, 'Month Start', m_a, m_b,

                        month_labels[m_a], month_labels[m_b])

                    st.subheader('Monthly Category Revenue Comparison')

                    st.dataframe(category_month_comparison, use_container_width=True, hide_index=True)

                    st.bar_chart(category_month_comparison.set_index('Category')[

                        [f'{month_labels[m_a]} Revenue', f'{month_labels[m_b]} Revenue']])

st.divider()

st.subheader('Daily, Weekly and Monthly Performance')

frequency = st.radio('Group by', ['Daily', 'Weekly', 'Monthly'], horizontal=True)

freq = {'Daily': 'D', 'Weekly': 'W-SAT', 'Monthly': 'MS'}[frequency]

series_parts = []

trends = pd.DataFrame()

for kind, columns in [

    ('Orders', ['Revenue', 'Units']),

    ('Ads', ['Spend', 'Ad Sales', 'Clicks', 'Impressions']),

    ('Traffic', ['Sessions', 'Page Views', 'Units Ordered', 'Ordered Sales']),

    ('Customer Returns', ['Returned Units', 'Refund Value'])]:

    frame = filtered.get(kind, pd.DataFrame())

    if frame.empty:

        continue

    frame = frame.dropna(subset=['Date'])

    if frame.empty:

        continue

    selected = [x for x in columns if x in frame]

    grouped = frame.groupby(pd.Grouper(key='Date', freq=freq))[selected].sum()

    grouped = grouped.rename(columns={c: f'{kind} | {c}' for c in selected})

    series_parts.append(grouped)

if series_parts:

    trends = pd.concat(series_parts, axis=1).sort_index()

    st.dataframe(trends.reset_index(), use_container_width=True, hide_index=True)

    plot_columns = [x for x in ['Orders | Revenue', 'Ads | Spend', 'Ads | Ad Sales'] if x in trends]

    if plot_columns:

        st.line_chart(trends[plot_columns])

else:

    st.info('No usable dates for period comparisons.')

st.subheader('Combined SKU and Category Performance')

level = st.radio('Analyze by', ['SKU', 'Category'], horizontal=True)

parts = []

for kind, cols in [

    ('Orders', ['Revenue', 'Units']),

    ('Ads', ['Spend', 'Ad Sales', 'Clicks', 'Impressions']),

    ('Traffic', ['Sessions', 'Page Views', 'Units Ordered', 'Ordered Sales']),

    ('Customer Returns', ['Returned Units', 'Refund Value']),

    ('FBA Inventory', ['Ending Balance', 'Sellable Ending Balance'])]:

    frame = stock if kind == 'FBA Inventory' else filtered.get(kind, pd.DataFrame())

    if frame.empty or level not in frame:

        continue

    frame = frame[frame[level].notna() & frame[level].ne('Unknown')]

    if frame.empty:

        continue

    cols = [c for c in cols if c in frame]

    part = frame.groupby(level)[cols].sum().rename(columns={c: f'{kind} | {c}' for c in cols})

    parts.append(part)

if parts:

    combined = pd.concat(parts, axis=1).fillna(0)

    if 'Orders | Revenue' in combined:

        combined['TACOS %'] = (combined.get('Ads | Spend', 0) /

                                combined['Orders | Revenue'].replace(0, float('nan')) * 100)

    if 'Ads | Ad Sales' in combined:

        combined['ACOS %'] = (combined.get('Ads | Spend', 0) /

                               combined['Ads | Ad Sales'].replace(0, float('nan')) * 100)

    if 'Traffic | Sessions' in combined:

        combined['Traffic Unit Session %'] = (

            combined.get('Traffic | Units Ordered', 0) /

            combined['Traffic | Sessions'].replace(0, float('nan')) * 100)

    sort_field = 'Orders | Revenue' if 'Orders | Revenue' in combined else combined.columns[0]

    combined = combined.sort_values(sort_field, ascending=False)

    st.dataframe(combined.reset_index(), use_container_width=True, hide_index=True)

else:

    combined = pd.DataFrame()

    st.info('No common SKU or category values available.')

if not stock.empty:

    st.subheader('FBA Warehouse Breakdown')

    if 'Zone' in stock:

        zone_summary = stock.groupby('Zone', as_index=False)[

            ['Ending Balance', 'Sellable Ending Balance']].sum()

        st.dataframe(zone_summary, use_container_width=True, hide_index=True)

    st.caption('Movement-ledger ending stock is not necessarily the same as available-to-sell inventory.')

if not ads.empty and 'Campaign' in ads:

    st.subheader('Advertising Campaign Performance')

    campaigns = ads[ads['Campaign'].ne('Unknown')].groupby('Campaign', as_index=False)[

        ['Spend', 'Ad Sales', 'Clicks', 'Impressions']].sum()

    if not campaigns.empty:

        campaigns['ACOS %'] = campaigns['Spend'] / campaigns['Ad Sales'].replace(0, float('nan')) * 100

        st.dataframe(campaigns.sort_values('Spend', ascending=False), use_container_width=True, hide_index=True)

# =====================================

# SKU ASP HISTORY AND THRESHOLD ANALYSIS

# =====================================

st.divider()

st.header('SKU Price / ASP Change Analysis')

st.caption('Order-derived ASP = revenue divided by units; it is not verified listing price history.')



sku_weekly_history = pd.DataFrame()

price_threshold = None
min_units = None

sku_monthly_history = pd.DataFrame()

price_crossings = pd.DataFrame()



if not orders.empty:

    dated_sku = orders.dropna(subset=['Date']).copy()

    dated_sku = dated_sku.loc[

        dated_sku['SKU'].notna() & dated_sku['SKU'].ne('Unknown')

        & dated_sku['Units'].gt(0)

    ].copy()

    if not dated_sku.empty:

        dated_sku['Week Start'] = (

            dated_sku['Date'].dt.normalize()

            - pd.to_timedelta((dated_sku['Date'].dt.dayofweek + 1) % 7, unit='D')

        )

        dated_sku['Month Start'] = dated_sku['Date'].dt.to_period('M').dt.to_timestamp()

        sku_weekly_history = (

            dated_sku.groupby(['SKU', 'Week Start'], as_index=False)

            .agg(Revenue=('Revenue', 'sum'), Units=('Units', 'sum'))

            .sort_values(['SKU', 'Week Start'])

        )

        sku_weekly_history['ASP'] = (sku_weekly_history['Revenue'] /

                                     sku_weekly_history['Units'].replace(0, float('nan')))

        sku_weekly_history['Week End'] = sku_weekly_history['Week Start'] + pd.Timedelta(days=6)

        sku_monthly_history = (

            dated_sku.groupby(['SKU', 'Month Start'], as_index=False)

            .agg(Revenue=('Revenue', 'sum'), Units=('Units', 'sum'))

            .sort_values(['SKU', 'Month Start'])

        )

        sku_monthly_history['ASP'] = (sku_monthly_history['Revenue'] /

                                      sku_monthly_history['Units'].replace(0, float('nan')))



        price_threshold = st.number_input(

            'ASP threshold (₹)', min_value=1.0, value=1600.0, step=100.0

        )

        min_units = st.number_input(

            'Minimum units in each comparison week', min_value=1, value=1, step=1

        )

        h = sku_weekly_history.copy()

        h['Previous Week'] = h.groupby('SKU')['Week Start'].shift(1)

        for measure in ['ASP', 'Units', 'Revenue']:

            h['Previous ' + measure] = h.groupby('SKU')[measure].shift(1)

        # Only true adjacent weeks. Never compare across missing reporting weeks.

        eligible = (

            h['Previous Week'].notna()

            & h['Week Start'].sub(h['Previous Week']).eq(pd.Timedelta(days=7))

            & h['Previous ASP'].lt(price_threshold)

            & h['ASP'].gt(price_threshold)

            & h['Previous Units'].ge(min_units)

            & h['Units'].ge(min_units)

        )

        price_crossings = h.loc[eligible, [

            'SKU', 'Previous Week', 'Week Start', 'Previous ASP', 'ASP',

            'Previous Units', 'Units', 'Previous Revenue', 'Revenue'

        ]].copy()

        if not price_crossings.empty:

            price_crossings['ASP Change %'] = (

                (price_crossings['ASP'] / price_crossings['Previous ASP']) - 1

            ) * 100

            price_crossings['Units Change %'] = (

                (price_crossings['Units'] / price_crossings['Previous Units']) - 1

            ) * 100

            price_crossings['Revenue Change %'] = (

                (price_crossings['Revenue'] / price_crossings['Previous Revenue'].replace(0, float('nan'))) - 1

            ) * 100

            price_crossings = price_crossings.sort_values(['Week Start', 'SKU'], ascending=[False, True])

        # Format date columns only for display. Original datetime data remains
        # unchanged for weekly comparisons, Gemini context, and Excel exports.
        def display_dates_without_time(frame):
            display_df = frame.copy()
            for column in ['Date', 'Week Start', 'Week End',
                           'Month Start', 'Previous Week']:
                if column in display_df.columns:
                    display_df[column] = pd.to_datetime(
                        display_df[column], errors='coerce'
                    ).dt.strftime('%Y-%m-%d')
            return display_df

        st.subheader('Weekly SKU ASP history')

        # Show only the SKU histories that crossed the selected threshold
        # between two consecutive weeks, with minimum units in both weeks.
        if not price_crossings.empty:
            matching_skus = price_crossings['SKU'].unique()
            filtered_history = sku_weekly_history[
                sku_weekly_history['SKU'].isin(matching_skus)
            ]
            display_df = display_dates_without_time(filtered_history)
            st.dataframe(
                display_df,
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info(
                f"No SKUs crossed the ₹{price_threshold:,.0f} "
                "ASP threshold between consecutive weeks."
            )

        st.subheader('Below-to-above threshold: consecutive-week comparisons')

        if price_crossings.empty:

            st.info('No matching consecutive-week SKU ASP crossings were found in the loaded period.')

        else:

            st.dataframe(
                display_dates_without_time(price_crossings),
                hide_index=True,
                use_container_width=True
            )

        st.caption('A crossing does not prove that the listing price was increased. Weeks may be partial; promotions, discounts and data coverage can change ASP.')

    else:

        st.info('Valid dated SKU order rows are required for price change analysis.')

else:

    st.info('Upload Orders reports for SKU-level price change analysis.')



# ==========================================

# GEMINI AI — COMBINED SALES INSIGHTS

# ==========================================

st.divider()

st.header('LangChain + Gemini AI — Combined Report Insights')

st.caption('LangChain sends selected aggregated data to Gemini, not raw customer/order identifiers.')

question = st.text_input(

    'Ask LangChain + Gemini about your Amazon sales',

    placeholder='Which SKUs crossed from below ₹1,600 ASP to above ₹1,600, and how did units change?'

)



def extract_requested_sales_date(message):
    """Interpret an explicit day in a user question (India DD-MM-YYYY by default)."""
    patterns = [
        (r'(?<!\d)(\d{1,2})[/-](\d{1,2})[/-](20\d{2})(?!\d)', 'dmy'),
        (r'(?<!\d)(20\d{2})-(\d{1,2})-(\d{1,2})(?!\d)', 'ymd'),
    ]
    for pattern, style in patterns:
        match = re.search(pattern, message)
        if match:
            a, b, c = match.groups()
            try:
                return pd.Timestamp(f'{c}-{b}-{a}' if style == 'dmy' else f'{a}-{b}-{c}').normalize()
            except (ValueError, TypeError):
                pass
    # English-written dates, e.g. 2 November 2026 / November 2, 2026.
    months = ('january|february|march|april|may|june|july|august|'
              'september|october|november|december')
    match = re.search(rf'\b(\d{{1,2}})\s+({months})\s+(20\d{{2}})\b', message, re.I)
    if match:
        try:
            return pd.Timestamp(f'{match[1]} {match[2]} {match[3]}').normalize()
        except ValueError:
            pass
    match = re.search(rf'\b({months})\s+(\d{{1,2}}),?\s+(20\d{{2}})\b', message, re.I)
    if match:
        try:
            return pd.Timestamp(f'{match[1]} {match[2]} {match[3]}').normalize()
        except ValueError:
            pass
    return None


def daily_sku_order_summary(order_frame, requested_day, message):
    """Aggregate actual filtered Orders data, never weekly/monthly approximations."""
    empty_columns = ['SKU', 'Sales (INR)', 'Units', 'ASP (INR)']
    if order_frame.empty:
        return pd.DataFrame(columns=empty_columns), None
    rows = order_frame.loc[
        pd.to_datetime(order_frame['Date'], errors='coerce').dt.normalize().eq(requested_day)
    ].copy()
    available_skus = rows['SKU'].dropna().astype(str).unique().tolist()
    mentioned = [sku for sku in available_skus if sku.lower() in message.lower()]
    if mentioned:
        rows = rows[rows['SKU'].isin(mentioned)]
    if rows.empty:
        return pd.DataFrame(columns=empty_columns), mentioned
    summary = (rows.groupby('SKU', dropna=False, as_index=False)
               .agg(**{'Sales (INR)': ('Revenue', 'sum'), 'Units': ('Units', 'sum')}))
    summary['ASP (INR)'] = summary['Sales (INR)'] / summary['Units'].replace(0, float('nan'))
    summary = summary.sort_values(['Sales (INR)', 'Units'], ascending=False)
    return summary, mentioned


if st.button('Ask LangChain + Gemini'):

    if not question.strip():

        st.warning('Please enter a question.')

    elif not os.getenv('GEMINI_API_KEY'):

        st.error('GEMINI_API_KEY missing from .env')

    else:

        try:
            requested_day = extract_requested_sales_date(question)
            daily_focus = requested_day is not None and any(
                word in question.lower() for word in ('sku', 'sale', 'unit', 'order', 'revenue', 'asp')
            )
            day_summary = pd.DataFrame()
            if daily_focus:
                day_summary, mentioned_skus = daily_sku_order_summary(orders, requested_day, question)
                st.subheader(f'SKU sales for {requested_day:%d-%m-%Y} — Orders only')
                if day_summary.empty:
                    st.warning('No matching order rows for this date in the currently selected '
                               'dashboard reporting period. Check that the day is included '
                               'in the date filter and its Orders files were loaded.')
                else:
                    st.dataframe(day_summary, use_container_width=True, hide_index=True)
                    st.caption(f'{len(day_summary)} SKUs | Sales: ₹{day_summary["Sales (INR)"].sum():,.2f} '
                               f'| Units: {day_summary["Units"].sum():,.0f}. '
                               'Cancelled and pending orders are excluded by the existing cleaning logic.')
                    st.download_button(
                        f'Download {requested_day:%d-%m-%Y} SKU Sales (CSV)',
                        day_summary.to_csv(index=False).encode('utf-8-sig'),
                        file_name=f'sku_sales_{requested_day:%Y%m%d}.csv', mime='text/csv',
                        key='daily_sku_sales_download',
                    )
            # Question-aware sections are kept separate; do not truncate a single

            # huge string at 30,000 characters and accidentally omit SKU history.

            if daily_focus:
                if day_summary.empty:
                    context_sections = [
                        f'NO MATCHING DAILY ORDERS for {requested_day:%d-%m-%Y} in the loaded, filtered data. '
                        'Never estimate daily figures from weekly or monthly data.'
                    ]
                else:
                    context_sections = [
                        f'EXACT DAILY SKU SALES FROM ORDERS for {requested_day:%d-%m-%Y}. '
                        'This table is authoritative for the specified date; do not replace '
                        'it with weekly/monthly totals.\n' + day_summary.to_string(index=False)
                    ]
            else:
                context_sections = [
                    'DASHBOARD KPIs:\n' + '\n'.join(f'{name}: {value}' for name, value in metrics)
                ]

            lower_question = question.lower()

            price_question = any(term in lower_question for term in (

                'price', 'pricing', 'asp', '1600', '1,600', 'increase', 'decrease',

                'before', 'after', 'threshold'

            ))

            # Compact and prioritize report summaries before sending to Gemini.
            # Date filtering has already happened above. This reduces request size,
            # not the number of days in the dashboard.
            max_context_chars = 30000

            def compact_frame(frame, max_rows=30, max_columns=12):
                if frame is None or frame.empty:
                    return ''
                view = frame.copy()
                # Preserve key identifying and performance fields first.
                priority = [
                    'SKU', 'Category', 'Campaign', 'Week', 'Month',
                    'Week Start', 'Month Start', 'Revenue', 'Units', 'ASP',
                    'Spend', 'Ad Sales', 'Clicks', 'Impressions', 'Sessions',
                    'Orders | Revenue', 'Orders | Units', 'Ads | Spend',
                    'Ads | Ad Sales', 'Traffic | Sessions', 'TACOS %',
                    'ACOS %', 'Traffic Unit Session %',
                    'FBA Inventory | Ending Balance',
                    'FBA Inventory | Sellable Ending Balance'
                ]
                ordered = [c for c in priority if c in view.columns]
                ordered.extend(c for c in view.columns if c not in ordered)
                view = view[ordered[:max_columns]]
                for field in view.select_dtypes(include='number').columns:
                    view[field] = view[field].round(2)
                # Biggest sales contributors first in the combined section.
                if 'Orders | Revenue' in view:
                    view = view.sort_values('Orders | Revenue', ascending=False)
                # For other tables, preserve their existing ordering.
                return view.head(max_rows).to_csv(index=False, float_format='%.2f')

            # Keep important context ahead of less relevant historical summaries.
            # Any pre-existing daily context remains the primary source for a daily question.
            if price_question and not daily_focus:
                if not price_crossings.empty:
                    context_sections.append(
                        'COMPUTED CONSECUTIVE-WEEK ASP CROSSINGS:\n' +
                        compact_frame(price_crossings, max_rows=50, max_columns=11)
                    )
                else:
                    context_sections.append('No SKU crossed the selected ASP threshold in consecutive observed weeks.')
                if not sku_weekly_history.empty:
                    relevant = (sku_weekly_history[
                        sku_weekly_history['SKU'].isin(price_crossings['SKU'])
                    ] if not price_crossings.empty else sku_weekly_history)
                    context_sections.append(
                        'WEEKLY REALIZED ASP HISTORY (sample; selected dates only):\n' +
                        compact_frame(relevant, max_rows=70, max_columns=6)
                    )
                if not sku_monthly_history.empty:
                    context_sections.append(
                        'MONTHLY REALIZED ASP HISTORY (sample; selected dates only):\n' +
                        compact_frame(sku_monthly_history, max_rows=40, max_columns=5)
                    )

            if not daily_focus:
                # Combined section first so it is not squeezed out by other summaries.
                if not combined.empty:
                    combined_context = combined.reset_index()
                    # Match an explicitly named SKU/category before ranking results.
                    match_cols = [c for c in ('SKU', 'Category') if c in combined_context.columns]
                    matches = pd.Series(False, index=combined_context.index)
                    for field in match_cols:
                        matches |= combined_context[field].astype(str).map(
                            lambda name: name.lower() in lower_question if len(name) >= 5 else False
                        )
                    if matches.any():
                        combined_context = combined_context.loc[matches]
                    context_sections.append(
                        'COMBINED PERFORMANCE (top rows, selected reporting period; '
                        'not an exhaustive list of SKUs):\n' +
                        compact_frame(combined_context, max_rows=35, max_columns=16)
                    )

                for label, frame, limit, cols in [
                    ('CATEGORY PERFORMANCE', category_performance, 30, 6),
                    ('TOP 20 SKUS', top20_skus, 20, 6),
                    ('WEEKLY SALES', weekly_sales, 20, 9),
                    ('WEEKLY CATEGORY PERFORMANCE', weekly_category, 35, 7),
                    ('MONTHLY SALES', monthly_sales, 15, 9),
                    ('MONTHLY CATEGORY PERFORMANCE', monthly_category, 30, 7),
                ]:
                    if frame is not None and not frame.empty:
                        context_sections.append(
                            label + ':\n' + compact_frame(frame, max_rows=limit, max_columns=cols)
                        )

            # Do not drop entire summaries just because one table is large.
            # Fit a smaller set of rows, always keeping the section heading.
            selected_sections = []
            remaining = max_context_chars
            shortened = []
            for section in context_sections:
                if remaining <= 150:
                    shortened.append(section.split('\n', 1)[0])
                    break
                if len(section) + 2 <= remaining:
                    selected_sections.append(section)
                    remaining -= len(section) + 2
                    continue
                heading, sep, rows = section.partition('\n')
                if sep:
                    lines = rows.splitlines()
                    kept = [heading, lines[0]] if lines else [heading]
                    for line in lines[1:]:
                        if len('\n'.join(kept + [line])) + 100 > remaining:
                            break
                        kept.append(line)
                    selected_sections.append('\n'.join(kept)[:remaining - 2])
                else:
                    selected_sections.append(section[:remaining - 2])
                shortened.append(heading)
                remaining = 0

            if shortened:
                st.caption(
                    'Gemini input was condensed to stay within the request size limit. '
                    'The dashboard and exported Excel file still retain their selected-period data; '
                    'AI answers may not include every SKU. Filter to a specific SKU/category '
                    'when you need exhaustive detail.'
                )

            context = '\n\n'.join(selected_sections)

            instructions = (

                'You are an Amazon India ecommerce performance analyst.\n'

                'Answer the user question first, using only the supplied data.\n'
                f'Selected dashboard reporting period: {start} through {end}. Do not use data outside this window except inventory snapshots explicitly labeled as-of.\n'
                'When a specific day is requested, use ONLY EXACT DAILY SKU SALES data for that date; '
                'never substitute weekly/monthly values. If the daily table is empty, state that clearly.\n'

                'Use exact SKU names, units, revenue, dates and percentage differences where available.\n'

                'Price history here is realized order ASP, NOT verified Amazon listing-price history. '

                'Never claim a listing price increased solely because ASP increased.\n'

                'For price crossing questions, prioritize the computed consecutive-week crossing table; '

                'compare Previous Units/Revenue with Units/Revenue and identify positive and negative changes.\n'

                'If no qualifying crossing is present, say so, and explain only relevant supported trends.\n'

                'Weeks run Sunday to Saturday. Partial or missing weeks may affect comparisons.\n'

                'Do not mix advertising-attributed sales with order revenue, or infer causation.\n'

                'For Top 20 SKUs by revenue, use the TOP 20 SKUS table and its Revenue Rank column; never infer revenue ranks from the COMBINED PERFORMANCE sample.\n'
                'Do not invent missing prices, SKUs or percentage changes.\n'

            )

            prompt = (instructions + '\nDATA:\n' + context

                      + '\n\nUSER QUESTION:\n' + question)

            # LANGCHAIN + GEMINI: retain the precomputed, date-filtered report context.
            # Use Gemini Developer API via GEMINI_API_KEY from .env.
            primary_model = (os.getenv('GEMINI_MODEL', 'gemini-2.5-flash').strip()
                             or 'gemini-2.5-flash')
            # Optional comma-separated fallback list in .env.
            fallback_models = [m.strip() for m in
                               os.getenv('GEMINI_FALLBACK_MODELS',
                                         'gemini-2.5-flash-lite').split(',')
                               if m.strip()]
            model_candidates = list(dict.fromkeys([primary_model] + fallback_models))
            st.caption(f'Primary Gemini Model: {primary_model} (LangChain)')

            # Prompt variables avoid accidental formatting of braces present in the data.
            prompt_template = ChatPromptTemplate.from_messages([
                ('system', '{analyst_instructions}'),
                ('human', 'DASHBOARD DATA:\n{report_context}\n\nUSER QUESTION:\n{user_question}')
            ])
            prompt_inputs = {
                'analyst_instructions': instructions,
                'report_context': context,
                'user_question': question,
            }

            answer = None
            used_model = None
            last_api_error = None

            def error_status(exc):
                """Support HTTP codes wrapped by LangChain / Google GenAI SDK."""
                for obj in (exc, getattr(exc, '__cause__', None),
                            getattr(exc, '__context__', None)):
                    if obj is None:
                        continue
                    for field in ('status_code', 'code'):
                        value = getattr(obj, field, None)
                        try:
                            if value is not None:
                                return int(value)
                        except (TypeError, ValueError):
                            pass
                match = re.search(r'\b(400|401|403|404|429|500|502|503|504)\b',
                                  str(exc))
                return int(match.group(1)) if match else None

            with st.spinner('LangChain + Gemini analyzing sales reports...'):
                for current_model in model_candidates:
                    # No nested SDK retries: our retry loop handles transient errors.
                    llm = ChatGoogleGenerativeAI(
                        model=current_model,
                        api_key=os.environ['GEMINI_API_KEY'],
                        temperature=0.2,
                        max_retries=0,
                        timeout=120,
                        max_output_tokens=2000,
                    )
                    chain = prompt_template | llm
                    for attempt in range(3):
                        try:
                            candidate = chain.invoke(prompt_inputs)
                            content = getattr(candidate, 'content', '')
                            if isinstance(content, str):
                                candidate_text = content
                            elif isinstance(content, list):
                                candidate_text = '\n'.join(
                                    part if isinstance(part, str)
                                    else str(part.get('text', ''))
                                    for part in content
                                    if isinstance(part, str) or
                                    (isinstance(part, dict) and part.get('type') == 'text')
                                )
                            else:
                                candidate_text = str(content or '')
                            if candidate_text.strip():
                                answer = candidate_text
                                used_model = current_model
                            else:
                                st.warning(f'{current_model} returned no text; trying fallback.')
                            break
                        except Exception as api_error:
                            last_api_error = api_error
                            status = error_status(api_error)
                            if status == 404:
                                st.warning(f'{current_model} unavailable (404); trying fallback.')
                                break
                            if status not in (429, 500, 502, 503, 504):
                                raise
                            if attempt == 2:
                                st.warning(f'{current_model}: API error {status} after retries; trying fallback.')
                                break
                            delay = min(20.0, 3.0 * (2 ** attempt) + random.uniform(0, 1))
                            st.warning(
                                f'{current_model}: temporary API error {status}. '
                                f'Retrying ({attempt + 1}/2) after {delay:.1f}s.'
                            )
                            time.sleep(delay)
                    if answer is not None:
                        break

            if answer is None:
                if last_api_error is not None:
                    raise last_api_error
                raise RuntimeError('No Gemini model returned a usable response.')
            if used_model != primary_model:
                st.info(f'Used fallback Gemini model: {used_model}')
            st.success('LangChain + Gemini analysis completed!')
            st.markdown(answer)

        except Exception as error:

            st.error(f'Gemini API Error: {type(error).__name__}: {error}')

            with st.expander('Show detailed error traceback'):

                st.code(traceback.format_exc(), language='text')

            if '404' in str(error) or 'NOT_FOUND' in str(error):
                st.info('Selected model is not available for this API key; check model access in Google AI Studio.')
            elif '503' in str(error) or 'UNAVAILABLE' in str(error):
                st.info('Gemini is overloaded even after retries. Your dashboard and Excel export remain usable.')
            elif '429' in str(error) or 'RESOURCE_EXHAUSTED' in str(error):
                st.info('Check rate limits and daily quota in Google AI Studio.')



# =====================================

# COMPLETE EXCEL EXPORT (MANUAL UPLOADS)

# =====================================

st.divider()

st.header('Download Complete Excel Report')

st.caption('Exports all currently selected dates, summary analyses and available report categories. '

           'It does not include data that was not uploaded.')

export_sheets = {

    'Category Performance': category_performance,

    'Top 20 SKUs': top20_skus,

    'SKU Weekly ASP': sku_weekly_history,

    'SKU Monthly ASP': sku_monthly_history,

    'SKU ASP Crossings': price_crossings,

    'Weekly Sales': weekly_sales,

    'Weekly Category': weekly_category,

    'Weekly Category Units': weekly_category_units.reset_index() if not weekly_category_units.empty else pd.DataFrame(),

    'Compare Two Weeks': week_comparison,

    'Weekly Category Compare': category_week_comparison,

    'Monthly Sales': monthly_sales,

    'Monthly Category': monthly_category,

    'Compare Two Months': month_comparison,

    'Monthly Category Compare': category_month_comparison,

    'Unified Period Trends': trends.reset_index() if series_parts else pd.DataFrame(),

    'Combined Analysis': combined.reset_index() if not combined.empty else pd.DataFrame(),

    'FBA Latest Stock': stock,

}

# Customize columns ONLY in the exported detail worksheets.
# Keep filtered data intact for dashboard, Gemini, FBA stock and other reports.
DETAIL_SHEET_NAMES = {
    'Orders': 'Orders Detail',
    'Ads': 'Ads Detail',
    'Traffic': 'Traffic Detail',
    'Customer Returns': 'Returns Detail',
    'FBA Inventory': 'FBA Detail',
}

DETAIL_COLUMNS_TO_REMOVE = {
    'Orders': ['FNSKU', 'Disposition', 'Campaign'],
    'Ads': ['FNSKU', 'Disposition', 'Zone'],
    'Traffic': ['FNSKU', 'Disposition', 'Campaign'],
    'Customer Returns': ['Campaign'],
    'FBA Inventory': ['Campaign'],
}

for kind, df in filtered.items():
    sheet_name = DETAIL_SHEET_NAMES.get(kind, kind)
    export_df = df.drop(
        columns=DETAIL_COLUMNS_TO_REMOVE.get(kind, []),
        errors='ignore'
    ).copy()

    # Orders already has a Zone column from prepare(). Keep it in the export.
    # If source Orders files have no recognized zone header, values are 'Unknown'.
    if kind == 'Orders' and 'Zone' not in export_df.columns:
        export_df['Zone'] = 'Unknown'

    export_sheets[sheet_name] = export_df

export_sheets['Dashboard KPIs'] = pd.DataFrame(metrics, columns=['Metric', 'Value'])

def excel_report(sheets):

    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd mmm yyyy',

                        date_format='dd mmm yyyy') as writer:

        workbook = writer.book

        header_fmt = workbook.add_format({'bold': True, 'bg_color':'#183153',

            'font_color':'white', 'border':0, 'valign':'vcenter'})

        money_fmt = workbook.add_format({'num_format':'₹#,##0.00;[Red](₹#,##0.00)'})

        count_fmt = workbook.add_format({'num_format':'#,##0'})

        pct_fmt = workbook.add_format({'num_format':'0.00"%";[Red](0.00"%")'})

        for sheet_name, data in sheets.items():

            if data is None or data.empty:

                continue

            # Excel's hard row limit; do not silently truncate data.

            if len(data) > 1_048_575:

                raise ValueError(f'{sheet_name}: more than 1,048,575 rows. Narrow the date filter.')

            clean = data.reset_index(drop=True).copy()

            for column in clean.select_dtypes(include=['datetimetz']).columns:

                clean[column] = clean[column].dt.tz_localize(None)

            clean.to_excel(writer, sheet_name=sheet_name[:31], index=False)

            ws = writer.sheets[sheet_name[:31]]

            ws.freeze_panes(1, 0)

            ws.set_row(0, 26, header_fmt)

            ws.autofilter(0, 0, len(clean), len(clean.columns)-1)

            for pos, field in enumerate(clean.columns):

                label = str(field)

                width = min(36, max(13, len(label)+3))

                if label in ('SKU', 'Category', 'Campaign', 'Source'):

                    width = min(32, max(width, 20))

                fmt = None

                if ('%' in label or 'Share' in label or 'Change %' in label):

                    fmt = pct_fmt

                elif any(x in label.lower() for x in ('revenue', 'sales', 'spend', 'refund', 'asp', 'difference')) and not any(x in label.lower() for x in ('units', 'quantity')):

                    fmt = money_fmt

                elif any(x in label.lower() for x in ('units', 'clicks', 'impressions', 'sessions', 'stock', 'balance')):

                    fmt = count_fmt

                ws.set_column(pos, pos, width, fmt)

    return output.getvalue()

# Generate large workbooks only when explicitly requested. Doing this on every

# Streamlit rerun can block rendering of sections further down the page.

# Invalidate an old prepared workbook when report signatures or filters change.

export_signature = repr((

    tuple((kind, tuple(state.get('signature', [])))

          for kind, state in sorted(st.session_state.unified_reports.items())),

    str(start), str(end),

    float(price_threshold) if not sku_weekly_history.empty else None,

    int(min_units) if not sku_weekly_history.empty else None,

))

if st.session_state.get('prepared_excel_signature') != export_signature:

    st.session_state.pop('prepared_excel_bytes', None)

    st.session_state['prepared_excel_signature'] = export_signature



if st.button('Prepare Complete Excel Report', type='primary'):

    try:

        with st.spinner('Building Excel workbook from uploaded reports...'):

            st.session_state['prepared_excel_bytes'] = excel_report(export_sheets)

        st.success('Excel report is ready to download.')

    except Exception as export_error:

        st.error(f'Excel export issue: {export_error}')



if st.session_state.get('prepared_excel_bytes') is not None:

    st.download_button(

        'Download Complete Excel Report (.xlsx)',

        data=st.session_state['prepared_excel_bytes'],

        file_name='Amazon_Unified_Sales_Report.xlsx',

        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

        use_container_width=True,

    )

st.caption('The workbook is prepared only when you click Prepare Complete Excel Report.')
