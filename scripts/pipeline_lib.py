"""
商品リサーチツール 共通データパイプライン（GitHub Actions実行用）。
クローズアウト版(CO_DEPTS指定)・全部門版(co_depts=None)の両方で共用する。

ローカル(PCデバイスブリッジ)実行版の build_research_tool.py / build_research_data.py と
ロジックは同一。相違点は、月の一覧をハードコードせず、ダウンロードしたファイル名から
動的に検出する点（新しい月のファイルが増えても自動で追従するため）。
"""
import re
import warnings

import numpy as np
import pandas as pd

CO_DEPTS = ['サニタリー', '健康食品', '化粧品Ａ', '化粧品Ｂ', '医療雑貨', '家庭用品', '日用雑貨']

STORE_ORDER = [
    '本部倉庫', '鶴来本店', '久安店', '安養寺店', '御経塚店', '小立野店', '月橋店', '高尾店',
    '山島台店', '県庁前店', '水戸町店', 'かわち店', '窪店', '根上店', '志賀店', '三納店',
    '駅西店', 'アピタ店', '加賀店', '白峰店', 'かほく店', '泉店', '吉野谷店', 'ＦＭ三納店', 'ＦＭ武蔵店', '桜田店',
]
STORE_ABBR = {
    '本部倉庫': '倉', '鶴来本店': '本', '久安店': '久', '安養寺店': '安', '御経塚店': '御',
    '小立野店': '小', '月橋店': '月', '高尾店': '高', '山島台店': '山', '県庁前店': '県',
    '水戸町店': '水', 'かわち店': 'か', '窪店': '窪', '根上店': '根', '志賀店': '志',
    '三納店': '三', '駅西店': '駅', 'アピタ店': 'ア', '加賀店': '加', '白峰店': '峰', 'かほく店': 'ほ',
    '泉店': '泉', '吉野谷店': '吉', 'ＦＭ三納店': 'F3', 'ＦＭ武蔵店': '武', '桜田店': '桜',
}
CHANNELS = ['外販課', 'ＥＣ店', '配置薬', 'ネット店']
CHANNEL_ABBR = {'外販課': '外', 'ＥＣ店': 'EC', '配置薬': '配', 'ネット店': 'ネ'}
STORE_ORDER_FULL = STORE_ORDER + CHANNELS
STORE_ABBR_FULL = {**STORE_ABBR, **CHANNEL_ABBR}
STORE_INDEX = {name: i for i, name in enumerate(STORE_ORDER_FULL)}

# 入荷案内.xlsx の店舗別数量列（短縮名）→ 正式店舗名
STORE_MAP = {
    '本部': '本部倉庫', '本店': '鶴来本店', '久安': '久安店', '安養寺': '安養寺店',
    '御経塚': '御経塚店', '小立野': '小立野店', '月橋': '月橋店', '高尾': '高尾店',
    '山島': '山島台店', '県庁前': '県庁前店', '水戸町': '水戸町店', 'かわち': 'かわち店',
    '窪': '窪店', '根上': '根上店', '志賀': '志賀店', '三納': '三納店',
    'アピタ': 'アピタ店', '加賀': '加賀店', '泉': '泉店', '武蔵': 'ＦＭ武蔵店',
    '吉野谷': '吉野谷店', '桜田': '桜田店',
}
DEPT_NORM = {'化粧品B': '化粧品Ｂ', '化粧品Ａ': '化粧品Ａ', '化粧品Ｂ': '化粧品Ｂ'}


def jan13(v):
    try:
        return str(int(round(float(v)))).zfill(13)
    except Exception:
        s = str(v).strip()
        return s.zfill(13) if s.isdigit() else None


def norm_dept(v):
    return DEPT_NORM.get(v, v)


def detect_months(filenames, prefix, suffix='.csv'):
    """ファイル名一覧から '202503' のような年月を検出し、month -> ファイル名 の辞書を返す。
    同じ月に '_copy' 付きファイルがあれば、そちらを優先する（データ差し替え用の慣習）。"""
    pattern = re.compile(re.escape(prefix) + r'(\d{6})(_copy)?' + re.escape(suffix) + r'$')
    result = {}
    for name in filenames:
        m = pattern.match(name)
        if not m:
            continue
        month, is_copy = m.group(1), bool(m.group(2))
        if month not in result or is_copy:
            result[month] = name
    return dict(sorted(result.items()))


def fmt_date_cell(v):
    import datetime as _dt
    if isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date)):
        ts = pd.Timestamp(v)
        if 2015 <= ts.year <= 2035:
            return ts.strftime('%Y-%m-%d')
        return None
    try:
        iv = int(float(v))
    except Exception:
        return None
    s = str(iv)[:8]
    if len(s) != 8:
        return None
    y, m, d = int(s[0:4]), int(s[4:6]), int(s[6:8])
    if not (2015 <= y <= 2035 and 1 <= m <= 12 and 1 <= d <= 31):
        return None
    return f'{y:04d}-{m:02d}-{d:02d}'


def clean_qty(s):
    return pd.to_numeric(s.astype(str).str.replace(',', '').str.strip(), errors='coerce')


def num_or_none(v):
    try:
        fv = float(v)
    except Exception:
        return None
    if np.isnan(fv) or np.isinf(fv):
        return None
    return fv


def build_master(mst_csv_path, co_depts=None):
    mst_cols = ['商品CD', '商品名', '部門名', 'POS部門名', 'クラス名', 'サブクラス名',
                '仕入先CD', '仕入先名', 'メーカー名', '原単価', '売単価']
    mst = pd.read_csv(mst_csv_path, encoding='cp932', usecols=mst_cols)
    if co_depts:
        mst = mst[mst['部門名'].isin(co_depts)].copy()
    mst['jan'] = mst['商品CD'].apply(jan13)
    mst = mst.drop_duplicates('jan', keep='last')
    mst = mst.rename(columns={'原単価': 'cost', '売単価': 'price'})
    mst['cost'] = pd.to_numeric(mst['cost'], errors='coerce').abs()
    mst['price'] = pd.to_numeric(mst['price'], errors='coerce').abs()
    for c in ['部門名', 'POS部門名', 'クラス名', 'サブクラス名', '仕入先名', 'メーカー名']:
        mst[c] = mst[c].fillna('').astype(str).str.strip()
    return mst


def build_genzaiko(genzaiko_csv_path, jan_set):
    with open(genzaiko_csv_path, encoding='cp932') as f:
        header0 = f.readline()
    period_m = re.search(r'(\d{4}/\d{1,2}/\d{1,2})\s*[～~]\s*(\d{4}/\d{1,2}/\d{1,2})', header0)
    period = f'{period_m.group(1)} 〜 {period_m.group(2)}' if period_m else None

    genzaiko = pd.read_csv(genzaiko_csv_path, encoding='cp932', skiprows=3, low_memory=False)
    genzaiko = genzaiko.rename(columns={'商品CD': 'jan', '商品名': 'name', '店舗名': 'row_type'})
    present_store_cols = [c for c in STORE_ORDER if c in genzaiko.columns]
    missing_stores = [c for c in STORE_ORDER if c not in genzaiko.columns]
    if missing_stores:
        print('警告: 現在庫ファイルに見つからない店舗列（閉店等の可能性。集計から除外）:', missing_stores)
    extra_stores = [c for c in genzaiko.columns
                     if c not in STORE_ORDER and c not in
                     ('jan', 'name', 'row_type', '外販課', 'ＥＣ店', '配置薬')
                     and not str(c).startswith('Unnamed') and '取扱店舗数' not in str(c)]
    if extra_stores:
        print('警告: STORE_ORDER未登録の新しい店舗列を検出（現在は集計対象外）:', extra_stores)

    genzaiko['jan'] = genzaiko['jan'].apply(jan13)
    genzaiko = genzaiko.dropna(subset=['jan'])
    for c in present_store_cols:
        genzaiko[c] = pd.to_numeric(genzaiko[c].astype(str).str.strip(), errors='coerce')
    genzaiko = genzaiko[genzaiko['jan'].isin(jan_set)].copy()
    genzaiko['stock_total'] = genzaiko[present_store_cols].sum(axis=1, skipna=True)
    return genzaiko, period, present_store_cols


def build_sales_monthly(main_files, tanpin_files, jan_set, co_depts=None):
    """main_files / tanpin_files: {month: local_path} の辞書"""
    monthly_rows = []
    for m, path in sorted(main_files.items()):
        df = pd.read_csv(path, encoding='cp932', usecols=['店舗名', '部門名', '商品CD', '売上数量'])
        if co_depts:
            df = df[df['部門名'].isin(co_depts)].copy()
        df['jan'] = df['商品CD'].apply(jan13)
        df = df[df['jan'].isin(jan_set)]
        df['qty'] = clean_qty(df['売上数量'])
        agg = df.groupby(['jan', '店舗名'], as_index=False)['qty'].sum()
        agg['month'] = m
        monthly_rows.append(agg)
        print(m, 'メイン: 対象行', len(df), '集計後', len(agg))

    for m, path in sorted(tanpin_files.items()):
        df = pd.read_csv(path, encoding='cp932', skiprows=3, usecols=['店舗名', '部門名', '商品CD', '売上数量'])
        if co_depts:
            df = df[df['部門名'].isin(co_depts)].copy()
        df['jan'] = df['商品CD'].apply(jan13)
        df = df[df['jan'].isin(jan_set)]
        df['qty'] = clean_qty(df['売上数量'])
        agg = df.groupby(['jan', '店舗名'], as_index=False)['qty'].sum()
        agg['month'] = m
        monthly_rows.append(agg)
        print(m, '単品(桜田/FM武蔵): 対象行', len(df), '集計後', len(agg))

    sales_monthly = pd.concat(monthly_rows, ignore_index=True)
    sales_monthly = sales_monthly.groupby(['jan', '店舗名', 'month'], as_index=False)['qty'].sum()
    return sales_monthly


def build_nyuka(nyuka_xlsx_path, jan_set):
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        nyuka_raw = pd.read_excel(nyuka_xlsx_path, sheet_name='政策Ｃ')
    nyuka_raw = nyuka_raw.rename(columns={
        '原価\n（税抜）': '原価_税抜', '売価\n（税抜）': '売価_税抜', '取引先\nコード': '取引先CD',
    })
    nyuka_raw['jan'] = nyuka_raw['JANCD'].apply(jan13)
    nyuka_raw = nyuka_raw.dropna(subset=['jan'])

    lot_rows = []
    for r in nyuka_raw.itertuples(index=False):
        jan = r.jan
        if jan not in jan_set:
            continue
        date_s = fmt_date_cell(getattr(r, '納品日'))
        stores = {}
        for short, full in STORE_MAP.items():
            v = getattr(r, short, None)
            if pd.notna(v) and v not in (0, '0'):
                try:
                    qv = int(round(float(v)))
                except Exception:
                    continue
                if qv != 0:
                    stores[full] = qv
        total = sum(stores.values())
        if total <= 0:
            continue
        cost_v = getattr(r, '原価_税抜', None)
        price_v = getattr(r, '売価_税抜', None)
        lot_rows.append({
            'jan': jan,
            'date': date_s,
            'dept': norm_dept(getattr(r, '部門')),
            'cost': cost_v if pd.notna(cost_v) else None,
            'price': price_v if pd.notna(price_v) else None,
            'total': total,
            'stores': stores,
        })
    lot_df = pd.DataFrame(lot_rows)
    return nyuka_raw, lot_df


def build_dict(series):
    vals = sorted(set(series.fillna('').astype(str).str.strip()))
    index = {v: i for i, v in enumerate(vals)}
    return vals, index


def merge_into_data(mst, genzaiko, genzaiko_period, present_store_cols,
                     sales_monthly, lots, active_jans, months):
    depts, dept_idx = build_dict(mst['部門名'])
    pos_depts, pos_dept_idx = build_dict(mst['POS部門名'])
    classes, class_idx = build_dict(mst['クラス名'])
    subclasses, subclass_idx = build_dict(mst['サブクラス名'])
    suppliers, supplier_idx = build_dict(mst['仕入先名'])
    makers, maker_idx = build_dict(mst['メーカー名'])

    master = []
    for r in mst.itertuples(index=False):
        jan = r.jan
        master.append([
            jan, r.商品名,
            dept_idx[r.部門名], pos_dept_idx[r.POS部門名],
            class_idx[r.クラス名], subclass_idx[r.サブクラス名],
            supplier_idx[r.仕入先名], maker_idx[r.メーカー名],
            num_or_none(r.cost), num_or_none(r.price),
            1 if jan in active_jans else 0,
        ])

    genzaiko_by_jan = genzaiko.set_index('jan')
    sales_by_jan = {jan: g for jan, g in sales_monthly.groupby('jan')}
    lots_by_jan = {}
    for r in lots.itertuples(index=False):
        lots_by_jan.setdefault(r.jan, []).append(r)

    month_index = {m: i for i, m in enumerate(months)}
    active = {}
    for jan in sorted(active_jans):
        entry = {}
        if jan in genzaiko_by_jan.index:
            grow = genzaiko_by_jan.loc[jan]
            pairs = []
            total = 0.0
            for store in present_store_cols:
                v = grow.get(store)
                if v is not None and not (isinstance(v, float) and np.isnan(v)) and v != 0:
                    qv = num_or_none(v)
                    if qv:
                        pairs.append([STORE_INDEX[store], qv])
                        total += qv
            if pairs:
                entry['stock'] = [pairs, total]

        if jan in sales_by_jan:
            g = sales_by_jan[jan]
            rows = []
            for r in g.itertuples(index=False):
                if r.qty and r.qty > 0 and r.店舗名 in STORE_INDEX and r.month in month_index:
                    rows.append([STORE_INDEX[r.店舗名], month_index[r.month], num_or_none(r.qty)])
            if rows:
                rows.sort(key=lambda x: (x[1], x[0]))
                entry['sales'] = rows

        if jan in lots_by_jan:
            lot_rows = []
            arrivals_acc = {}

            def _datekey(x):
                d = x.date
                return d if isinstance(d, str) else ''

            for r in sorted(lots_by_jan[jan], key=_datekey):
                store_pairs = []
                date_s = r.date if isinstance(r.date, str) else None
                for name, qty in (r.stores or {}).items():
                    if name not in STORE_INDEX:
                        continue
                    si = STORE_INDEX[name]
                    store_pairs.append([si, qty])
                    if date_s:
                        month = date_s[0:4] + date_s[5:7]
                        key = (si, month)
                        arrivals_acc[key] = arrivals_acc.get(key, 0) + qty
                lot_rows.append([
                    date_s, dept_idx.get(r.dept if isinstance(r.dept, str) else '', dept_idx.get('', 0)),
                    num_or_none(r.cost), num_or_none(r.price), num_or_none(r.total),
                    store_pairs,
                ])
            if lot_rows:
                entry['lots'] = lot_rows
            if arrivals_acc:
                arrivals = [[si, month, qty] for (si, month), qty in arrivals_acc.items()]
                arrivals.sort(key=lambda x: (x[1], x[0]))
                entry['arrivals'] = arrivals

        if entry:
            active[jan] = entry

    import datetime
    meta = {
        'generated_at': datetime.datetime.now().isoformat(),
        'genzaiko_period': genzaiko_period,
        'months': months,
        'current_partial_month': months[-1] if months else None,
        'store_order': STORE_ORDER_FULL,
        'store_abbr': STORE_ABBR_FULL,
        'depts': depts,
        'posDepts': pos_depts,
        'classes': classes,
        'subclasses': subclasses,
        'suppliers': suppliers,
        'makers': makers,
    }
    return {'meta': meta, 'master': master, 'active': active}
