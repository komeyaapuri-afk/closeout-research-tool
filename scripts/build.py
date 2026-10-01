"""
GitHub Actions から実行されるメインスクリプト。
1. Google Driveから必要なファイルを取得
2. クローズアウト7部門版・全部門版、それぞれのデータを構築
3. テンプレートに埋め込み、closeout.html / zenbumon.html を生成
"""
import base64
import gzip
import json
import os
import sys

import drive_io
import pipeline_lib as pl

DATA_DIR = 'data'
POS_FOLDER_ID = os.environ['DRIVE_FOLDER_POS']
GENZAIKO_FOLDER_ID = os.environ['DRIVE_FOLDER_GENZAIKO']
NYUKA_FOLDER_ID = os.environ['DRIVE_FOLDER_NYUKA']

NYUKA_FILENAME = '【医療・雑貨クローズアウト】入荷案内.xlsx'
GENZAIKO_FILENAME = '現在庫.csv'
MASTER_FILENAME = '商品マスタ.csv'


def fetch_all():
    service = drive_io.get_drive_service()
    os.makedirs(DATA_DIR, exist_ok=True)

    print('=== Google Driveからファイル一覧を取得 ===')
    pos_files = drive_io.list_files_in_folder(service, POS_FOLDER_ID)
    genzaiko_files = drive_io.list_files_in_folder(service, GENZAIKO_FOLDER_ID)
    nyuka_files = drive_io.list_files_in_folder(service, NYUKA_FOLDER_ID)
    pos_names = [f['name'] for f in pos_files]
    print('POSジャーナルフォルダ内ファイル数:', len(pos_names))

    main_months = pl.detect_months(pos_names, 'POSジャーナル_', '.csv')
    tanpin_months = pl.detect_months(pos_names, '単品ジャーナル(桜田FM武蔵)_', '.csv')
    print('検出したメイン月:', list(main_months.keys()))
    print('検出した単品(桜田/FM武蔵)月:', list(tanpin_months.keys()))
    if not main_months:
        print('エラー: POSジャーナルの月次ファイルが1件も検出できませんでした', file=sys.stderr)
        sys.exit(1)

    print('商品マスタ・POSジャーナル各月をダウンロード中...')
    mst_path = drive_io.download_by_name(service, pos_files, MASTER_FILENAME, DATA_DIR)
    main_paths = {m: drive_io.download_by_name(service, pos_files, name, DATA_DIR)
                  for m, name in main_months.items()}
    tanpin_paths = {m: drive_io.download_by_name(service, pos_files, name, DATA_DIR)
                    for m, name in tanpin_months.items()}

    print('現在庫・入荷案内をダウンロード中...')
    genzaiko_path = drive_io.download_by_name(service, genzaiko_files, GENZAIKO_FILENAME, DATA_DIR)
    nyuka_path = drive_io.download_by_name(service, nyuka_files, NYUKA_FILENAME, DATA_DIR)

    months = sorted(set(main_months.keys()) | set(tanpin_months.keys()))
    return {
        'mst_path': mst_path,
        'main_paths': main_paths,
        'tanpin_paths': tanpin_paths,
        'genzaiko_path': genzaiko_path,
        'nyuka_path': nyuka_path,
        'months': months,
    }


def build_scope(sources, co_depts, label):
    print(f'\n=== {label}: データ構築開始 ===')
    mst = pl.build_master(sources['mst_path'], co_depts=co_depts)
    jan_set = set(mst['jan'])
    print(f'{label}: 商品マスタ件数', len(mst))

    genzaiko, genzaiko_period, present_store_cols = pl.build_genzaiko(sources['genzaiko_path'], jan_set)
    print(f'{label}: 現在庫あり商品数', (genzaiko['stock_total'] > 0).sum())

    sales_monthly = pl.build_sales_monthly(sources['main_paths'], sources['tanpin_paths'], jan_set, co_depts=co_depts)
    print(f'{label}: 販売月次集計 総行数', len(sales_monthly))

    nyuka_raw, lots = pl.build_nyuka(sources['nyuka_path'], jan_set)
    print(f'{label}: 仕入れロット件数', len(lots))

    genzaiko_active = set(genzaiko.loc[genzaiko['stock_total'] > 0, 'jan'])
    sales_active = set(sales_monthly.loc[sales_monthly['qty'] > 0, 'jan'].unique())
    nyuka_active = set(lots['jan'].unique()) if len(lots) else set()
    active_jans = genzaiko_active | sales_active | nyuka_active
    print(f'{label}: 動いている商品 母集団', len(active_jans))

    data = pl.merge_into_data(mst, genzaiko, genzaiko_period, present_store_cols,
                               sales_monthly, lots, active_jans, sources['months'])
    print(f'{label}: active件数', len(data['active']))
    return data


def embed(data, template_path, out_path):
    raw = json.dumps(data, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    comp = gzip.compress(raw, compresslevel=9)
    b64 = base64.b64encode(comp).decode('ascii')
    with open(template_path, encoding='utf-8') as f:
        tpl = f.read()
    if '%%DATA_GZ_B64%%' not in tpl:
        print(f'エラー: {template_path} にプレースホルダーが見つかりません', file=sys.stderr)
        sys.exit(1)
    out = tpl.replace('%%DATA_GZ_B64%%', b64)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(out)
    print(f'生成: {out_path} ({os.path.getsize(out_path):,} bytes)')


def main():
    sources = fetch_all()

    co_data = build_scope(sources, pl.CO_DEPTS, 'クローズアウト7部門')
    embed(co_data, 'templates/research_template.html', 'closeout.html')

    all_data = build_scope(sources, None, '全部門')
    embed(all_data, 'templates/alldept_template.html', 'zenbumon.html')

    print('\n完了')


if __name__ == '__main__':
    main()
