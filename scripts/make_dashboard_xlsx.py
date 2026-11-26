"""Generate the original upload dashboard's two synthetic Excel workbooks."""
from pathlib import Path
import argparse
import json
from datetime import date
from calendar import monthrange

ROOT = Path(__file__).resolve().parents[1]
CATEGORIES = dict(zip(
    ['支架', '滤镜', '收纳包', '清洁套装', '转接头', '补光灯', '充电配件', '读卡设备', '镜头配件', '拍摄支撑'],
    ['数码支架', '滤镜', '相机包', '相机清洁/贴膜', '镜头附件', '闪光灯/手柄', '相机配件', '读卡器', '拍照配件', '三脚架/云台'],
))


def main():
    args = argparse.ArgumentParser(description=__doc__)
    args.add_argument('--out', type=Path, default=ROOT/'examples'/'dashboard-excel')
    out = args.parse_args().out
    try:
        from openpyxl import Workbook
    except ImportError:
        args.exit(2, 'Install the optional dependency: python -m pip install -r requirements-xlsx.txt\n')
    src = json.loads((ROOT/'examples'/'monthly-demo.json').read_text(encoding='utf-8'))
    if src.get('demo') is not True or src.get('format') != 'brand-monthly-demo':
        args.error('only known synthetic monthly fixtures can be converted')
    out.mkdir(parents=True, exist_ok=True)
    first, last = min(src['months']), max(src['months'])
    y, m = map(int, last.split('-'))
    stop = date(y, m, monthrange(y, m)[1]).isoformat()
    for idx, (mode, file) in enumerate([('自营','monthly-owner-demo.xlsx'), ('POP','monthly-pop-demo.xlsx')]):
        book = Workbook()
        book.properties.creator = 'Synthetic example'
        book.properties.title = 'Synthetic monthly data'
        rows = book.active
        rows.title = '数据'
        rows.append(['时间', '渠道', '三级品类', '标准品牌', '货盘类型', '销售额'])
        for month, brand, category, owner, pop, tb, tm, dy in src['rows']:
            when = f'{month[:4]}年{int(month[5:])}月'
            cat = CATEGORIES[category]
            pallet = ['一类货', '二类货', '三类货'][list(CATEGORIES).index(category) % 3]
            rows.append([when, 'JD', cat, brand, pallet, [owner, pop][idx]])
            for channel, amount in zip(['A-TB(非国际)', 'A-TM(大贸)', 'N-DY(非国际)'], [tb,tm,dy]):
                rows.append([when, channel, cat, brand, None, amount])
        scope = book.create_sheet('筛选条件')
        for line in [('筛选项','取值'),('时间范围',first+'-01 至 '+stop),('经营模式',mode),('是否含刷单','否'),('京东口径','模拟成交口径'),('是否包含二手','否'),('是否含清场','否'),('是否品类回算','否')]:
            scope.append(line)
        book.save(out/file)
        book.close()
    print('Wrote two synthetic XLSX workbooks.')


if __name__ == '__main__':
    main()
