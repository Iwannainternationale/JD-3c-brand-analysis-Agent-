"""Build public fixtures from invented names/numbers; no original project assets are read."""
from pathlib import Path
import csv
import json
import math
import sys

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = ROOT / 'examples'
BRANDS = ['样例棱镜', '样例风帆', '样例晨光', '样例云桥', '样例青石', '样例银杉', '样例星轨', '样例流萤',
          '样例远峰', '样例拾光', '样例灯塔', '样例涟漪', '样例荒原', '样例松塔', '样例初雪', '样例岩雀']
ABSENT = '样例外圈'
CATS = ['支架', '滤镜', '收纳包', '清洁套装', '转接头', '补光灯', '充电配件', '读卡设备', '镜头配件', '拍摄支撑']
# Signed differences for synthetic business export; not derived from the separate monthly data.
DELTA = [-6200,-1900,-4500,-2400,-2200,-850,-3700,-1200,-1800,-1100,-960,2900,0,0,0,2200,-1250]


def csv_file(name, header, rows):
    with (EXAMPLES / name).open('w', encoding='utf-8', newline='') as file:
        w = csv.writer(file, lineterminator='\n')
        w.writerow(header)
        w.writerows(rows)


def write_json(path, obj, indent=2):
    path.write_text(json.dumps(obj, ensure_ascii=False, allow_nan=False, indent=indent)+'\n', encoding='utf-8')


def main():
    EXAMPLES.mkdir(exist_ok=True)
    rows = []
    for i, brand in enumerate(BRANDS+[ABSENT]):
        base = 22600 + ((i * 1279) % 9200) if i != 16 else 3600
        cats = [CATS[(i+n)%10] for n in (0,3,7)]
        for n, cat in enumerate(cats):
            p = int(base * (.53,.31,.16)[n])
            c = p + int(DELTA[i] * (.54,.34,.12)[n])
            if i == 13:
                p = '' if n == 0 else 0
            if i == 14:
                p = -120 if n == 0 else 0
            rows.append(('棱镜配件（样例别名）' if i == 0 else brand, cat, c, p))
    csv_file('gmv.csv', ('brand','category','current_gmv','prior_year_gmv'), rows)
    csv_file('key-brands.csv', ('brand',), [(x,) for x in BRANDS[:8]+[ABSENT]])
    trend = [(0.265,0.241,0.218),(0.173,None,0.153),(0.306,0.273,0.251),(.246,.217,.197),(.123,.104,.111),(.151,.151,.142),(.194,.165,.14),(None,None,None),
             (.22,.2,.19),(0,None,.04),(.11,.135,.115)]
    trend += [(0.1,0.11,0.12)] * 5 + [(.198,.161,.128)]
    weeks = ['2025-W51','2025-W52','2026-W01']
    weekly = [(w, b, v) for b, vs in zip(BRANDS+[ABSENT], trend) if b != BRANDS[7] for w,v in zip(weeks,vs)]
    csv_file('weekly-share.csv', ('week','brand','share'), weekly)
    write_json(EXAMPLES/'aliases.json', {'棱镜配件（样例别名）': BRANDS[0]})
    months = [f'{2024+(11+i)//12:04d}-{(11+i)%12+1:02d}' for i in range(22)]
    sales = []
    for mois, month in enumerate(months):
        for i,b in enumerate(BRANDS):
            if (i == 1 and mois > 19) or (i == 15 and mois < 10):
                continue
            for n,offset in enumerate((0,3,7)):
                seed = 13500 + (i*1531)%16000
                factor = (.82, .83, .97, 1.00, 1.07, 1.29, .95, 1.0, 1.06, 1.11, 1.39, 1.25)[int(month[-2:])-1]
                month_factor = 1 + (mois/21)*(-.24 if i < 7 else .2) + .1*math.sin(mois*.62+i+n)
                jd = round(seed*factor*month_factor*(.63,.26,.15)[n])
                owner = round(jd*(.42+.04*(i%4)))
                competitors = [round(seed * factor * f*(.63,.26,.15)[n]*(1+.009*mois)) for f in (.65,.44,.29)]
                sales.append([month,b,CATS[(i+offset)%10],owner,jd-owner,*competitors])
    snapshot = {'format':'brand-monthly-demo','version':1,'demo':True,'months':months,'rows':sales}
    write_json(EXAMPLES/'monthly-demo.json', snapshot, None)
    template = (ROOT/'dashboard'/'template.html').read_text(encoding='utf-8')
    payload = json.dumps(snapshot, ensure_ascii=False, separators=(',', ':')).replace('<', r'\u003c')
    sys.path.insert(0, str(ROOT))
    from brand_agent import analyze, report, risk_list
    result = analyze(EXAMPLES/'gmv.csv', EXAMPLES/'key-brands.csv', EXAMPLES/'weekly-share.csv', period='2025-12（模拟）', source='本地模拟周市占', alias_path=EXAMPLES/'aliases.json', expected_latest='2026-W01', demo=True)
    expected = EXAMPLES/'expected'; expected.mkdir(exist_ok=True)
    list_data = risk_list(result)
    write_json(expected/'brand-risk-list.json', list_data)
    (expected/'report.md').write_text(report(result), encoding='utf-8', newline='\n')
    risk_data = json.dumps(list_data, ensure_ascii=False, separators=(',', ':')).replace('<', r'\u003c')
    assert template.count('/*__DEMO__* /') == template.count('/*__RISK__* /') == 1
    (ROOT/'dashboard'/'index.html').write_text(template.replace('/*__DEMO__* /',payload).replace('/*__RISK__* /',risk_data), encoding='utf-8',newline='\n')
    print(f'{len(months)} months / {len(BRANDS)} monthly brands / {len(CATS)} categories / {len(rows)} GMV rows')


if __name__ == '__main__':
    main()
