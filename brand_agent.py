"""Offline, deterministic analysis engine for the Brand Market Share Agent."""
from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, getcontext
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import unicodedata

getcontext().prec = 40
ZERO = Decimal(0)
GMV = ('brand', 'category', 'current_gmv', 'prior_year_gmv')
KEY = ('brand',)
SHARE = ('week', 'brand', 'share')
KINDS = ('负增长', '正增长', '持平', '无基数', '不可比')
LABELS = ('风险', '预警', '次要')


class InputError(ValueError):
    pass


def text(value):
    return '' if value is None else str(value).strip()


def norm(value):
    return unicodedata.normalize('NFKC', value).strip().casefold()


def name(value, what, limit=180):
    if not isinstance(value, str):
        raise InputError(f'{what}必须是文本')
    value = value.strip()
    if not value or len(value) > limit or any(unicodedata.category(c) in ('Cc', 'Cf') for c in value):
        raise InputError(f'{what}为空、过长或含控制字符')
    return value


def num(value, *, blank=None, limit=Decimal('1e15')):
    if isinstance(value, bool):
        raise InputError('数字列不能使用布尔值')
    value = text(value)
    if not value:
        return blank
    if len(value) > 50 or not re.fullmatch(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]{1,3})?', value):
        raise InputError(f'数值格式不正确：{value!r}（请用无千分符、无百分号的数字）')
    try:
        d = Decimal(value)
    except InvalidOperation as e:
        raise InputError('无法解析数值') from e
    if not d.is_finite() or abs(d) > limit or (d and d.adjusted() < -12):
        raise InputError('数值超出支持范围')
    return d


def table(path, required):
    """CSV (UTF-8/BOM) and optionally XLSX; reject ambiguous headers/sheets."""
    path = Path(path)
    if path.suffix.lower() == '.csv':
        with path.open(encoding='utf-8-sig', newline='') as f:
            yield from checked_rows(csv.reader(f, strict=True), required, path.name)
    elif path.suffix.lower() == '.xlsx':
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise InputError('读取 .xlsx 请安装 requirements-xlsx.txt，或另存为 UTF-8 CSV') from exc
        # Normal mode avoids trusting stale worksheet dimensions. data_only uses saved formula values.
        try:
            wb = load_workbook(path, data_only=True)
        except Exception as exc:
            raise InputError(f'无法打开 XLSX {path.name}：{exc}') from exc
        try:
            choices = []
            for sheet in wb.worksheets:
                iterator = sheet.values
                header = next(iterator, ())
                if set(required).issubset({text(v) for v in header}):
                    choices.append((sheet.title, header, iterator))
            if len(choices) != 1:
                raise InputError(f'{path.name} 需要恰好一个首行包含 {", ".join(required)} 的工作表（找到 {len(choices)}）')
            from itertools import chain
            title, header, iterator = choices[0]
            yield from checked_rows(chain([header], iterator), required, f'{path.name}/{title}')
        finally:
            wb.close()
    else:
        raise InputError('仅支持 .csv 与 .xlsx')


def checked_rows(iterator, required, source):
    header = [text(x) for x in next(iterator, [])]
    missing = set(required) - set(header)
    nonempty = [k for k in header if k]
    if missing or len(set(nonempty)) != len(nonempty):
        raise InputError(f'{source} 表头缺字段或重复；缺少：{", ".join(sorted(missing)) or "无"}')
    for line, row in enumerate(iterator, 2):
        if not any(text(x) for x in row):
            continue
        if len(row) != len(header):
            raise InputError(f'{source} 第 {line} 行列数不符（空字段也需保留分隔符）')
        yield f'{source} 第 {line} 行', dict(zip(header, row))


def aliases(path):
    if not path:
        return {}
    value = json.loads(Path(path).read_text(encoding='utf-8-sig'), object_pairs_hook=unique_object)
    if not isinstance(value, dict):
        raise InputError('别名字典必须是 JSON 对象')
    result = {}
    for a, b in value.items():
        k = norm(name(a, '别名'))
        if k in result:
            raise InputError(f'归一化后出现重复别名：{a}')
        result[k] = name(b, '标准名')
    for k, b in result.items():
        if norm(b) in result and norm(b) != k:
            raise InputError(f'不支持链式别名；请直接映射到标准名：{b}')
    return result


def unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise InputError(f'JSON 出现重复键：{key}')
        result[key] = value
    return result


def week(value):
    if not isinstance(value, str):
        raise InputError('周标识必须是文本')
    m = re.fullmatch(r'([0-9]{4})\s*(?:-W|年第)\s*([0-9]{1,2})(?:周)?', value.strip(), flags=re.I)
    if not m or int(m[1]) < 1000:
        raise InputError(f'无效周标识：{value!r}；使用 YYYY-Www 或 YYYY年第n周')
    try:
        return date.fromisocalendar(int(m[1]), int(m[2]), 1)
    except ValueError as exc:
        raise InputError(f'无效 ISO 周：{value!r}') from exc


def week_label(d):
    y, w, _ = d.isocalendar()
    return f'{y:04d}-W{w:02d}'


def analyze(gmv_file, key_file, share_file, *, period, source, alias_path=None, expected_latest=None, demo=False):
    period, source = name(period, '经营周期', 80), name(source, '周市占来源', 100)
    alias = aliases(alias_path)
    display = {}

    def brand(value):
        value = name(value, '品牌')
        canonical = alias.get(norm(value), value)
        k = norm(canonical)
        display.setdefault(k, canonical)
        return k

    keys = set()
    for where, row in table(key_file, KEY):
        try:
            keys.add(brand(row['brand']))
        except InputError as e:
            raise InputError(f'{where}：{e}') from e
    if not keys:
        raise InputError('重点品牌名录为空；请核对导出是否完整')

    shares, days = {}, set()
    for where, row in table(share_file, SHARE):
        try:
            d, k = week(row['week']), brand(row['brand'])
            v = num(row['share'], limit=Decimal(1))
            if v is not None and v < 0:
                raise InputError('市占必须在 0–1 之间（如 12% 写 0.12）')
            if v is not None and v != v.quantize(Decimal('0.000000000001')):
                raise InputError('市占小数最多保留 12 位，以保证报告和 JSON 的趋势一致')
            if (d, k) in shares:
                raise InputError('同品牌同周出现重复记录（含别名合并后）；请先去重')
            days.add(d)
            shares[d, k] = v
        except InputError as e:
            raise InputError(f'{where}：{e}') from e
    window = sorted(days)[-3:]
    if len(window) != 3 or any((b - a).days != 7 for a, b in zip(window, window[1:])):
        raise InputError('周市占表最新三条周标签不是连续三周；不生成分层结果')
    if expected_latest and window[-1] != week(expected_latest):
        raise InputError(f'最新周不符：读到 {week_label(window[-1])}，预期 {expected_latest}')

    totals = defaultdict(lambda: defaultdict(lambda: [ZERO, ZERO]))
    profile = {'gmv_rows': 0, 'blank_prior_rows': 0, 'negative_current_rows': 0, 'negative_prior_rows': 0, 'repeated_brand_category_rows': 0, 'share_rows': len(shares), 'key_brands': len(keys)}
    for where, row in table(gmv_file, GMV):
        try:
            k, category = brand(row['brand']), name(row['category'], '类目')
            c, p = num(row['current_gmv']), num(row['prior_year_gmv'])
            if c is None:
                raise InputError('本期 GMV 为空')
            profile['gmv_rows'] += 1
            profile['blank_prior_rows'] += p is None
            profile['negative_current_rows'] += c < 0
            profile['negative_prior_rows'] += p is not None and p < 0
            profile['repeated_brand_category_rows'] += category in totals[k]
            pairs = totals[k][category]
            pairs[0] += c
            pairs[1] += p if p is not None else ZERO
        except InputError as e:
            raise InputError(f'{where}：{e}') from e
    if not totals:
        raise InputError('经营表没有数据行')

    rows = []
    for k, categories in totals.items():
        c = sum((x[0] for x in categories.values()), ZERO)
        p = sum((x[1] for x in categories.values()), ZERO)
        growth = (c - p) / p if p > 0 else None
        kind = ('不可比' if p < 0 else '无基数' if p == 0 else '负增长' if c < p else '持平' if c == p else '正增长')
        vs = [shares.get((d, k)) for d in window]
        present = any((d, k) in shares for d in window)
        label = reason = None
        if kind == '负增长':
            if k not in keys:
                label, reason = '次要', '经营同比为负；非重点品牌。'
            elif any(v is None for v in vs):
                label = '预警'
                reason = '重点品牌，经营同比为负；' + ('部分或全部周市占缺数，无法判断。' if present else '未匹配到该品牌在此窗口的周市占记录，无法判断。')
            elif vs[0] > vs[1] > vs[2]:
                label, reason = '风险', '重点品牌，经营同比为负；近三周市占持续走低。'
            else:
                label, reason = '预警', '重点品牌，经营同比为负；近三周市占未持续走低。'
        rows.append(dict(brand=display[k], keyBrand=k in keys, current=c, prior=p, yoy=growth, kind=kind, shares=vs, label=label, reason=reason, categories=categories))
    rows.sort(key=lambda b: (-b['current'], norm(b['brand']), b['brand']))
    return dict(period=period, source=source, weeks=[week_label(d) for d in window], rows=rows, demo=demo, profile=profile)


def risk_list(result):
    def wire_share(v):
        return None if v is None else float(v)
    content = dict(format='brand-risk-list', version=1, demo=result['demo'], gmvPeriod=result['period'], shareSource=result['source'], weeks=result['weeks'], brands=[])
    for b in result['rows']:
        if b['label']:
            content['brands'].append(dict(brand=b['brand'], label=b['label'], keyBrand=b['keyBrand'], yoy=float(b['yoy']), shares=[wire_share(v) for v in b['shares']], reason=b['reason']))
    return content


def money(x):
    return f'{x:,.2f}'


def pct(x, places=2):
    return '—' if x is None else f'{(x * 100).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP):.{places}f}%'


def week_pct(x):
    return '缺数' if x is None else f'{x*100:.10f}'.rstrip('0').rstrip('.')+'%'


def md(x):
    # Report fields are data, not HTML/Markdown syntax.
    from html import escape
    value = escape(' '.join(str(x).split()), quote=False)
    return re.sub(r'([\\|`*_\[\]~])', r'\\\1', value)


def report(result):
    rows, w = result['rows'], result['weeks']
    neg = [b for b in rows if b['kind'] == '负增长']
    other = [b for b in rows if b['kind'] != '负增长']
    risk = [b for b in neg if b['label'] == '风险']
    def total(items):
        return sum((b['current'] for b in items), ZERO)
    def clipped(items, n):
        return f'共 {len(items)} 个，展示 {min(n, len(items))} 个；未展示 {max(0, len(items)-n)} 个，合计当期 GMV {money(total(items[n:]))}。'
    def line(values):
        return '| ' + ' | '.join(md(x) for x in values) + ' |'
    def heading(header):
        return [line(header), line(['---'] * len(header))]
    def brand_row(b):
        return [b['brand'], b['label'] or b['kind'], money(b['current']), money(b['prior']), pct(b['yoy']), *[week_pct(v) for v in b['shares']]]
    cols = ['品牌', '标记', '当期 GMV', '同期 GMV', '同比', *w]
    parts = ['# 品牌市占分析报告', '', ('**模拟数据，不用于经营决策。**' if result['demo'] else '**依据本次输入生成。**'), '',
             f'经营周期：{md(result["period"])}；周市占：{md(result["source"])}；观察窗口：{", ".join(w)}。', '',
             '## 一、总览', '', *heading(['分组', '品牌数', '当期 GMV'])]
    for label in LABELS:
        group = [b for b in neg if b['label'] == label]
        parts.append(line([f'负增长 · {label}', len(group), money(total(group))]))
    for kind in KINDS[1:]:
        group = [b for b in rows if b['kind'] == kind]
        parts.append(line([kind, len(group), money(total(group))]))
    parts += [line(['合计', len(rows), money(total(rows))]), '', f'风险品牌 {len(risk)} 个；完整标签见风险 JSON。', '',
              '## 二、同比负增长品牌（GMV 降序 Top 100）', '', clipped(neg, 100), '', *heading(cols)]
    parts += [line(brand_row(b)) for b in neg[:100]]
    parts += ['', '## 三、深度分析（同比负增长品牌按 GMV 降序 Top 30）', '', clipped(neg, 30), '', '风险标签不参与排序。']
    for b in neg[:30]:
        parts += ['', f'### {md(b["brand"])}（{b["label"]}）', '', md(b['reason']), '',
                  '周市占：' + ' → '.join(week_pct(v) for v in b['shares']) + '。', '']
        categories = sorted(b['categories'].items(), key=lambda pair: (-pair[1][0], pair[0]))
        parts += heading(['类目', '本期 GMV', '同期 GMV', '差额（本期−同期）', '类目同比'])
        for cat, (c, p) in categories[:8]:
            parts.append(line([cat, money(c), money(p), money(c-p), pct((c-p)/p if p > 0 else None)]))
        if len(categories) > 8:
            tail = categories[8:]
            c, p = (sum((v[i] for _, v in tail), ZERO) for i in range(2))
            parts.append(line([f'其余 {len(tail)} 项', money(c), money(p), money(c-p), pct((c-p)/p if p > 0 else None)]))
        gaps = [(k, c-p) for k,(c,p) in categories if c < p]
        gaps.sort(key=lambda item: (item[1], item[0]))
        parts += ['', '经营 GMV 缺口：' + ('；'.join(f'{md(k)}（{money(x)}）' for k,x in gaps[:3]) or '无') + '。该差额不证明周市占变化原因。']
    parts += ['', '## 四、其他品牌（GMV 降序 Top 30）', '', clipped(other, 30), '', *heading(cols)]
    parts += [line(brand_row(b)) for b in other[:30]]
    parts += ['', '## 附：口径与输入校验', '',
              '- 同比 = 品牌汇总本期 GMV ÷ 汇总去年同期 GMV − 1；同期空值计 0，整列缺失和非空非法数值报错。同期合计为 0 记无基数，为负记不可比；实际相等记持平。',
              '- 近三周市占持续走低指两个周环比连续下降，到第三周确认；持平不算下降；未匹配或单周缺数不补 0。',
              '- 市占来自独立的周度输入，不由月表推算；本报告的经营类目差额仅用来定位排查线索。',
              '- 输入画像：' + '；'.join(f'{label} {result["profile"][key]}' for key,label in [('gmv_rows','经营行数'),('blank_prior_rows','同期空值行'),('negative_current_rows','负本期 GMV 行'),('negative_prior_rows','负同期 GMV 行'),('repeated_brand_category_rows','重复品牌类目行'),('share_rows','周市占记录'),('key_brands','重点品牌数')]) + '。',
              '- 负 GMV 和重复品牌类目行按原值加总；使用前请确认这些不是错误记录。', '']
    return '\n'.join(parts)


def write_atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix='.brand-agent-', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as file:
            file.write(data)
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    for field in ('gmv', 'key-brands', 'weekly-share', 'period'):
        p.add_argument('--'+field, required=True)
    p.add_argument('--share-source', default='本地周市占快照')
    p.add_argument('--aliases')
    p.add_argument('--expected-latest-week')
    p.add_argument('--demo', action='store_true')
    p.add_argument('--out', default='output', type=Path)
    a = p.parse_args(argv)
    try:
        result = analyze(a.gmv, a.key_brands, a.weekly_share, period=a.period, source=a.share_source, alias_path=a.aliases, expected_latest=a.expected_latest_week, demo=a.demo)
        structured = risk_list(result)
        markdown = report(result)
        encoded = json.dumps(structured, ensure_ascii=False, allow_nan=False, indent=2)+'\n'
        if len(structured['brands']) > 10000 or len(encoded.encode('utf-8')) > 5*1024*1024:
            raise InputError('风险清单超出 v1 支持范围（10,000 品牌或 5 MiB），请缩小输入范围；未截断')
        write_atomic(a.out/'report.md', markdown)
        write_atomic(a.out/'brand-risk-list.json', encoded)
    except (ValueError, OSError, csv.Error, ArithmeticError) as err:
        print(f'分析失败：{err}', file=sys.stderr)
        return 2
    print(f'已生成报告和风险清单：{a.out}（{len(result["rows"])} 个品牌；负增长 {len(structured["brands"])} 个）')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
