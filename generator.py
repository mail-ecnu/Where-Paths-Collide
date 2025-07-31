import argparse
import re
from pathlib import Path
from datetime import datetime

try:
    from zoneinfo import ZoneInfo   # Python 3.9+
except Exception:
    ZoneInfo = None

import pandas as pd


# ---------- 工具：安全字符串 ----------
def s(v):
    """将任意单元格值安全转为去首尾空格的字符串；None/NaN -> ''。"""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


# ---------- 标签外观（与示例一致） ----------
TAG_STYLES = {
    'Search-Based':            {'color': '#E4D9EE', 'icon': 'fa-book'},
    'Reinforcement learning':  {'color': '#F4C9C9', 'icon': 'fa-chess-knight'},
    'Supervised learning':     {'color': '#F4C9C9', 'icon': 'fa-chess-knight'},
    'Curriculum learning':     {'color': '#F4C9C9', 'icon': 'fa-chess-knight'},
    'CBS':                     {'color': '#F4C9C9', 'icon': 'fa-chess-knight'},
    'Optimal':                 {'color': '#FFDFC2', 'icon': 'fa-puzzle-piece'},
    'Sub-optimal':             {'color': '#FFDFC2', 'icon': 'fa-puzzle-piece'},
    'One-shot':                {'color': '#C7DDEC', 'icon': 'fa-atom'},
    'Lifelong':                {'color': '#C7DDEC', 'icon': 'fa-infinity'},
    'Discrete space':          {'color': '#CAE7CA', 'icon': 'fa-earth-americas'},
    'Continuous space':        {'color': '#D2ECEF', 'icon': 'fa-earth-americas'},
}

# 标签展示顺序（先方法，再算法，再最优性，再任务类型，再空间）
TAG_ORDER_PRIORITY = [
    ['Search-Based', 'Reinforcement learning', 'Supervised learning', 'Curriculum learning'],
    ['CBS'],
    ['Optimal', 'Sub-optimal'],
    ['One-shot', 'Lifelong'],
    ['Discrete space', 'Continuous space'],
]

# CCF 排序优先级
CCF_RANK = {'A': 0, 'B': 1, 'C': 2, 'None': 3, 'Preprint': 4}

# Excel 的 Type → 模板中的锚点
CATEGORY_ANCHOR_MAP = {
    'survey': 'survey',
    'benchmark': 'benchmark',                      # 兼容 singular
    'benchmarks': 'benchmark',
    'classical': 'classic',
    'augmented': 'Learning-Augmented Classic Solvers',  # 模板中的真实锚点名
    'learning': 'learning',
}

# 锚点 → 目录中 title 的 CSS 类名（用于定位 ToC 每个分类的年份行）
ANCHOR_TOC_CLASS = {
    'survey': 'survey',
    'benchmark': 'benchmark',
    'classic': 'classic',
    'Learning-Augmented Classic Solvers': 'augmented',
    'learning': 'learning',
}


def get_entry_rank(row):
    """根据 CCF 等级返回排序权重（A最高，Preprint最低）。"""
    conf_ccf = s(row.get('CCF', ''))
    journal_ccf = s(row.get('CCF.1', ''))
    conf = s(row.get('Conference', ''))
    journal = s(row.get('Journal', ''))
    if conf:
        rating = conf_ccf or "None"
    elif journal:
        rating = journal_ccf or "None"
    else:
        rating = conf_ccf or journal_ccf or "Preprint"
    return CCF_RANK.get(rating if rating in CCF_RANK else "None", 3)


def get_venue_sort_key(row):
    """同一年内排序：先按 CCF，再按 venue 名称。"""
    rank = get_entry_rank(row)
    conf = s(row.get('Conference', ''))
    journal = s(row.get('Journal', ''))
    venue = conf if conf else (journal if journal else "ZZZ")
    return (rank, venue)


def to_year(v):
    """从 Release Time 中提取 4 位年份（支持 'IJCAI, 2021' 或 '2021' 等格式）。"""
    vs = s(v)
    m = re.search(r'(\d{4})', vs)
    return int(m.group(1)) if m else None


def build_entry_html(entry, year):
    """生成一条论文的 <p>...</p> HTML。"""
    title   = s(entry.get('Paper Title', ''))
    link    = s(entry.get('doi', '')) or s(entry.get('Google Scholar', '')) or "#"
    authors = s(entry.get('Authors (Split by , and space)', '')).rstrip(',')
    conf    = s(entry.get('Conference', ''))
    journal = s(entry.get('Journal', ''))

    # venue 文本：会议优先（补年），否则期刊，再否则 arXiv
    if conf:
        conf_clean = re.sub(r',\s*(\d{4})$', r' \1', conf)
        venue_text = conf_clean if re.search(r'\d{4}$', conf_clean) else f"{conf_clean} {year}"
    elif journal:
        venue_text = journal
    else:
        venue_text = f"arXiv {year}"

    parts = []
    parts.append('<p>')
    parts.append(
        f'  <a href="{link}" style="color: #54a1cb; text-decoration: none; font-weight: 600;">{title}</a>'
    )

    # CCF 徽章
    ccf_value = ""
    conf_ccf = s(entry.get('CCF', ''))
    journal_ccf = s(entry.get('CCF.1', ''))
    if conf and conf_ccf:
        ccf_value = conf_ccf
    elif journal and journal_ccf:
        ccf_value = journal_ccf
    elif conf_ccf == 'Preprint' or journal_ccf == 'Preprint':
        ccf_value = 'Preprint'

    if ccf_value:
        if ccf_value == 'A':
            color = label_color = "red"
        elif ccf_value == 'B':
            color = label_color = "orange"
        elif ccf_value == 'C':
            color = label_color = "green"
        else:
            color = label_color = "lightgrey"
            if ccf_value not in ['None', 'Preprint']:
                ccf_value = "None"
        parts.append(
            f'  <img src="https://img.shields.io/badge/CCF-{ccf_value}-{color}?style=flat&labelColor={label_color}" '
            f'alt="CCF {ccf_value}" style="height:1.2em; width:auto; vertical-align:middle;" />'
        )

    # 作者/venue 行
    parts.append('<br>')
    av = []
    if authors: av.append(authors)
    if venue_text: av.append(venue_text)
    if av:
        parts.append(f'  <span style="color: #808080;">{" &#8212; ".join(av)}</span><br>')

    # 右侧徽标/标签区域
    span_bits = []

    # GitHub 星标（Repo）
    repo = s(entry.get('Repo', ''))
    if repo:
        repo_main = repo.split()[0]   # 多个链接取第一个
        if 'github.com/' in repo_main:
            repo_path = repo_main.split('github.com/')[1].strip().strip('/')
            if repo_path:
                span_bits.append(
                    f'<a href="{repo_main}">\n'
                    f'    <img src="https://img.shields.io/github/stars/{repo_path}?style=social&label=Official" '
                    f'alt="Official" style="height:1.2em; width:auto; vertical-align:middle;" />\n'
                    f'  </a>'
                )

    # 标签：方法/算法/最优性/任务类型/环境
    tags = []
    typ = s(entry.get('Type', '')).lower()
    if typ in ['classical', 'augmented']:
        tags.append('Search-Based')

    methods_str = (s(entry.get('Methods1', '')) + ' ' + s(entry.get('Methods2', ''))).lower()
    if 'reinforcement learning' in methods_str:
        tags.append('Reinforcement learning')
    if 'supervised learning' in methods_str:
        tags.append('Supervised learning')
    if 'curriculum learning' in methods_str:
        tags.append('Curriculum learning')
    if 'cbs' in methods_str and 'CBS' not in tags:
        tags.append('CBS')

    qual_raw = s(entry.get('Solution Quality', ''))
    if qual_raw:
        span_bits.append(
            f'<a href="" style="text-decoration: none;">'
            f'<span style="background-color: #FFDFC2; padding: 4px 8px; '
            f'border-radius: 50px; font-size: calc(100% - 4pt);">'
            f'<i class="fas fa-puzzle-piece"></i> {qual_raw} </span> </a>'
        )

    mission = s(entry.get('Mission Type', ''))
    if mission in ['One-shot', 'Lifelong']:
        tags.append(mission)

    env = s(entry.get('Env.', ''))
    if env in ['Discrete space', 'Continuous space']:
        tags.append(env)

    # 去重并按优先级排序
    seen = set()
    tags = [t for t in tags if not (t in seen or seen.add(t))]
    sorted_tags, rest = [], tags[:]
    for group in TAG_ORDER_PRIORITY:
        for t in group:
            if t in rest:
                sorted_tags.append(t)
                rest.remove(t)
    sorted_tags.extend(rest)

    for tag in sorted_tags:
        style = TAG_STYLES.get(tag)
        if style:
            span_bits.append(
                f'<a href="" style="text-decoration: none;">'
                f'<span style="background-color: {style["color"]}; padding: 4px 8px; '
                f'border-radius: 50px; font-size: calc(100% - 4pt);">'
                f'<i class="fas {style["icon"]}"></i> {tag} </span> </a>'
            )

    if span_bits:
        parts.append('  <span style="display:inline-flex; gap:8px; align-items:center;">')
        parts.extend(['  ' + b for b in span_bits])
        parts.append('  </span>')

    parts.append('</p>')
    return "\n".join(parts)


def rebuild_toc(html: str, years_by_anchor: dict) -> str:
    """
    重建 Table of Contents 内每个分类的年份行：
    仅保留有论文的年份；若没有年份则年份行为空（保留分类标题行以稳布局）。
    """
    out = html

    for anchor, years in years_by_anchor.items():
        css = ANCHOR_TOC_CLASS.get(anchor)
        if not css:
            continue

        # 找到该分类所在的 ToC 表格的“标题<tr>…</tr>”及其后面的“年份<tr>…</tr>”
        title_marker = f'<td class="title {css}"'
        title_idx = out.find(title_marker)
        if title_idx == -1:
            continue

        # 找标题行 <tr ...> 与 </tr>
        tr_title_start = out.rfind('<tr', 0, title_idx)
        tr_title_end = out.find('</tr>', title_idx)
        if tr_title_start == -1 or tr_title_end == -1:
            continue
        tr_title_end += len('</tr>')

        # 在同一个表格内找到紧随其后的年份行
        table_end = out.find('</table>', tr_title_end)
        if table_end == -1:
            table_end = len(out)
        tr_year_start = out.find('<tr', tr_title_end, table_end)
        tr_year_end = -1
        if tr_year_start != -1:
            tr_year_end = out.find('</tr>', tr_year_start, table_end)
            if tr_year_end != -1:
                tr_year_end += len('</tr>')
        # 若没有找到年份行，就在标题行后插入
        insert_after = tr_title_end
        replace_start = tr_year_start if tr_year_start != -1 else tr_title_end
        replace_end = tr_year_end if tr_year_end != -1 else tr_title_end

        # 生成新的年份行（降序；无年份则空行）
        year_cells = []
        for y in sorted(set(int(v) for v in years), reverse=True):
            year_cells.append(f'      <td><a href="#{anchor}_{y}">{y}</a></td>')
        new_year_row = "<tr>\n" + ("\n".join(year_cells)) + "\n</tr>"

        out = out[:replace_start] + new_year_row + out[replace_end:]

    return out


def generate_html_from_excel(excel_path, template_html_path, output_html_path, tz_name="Asia/Singapore"):
    """主函数：按分类与年份将 Excel 数据插入模板 HTML；并同步更新目录年份与更新时间戳。"""
    # 读取 Excel
    xls = pd.ExcelFile(excel_path)
    try:
        df = pd.read_excel(xls, sheet_name='Total')
    except Exception:
        # 如果没有 Total 表，则读取第一个表
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    df.columns = [str(c) for c in df.columns]

    # 按分类与年份分组，并在每年内排序
    grouped_entries = {}
    for raw_cat, anchor in CATEGORY_ANCHOR_MAP.items():
        df_cat = df[df['Type'].map(lambda x: s(x).lower()) == raw_cat]
        if df_cat.empty:
            continue

        df_cat = df_cat.copy()
        df_cat['Year'] = df_cat['Release Time'].map(to_year)
        df_cat = df_cat[df_cat['Year'].notna()]

        years = sorted(df_cat['Year'].unique().tolist(), reverse=True)
        entries_by_year = {}
        for year in years:
            entries = df_cat[df_cat['Year'] == year].to_dict('records')
            entries.sort(key=get_venue_sort_key)
            entries_by_year[int(year)] = entries

        grouped_entries[anchor] = {'years': years, 'entries_by_year': entries_by_year}

    # 读取模板 HTML
    template_html = Path(template_html_path).read_text(encoding='utf-8', errors='ignore')

    # 为每个分类构建新内容（保留原有分类标题；仅输出有论文的年份）
    new_sections = {}
    years_by_anchor = {}  # 用于重建 ToC
    for anchor, data in grouped_entries.items():
        years = [int(y) for y in data['years']]
        entries_by_year = data['entries_by_year']
        if not years:
            continue

        # 将模板中该分类的标题块（从 <a name="anchor" 到后续第一个 </h3>）原样保留
        start_anchor_idx = template_html.find(f'<a name="{anchor}"')
        if start_anchor_idx != -1:
            h3_start = template_html.find('<h3', start_anchor_idx)
            h3_end = template_html.find('</h3>', h3_start)
            if h3_start != -1 and h3_end != -1:
                header_block = template_html[start_anchor_idx: h3_end + 5]
            else:
                header_block = f'<a name="{anchor}"></a>\n<h3>{anchor}</h3>'
        else:
            header_block = f'<a name="{anchor}"></a>\n<h3>{anchor}</h3>'

        out_lines = [header_block]

        # ✅ 只输出有论文的年份
        valid_years = sorted(entries_by_year.keys(), reverse=True)
        for year in valid_years:
            out_lines.append(f'\n    <a name="{anchor}_{year}"></a>')
            out_lines.append(
                f'<h3 class="wp-block-heading" style="font-size: 24px;"><span style="color:#000000">'
                f' <i class=""></i> {year} </span></h3>'
            )
            for entry in entries_by_year[year]:
                out_lines.append(build_entry_html(entry, year))

        new_sections[anchor] = "\n".join(out_lines)
        years_by_anchor[anchor] = valid_years[:]  # 供目录同步

    # 将新内容写回模板：按分类锚点顺序替换每段内容
    output_html = template_html
    anchors_in_order = [a for a in [
        CATEGORY_ANCHOR_MAP['survey'],
        CATEGORY_ANCHOR_MAP['benchmarks'],
        CATEGORY_ANCHOR_MAP['classical'],
        CATEGORY_ANCHOR_MAP['augmented'],
        CATEGORY_ANCHOR_MAP['learning'],
    ] if a in output_html]

    for i, anchor in enumerate(anchors_in_order):
        if anchor not in new_sections:
            # 若该分类在数据里无条目，则跳过替换（保留原模板内容）
            continue
        start = output_html.find(f'<a name="{anchor}"')
        if start == -1:
            continue
        if i + 1 < len(anchors_in_order):
            next_anchor = anchors_in_order[i + 1]
            end = output_html.find(f'<a name="{next_anchor}"')
            if end == -1:
                end = len(output_html)
        else:
            # 最后一段：若存在 "Last updated" 作为尾界标则取其前，否则到文件末尾
            end = output_html.find('Last updated')
            if end == -1:
                end = len(output_html)
        output_html = output_html[:start] + new_sections.get(anchor, '') + output_html[end:]

    # ---- 同步重建 Table of Contents 的年份行（只显示有论文的年份）----
    output_html = rebuild_toc(output_html, years_by_anchor)

    # ---- 更新时间戳（Last updated: <Month DD, YYYY>）----
    if ZoneInfo and tz_name:
        now = datetime.now(ZoneInfo(tz_name))
    else:
        now = datetime.now()
    date_str = now.strftime("%B %d, %Y")

    if "Last updated" in output_html:
        # 替换第一处 Last updated 的日期
        output_html = re.sub(r'(Last updated:\s*)([^<\n]*)',
                             r'\1' + date_str,
                             output_html,
                             count=1)
    else:
        # 模板中没有该字样，则在文末追加
        output_html += f'\n<p style="text-align:center; color:#808080;">Last updated: {date_str}</p>\n'

    # 写出文件
    Path(output_html_path).write_text(output_html, encoding='utf-8')
    print(f"[OK] Wrote: {output_html_path}")


def main():
    parser = argparse.ArgumentParser(description="Generate HTML from Excel and insert into template (with ToC sync).")
    parser.add_argument("--excel", required=True, help="Path to Excel file (e.g., mapf汇总.xlsx)")
    parser.add_argument("--template", required=True, help="Path to template HTML (e.g., demo.html)")
    parser.add_argument("--output", required=True, help="Path to write updated HTML (e.g., index.html)")
    parser.add_argument("--tz", default="Asia/Singapore", help="Timezone for 'Last updated' (default: Asia/Singapore)")
    args = parser.parse_args()

    generate_html_from_excel(args.excel, args.template, args.output, tz_name=args.tz)


if __name__ == "__main__":
    main()