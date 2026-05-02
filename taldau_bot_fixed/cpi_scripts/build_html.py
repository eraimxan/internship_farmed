"""
cpi_scripts/build_html.py — Build interactive HTML visualization of CPI data.

Produces a single self-contained HTML file with:
  • Sidebar: oblast → district navigation (collapsible)
  • Main panel: searchable, sortable table with hierarchical rows
  • Toggle between the two comparison period types
  • Expand / collapse / level-filter buttons
  • Color-coded depth levels matching the Excel output

Usage:
    python build_html.py [--csv output/cpi_data.csv] [--out output/cpi_kazakhstan.html]
"""

import sys
import re
import json
import argparse
import logging
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from cpi_config import (
    DATA_CSV, HTML_OUT,
    TARGET_PERIODS, PERIOD_LABELS,
    ALL_REGIONS,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def _period_sort_key(key: str) -> int:
    m = re.match(r"^y(\d+)$", key)
    if not m:
        return 0
    return int(m.group(1))


def _period_label(key: str, period_name_map: dict) -> str:
    if key in period_name_map:
        return period_name_map[key]
    m = re.match(r"^y(\d+)$", key)
    if not m:
        return key
    s = m.group(1)
    if len(s) == 6:
        month, year = int(s[:2]), int(s[2:])
        q = (month - 1) // 3 + 1
        return f"Q{q} {year}"
    elif len(s) == 5:
        q, year = int(s[0]), int(s[1:])
        return f"Q{q} {year}"
    return key


def _esc(s) -> str:
    """Escape string for safe embedding in JavaScript string literals."""
    return (
        str(s)
        .replace("\\", "\\\\")
        .replace("'", "\\'")
        .replace("\n", " ")
        .replace("\r", "")
    )


def _dfs_order(rows: list[dict]) -> list[dict]:
    """
    Reorder rows so each parent is immediately followed by its children (DFS).
    Uses (parent_text, depth-1) → id lookup to wire parent_id, then walks the tree.
    """
    if not rows:
        return rows

    text_to_id: dict = {}
    for r in rows:
        text_to_id[(r["text"], r["depth"])] = r["id"]

    for r in rows:
        pt = r.get("parent_text", "")
        if pt and r["depth"] > 0:
            r["_pid"] = text_to_id.get((pt, r["depth"] - 1))
        else:
            r["_pid"] = None

    by_id: dict = {r["id"]: r for r in rows}
    children: dict = {}
    for r in rows:
        children.setdefault(r["_pid"], []).append(r["id"])

    all_ids = set(by_id.keys())
    roots = [rid for rid in all_ids if by_id[rid]["_pid"] not in all_ids]

    result: list = []
    visited: set = set()

    def dfs(rid):
        if rid in visited:
            return
        visited.add(rid)
        result.append(by_id[rid])
        for child in sorted(children.get(rid, []), key=lambda x: str(x)):
            dfs(child)

    for root in sorted(roots, key=lambda x: str(x)):
        dfs(root)
    for rid in all_ids - visited:
        result.append(by_id[rid])

    for r in result:
        r.pop("_pid", None)
    return result


# ─────────────────────────────────────────────
# DATA BUILDER
# ─────────────────────────────────────────────

def build_data_structure(df: pd.DataFrame, period_name_map: dict) -> dict:
    """
    Convert DataFrame into a JS-ready dict:
    {
      oblast_name: {
        region_name: {
          "yoy": [
            {id, text, depth, parent_text, vals: {period_key: value, ...}},
            ...
          ],
          "prev_year": [...]
        }
      }
    }
    Also returns sorted list of period keys.
    """
    period_cols = sorted(
        [c for c in df.columns if re.match(r"^y\d+$", c)],
        key=_period_sort_key,
    )

    comp_keys = [k for k in TARGET_PERIODS if k in df["comparison_type"].unique()]
    if not comp_keys:
        comp_keys = df["comparison_type"].dropna().unique().tolist()

    # Build nested structure
    # Support both "oblast" and "region" columns (CPI only has region=oblast)
    oblast_col = "oblast" if "oblast" in df.columns else "region"
    data: dict = {}
    for oblast in list(dict.fromkeys(df[oblast_col].dropna().tolist())):
        ob_df = df[df[oblast_col] == oblast]
        data[oblast] = {}
        for region in list(dict.fromkeys(ob_df["region"].dropna().tolist())):
            reg_df = ob_df[ob_df["region"] == region]
            data[oblast][region] = {}
            for ck in comp_keys:
                sub = reg_df[reg_df["comparison_type"] == ck]
                rows = []
                for _, row in sub.iterrows():
                    vals = {}
                    for pc in period_cols:
                        v = row.get(pc)
                        if v is not None and str(v).lower() not in ("", "nan", "none"):
                            try:
                                vals[pc] = round(float(v), 2)
                            except Exception:
                                vals[pc] = str(v)
                    rows.append({
                        "id":          str(row.get("id", "")),
                        "text":        str(row.get("text", "")).strip(),
                        "depth":       int(float(row.get("depth", 0))),
                        "parent_text": str(row.get("parent_text", "") or ""),
                        "vals":        vals,
                    })
                # DFS order: each parent immediately followed by its children (like the website tree)
                rows = _dfs_order(rows)
                data[oblast][region][ck] = rows

    return data, period_cols, comp_keys


# ─────────────────────────────────────────────
# HTML GENERATOR
# ─────────────────────────────────────────────

def build_html(df: pd.DataFrame, params_path: Path | None, out_path: Path) -> None:
    """Generate interactive HTML from CPI DataFrame."""

    # Load period name map
    period_name_map: dict[str, str] = {}
    if params_path and params_path.exists():
        try:
            p = json.loads(params_path.read_text(encoding="utf-8"))
            dl = p.get("date_list", [])
            nl = p.get("period_name_list", [])
            for k, n in zip(dl, nl):
                period_name_map[f"y{k}"] = n
        except Exception:
            pass

    data, period_cols, comp_keys = build_data_structure(df, period_name_map)

    period_labels_js = {
        pc: _esc(_period_label(pc, period_name_map))
        for pc in period_cols
    }

    # Count totals for subtitle
    total_oblasts = len(data)
    total_regions = sum(len(v) for v in data.values())
    is_single_level = all(len(v) == 1 and list(v.keys())[0] == ob for ob, v in data.items())

    comp_label_str = " / ".join(PERIOD_LABELS.get(k, k) for k in comp_keys)

    # Serialize data to JSON for embedding
    js_data = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    js_periods = json.dumps(period_cols, ensure_ascii=False)
    js_period_labels = json.dumps(period_labels_js, ensure_ascii=False)
    js_comp_keys = json.dumps(comp_keys, ensure_ascii=False)
    js_comp_labels = json.dumps({k: PERIOD_LABELS.get(k, k) for k in comp_keys}, ensure_ascii=False)

    # Sidebar structure (oblast → districts list)
    sidebar_oblasts = list(data.keys())
    # Map oblast → list of districts
    oblast_districts = {ob: list(data[ob].keys()) for ob in sidebar_oblasts}
    js_sidebar = json.dumps(
        [{"name": ob, "districts": oblast_districts[ob]} for ob in sidebar_oblasts],
        ensure_ascii=False,
    )
    # For CPI, data is at oblast level (each "district" IS the oblast itself)
    js_single_level = "true" if is_single_level else "false"

    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Индексы потребительских цен — Казахстан</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#1a1a2e;overflow:hidden;height:100vh}}
.hdr{{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:#fff;padding:10px 20px;display:flex;align-items:center;gap:16px;height:54px;flex-shrink:0}}
.hdr h1{{font-size:15px;font-weight:700;white-space:nowrap}}
.hdr-sub{{font-size:11px;opacity:.8}}
.search-box{{flex:1;max-width:340px}}
.search-box input{{width:100%;padding:6px 14px;border-radius:16px;border:none;font-size:13px;outline:none}}
.layout{{display:flex;height:calc(100vh - 54px)}}
.sidebar{{width:290px;background:#fff;border-right:1px solid #dde;overflow-y:auto;flex-shrink:0;font-size:13px}}
.sb-title{{padding:8px 14px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#888;background:#f8f9fa;border-bottom:1px solid #eee;position:sticky;top:0;z-index:5}}
.ob-hdr{{display:flex;align-items:center;padding:8px 14px;font-weight:600;color:#1F4E79;cursor:pointer;border-bottom:1px solid #f0f0f0;gap:6px}}
.ob-hdr:hover{{background:#EBF3FB}}
.ob-hdr.sel{{background:#1F4E79;color:#fff}}
.ob-arr{{font-size:9px;transition:transform .2s;flex-shrink:0}}
.ob-arr.open{{transform:rotate(90deg)}}
.ob-cnt{{margin-left:auto;font-size:10px;background:rgba(0,0,0,.12);border-radius:8px;padding:1px 7px;flex-shrink:0}}
.ob-hdr.sel .ob-cnt{{background:rgba(255,255,255,.25)}}
.dist-list{{display:none}}
.dist-list.open{{display:block}}
.dist-item{{padding:6px 14px 6px 30px;font-size:12px;cursor:pointer;color:#555;border-bottom:1px solid #f8f8f8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.dist-item:hover{{background:#f0f7ff}}
.dist-item.sel{{background:#2E75B6;color:#fff}}
.main{{flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0}}
.main-hdr{{padding:8px 16px;background:#fff;border-bottom:1px solid #dde;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap}}
.main-hdr h2{{font-size:14px;font-weight:700;color:#1F4E79;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.ob-name{{font-size:11px;color:#888;white-space:nowrap}}
.ctrl-btns{{display:flex;gap:6px;margin-left:auto;flex-shrink:0;flex-wrap:wrap}}
.btn{{padding:4px 10px;border:1px solid #ccc;border-radius:4px;background:#fff;cursor:pointer;font-size:11px;white-space:nowrap}}
.btn:hover{{background:#EBF3FB;border-color:#2E75B6;color:#2E75B6}}
.btn.pr{{background:#2E75B6;color:#fff;border-color:#2E75B6}}
.btn.pr:hover{{background:#1F4E79}}
.btn.act{{background:#1F4E79;color:#fff;border-color:#1F4E79}}
.comp-toggle{{display:flex;gap:0;border-radius:6px;overflow:hidden;border:1px solid #2E75B6}}
.comp-btn{{padding:4px 12px;border:none;cursor:pointer;font-size:11px;background:#fff;color:#2E75B6;transition:all .15s}}
.comp-btn:not(:last-child){{border-right:1px solid #2E75B6}}
.comp-btn.act{{background:#2E75B6;color:#fff}}
.tbl-wrap{{flex:1;overflow:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
thead tr th{{position:sticky;z-index:10;background:#1F4E79;color:#fff;padding:7px 8px;text-align:left;font-size:11px;white-space:nowrap;border-right:1px solid rgba(255,255,255,.15)}}
thead tr:first-child th{{top:0}}
thead tr:last-child th{{top:29px}}
th.num{{text-align:right}}
th.grp-yoy{{background:#2E75B6}}
th.grp-prev{{background:#14375E}}
th.per-yoy{{background:#4472C4}}
th.per-prev{{background:#1F3864}}
.lv0{{background:#1F4E79}}.lv0 td{{color:#fff!important;font-weight:700}}
.lv1{{background:#2E75B6}}.lv1 td{{color:#fff!important;font-weight:600}}
.lv2{{background:#BDD7EE;font-weight:600}}
.lv3{{background:#DDEBF7}}
.lv4{{background:#F2F8FD}}
.lv5{{background:#FAFCFF;color:#555}}
.lv6{{background:#fff;color:#666}}
tr.hid{{display:none}}
tr:hover td{{filter:brightness(.95)}}
td{{padding:5px 8px;border-bottom:1px solid #f0f0f0;border-right:1px solid #e8e8e8;white-space:nowrap}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
td.name-cell{{max-width:440px;overflow:hidden;text-overflow:ellipsis}}
.tog{{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;border-radius:3px;background:rgba(255,255,255,.25);font-size:9px;cursor:pointer;flex-shrink:0;margin-right:3px;vertical-align:middle}}
.lv2 .tog,.lv3 .tog,.lv4 .tog,.lv5 .tog,.lv6 .tog{{background:rgba(0,0,0,.1)}}
td.id-cell{{font-size:10px;color:#999;text-align:center;min-width:65px}}
.lv0 td.id-cell,.lv1 td.id-cell{{color:rgba(255,255,255,.6)}}
.empty{{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:#bbb;gap:10px;font-size:15px}}
.empty .big{{font-size:56px}}
</style>
</head>
<body>
<div class="hdr">
  <div>
    <h1>🇰🇿 Индексы потребительских цен — Казахстан</h1>
    <div class="hdr-sub">{comp_label_str} · {total_oblasts} областей · {total_regions} районов · данные как на сайте taldau.stat.gov.kz</div>
  </div>
  <div class="search-box"><input id="srch" placeholder="🔍 Поиск по показателю..." oninput="doSearch(this.value)"></div>
</div>
<div class="layout">
  <div class="sidebar">
    <div class="sb-title">Области и районы</div>
    <div id="sb"></div>
  </div>
  <div class="main">
    <div class="main-hdr">
      <div style="min-width:0">
        <div class="ob-name" id="obName"></div>
        <h2 id="distName">Выберите район слева</h2>
      </div>
      <div class="ctrl-btns">
        <div class="comp-toggle" id="compToggle"></div>
        <button class="btn" onclick="setLevel(1)">L1</button>
        <button class="btn" onclick="setLevel(2)">L2</button>
        <button class="btn" onclick="setLevel(3)">L3</button>
        <button class="btn" onclick="colAll()">Свернуть</button>
        <button class="btn pr" onclick="expAll()">Развернуть всё</button>
      </div>
    </div>
    <div class="tbl-wrap" id="tblWrap">
      <div class="empty"><div class="big">📊</div><div>Выберите район в левой панели</div></div>
    </div>
  </div>
</div>
<script>
const DATA=%%DATA%%;
const PERIODS=%%PERIODS%%;
const PERIOD_LABELS=%%PERIOD_LABELS%%;
const COMP_KEYS=%%COMP_KEYS%%;
const COMP_LABELS=%%COMP_LABELS%%;
const SIDEBAR=%%SIDEBAR%%;
const SINGLE_LEVEL=%%SINGLE_LEVEL%%; // CPI: each sidebar item IS the oblast (no sub-regions)

let curOblast=null, curDistrict=null, curComp=COMP_KEYS[0]||'yoy';
let searchStr='';

// ── Build sidebar ────────────────────────────
function buildSidebar(){{
  const sb=document.getElementById('sb');
  sb.innerHTML='';
  SIDEBAR.forEach(ob=>{{
    if(SINGLE_LEVEL){{
      // CPI: oblast IS the data unit — single clickable item
      const hdr=document.createElement('div');
      hdr.className='ob-hdr';
      hdr.id='obhdr-'+escId(ob.name);
      hdr.innerHTML=`<span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis">${{ob.name}}</span>`;
      hdr.onclick=()=>selectDistrict(ob.name,ob.name);
      sb.appendChild(hdr);
    }} else {{
      const hdr=document.createElement('div');
      hdr.className='ob-hdr';
      hdr.innerHTML=`<span class="ob-arr" id="arr-${{escId(ob.name)}}">▶</span>
        <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis">${{ob.name}}</span>
        <span class="ob-cnt">${{ob.districts.length}}</span>`;
      hdr.onclick=()=>toggleOblast(ob.name, hdr);
      sb.appendChild(hdr);
      const dl=document.createElement('div');
      dl.className='dist-list';
      dl.id='dl-'+escId(ob.name);
      ob.districts.forEach(d=>{{
        const di=document.createElement('div');
        di.className='dist-item';
        di.textContent=d;
        di.id='di-'+escId(ob.name)+'-'+escId(d);
        di.onclick=()=>selectDistrict(ob.name,d);
        dl.appendChild(di);
      }});
      sb.appendChild(dl);
    }}
  }});
}}

function escId(s){{return s.replace(/[^a-zA-Z0-9]/g,'_');}}

function toggleOblast(name, hdrEl){{
  const dl=document.getElementById('dl-'+escId(name));
  if(!dl)return;
  const isOpen=dl.classList.contains('open');
  const arr=document.getElementById('arr-'+escId(name));
  if(isOpen){{dl.classList.remove('open');if(arr)arr.classList.remove('open');}}
  else{{dl.classList.add('open');if(arr)arr.classList.add('open');}}
}}

function selectDistrict(oblast, dist){{
  document.querySelectorAll('.dist-item.sel,.ob-hdr.sel').forEach(el=>el.classList.remove('sel'));
  if(SINGLE_LEVEL){{
    const el=document.getElementById('obhdr-'+escId(oblast));
    if(el)el.classList.add('sel');
    document.getElementById('obName').textContent='';
  }} else {{
    const di=document.getElementById('di-'+escId(oblast)+'-'+escId(dist));
    if(di)di.classList.add('sel');
    document.getElementById('obName').textContent=oblast;
  }}
  curOblast=oblast; curDistrict=dist;
  document.getElementById('distName').textContent=dist;
  renderTable();
}}

// ── Comparison toggle ────────────────────────
function buildCompToggle(){{
  const ct=document.getElementById('compToggle');
  ct.innerHTML='';
  COMP_KEYS.forEach(k=>{{
    const b=document.createElement('button');
    b.className='comp-btn'+(k===curComp?' act':'');
    b.textContent=COMP_LABELS[k]||k;
    b.onclick=()=>{{curComp=k;buildCompToggle();renderTable();}};
    ct.appendChild(b);
  }});
}}

// ── Render table ─────────────────────────────
function renderTable(){{
  if(!curOblast||!curDistrict)return;
  const rows=(DATA[curOblast]||{{}})[curDistrict]||{{}};
  const data=rows[curComp]||[];

  const wrap=document.getElementById('tblWrap');
  if(!data.length){{
    wrap.innerHTML='<div class="empty"><div class="big">🔍</div><div>Нет данных</div></div>';
    return;
  }}

  // Determine which periods have any value
  const livePeriods=PERIODS.filter(p=>data.some(r=>r.vals[p]!=null));

  let html='<table><thead>';

  // Header row 1: "Показатель" | period labels
  html+='<tr>';
  html+='<th rowspan="1" style="min-width:40px">ID</th>';
  html+='<th rowspan="1" style="min-width:380px">Наименование</th>';
  livePeriods.forEach(p=>{{
    html+=`<th class="num">${{PERIOD_LABELS[p]||p}}</th>`;
  }});
  html+='</tr></thead><tbody id="tbl-body">';

  // Parent tracking for expand/collapse
  let togCols='';
  data.forEach((r,i)=>{{
    const cls='lv'+Math.min(r.depth,6);
    const hasChildren=data.some(c=>c.parent_text===r.text&&c.depth===r.depth+1);
    const togHtml=hasChildren
      ? `<span class="tog" onclick="togRow(${{i}},event)">−</span>`
      : '<span style="display:inline-block;width:17px"></span>';
    const indent='&nbsp;'.repeat(r.depth*4);
    const vals=livePeriods.map(p=>{{
      const v=r.vals[p];
      if(v===null||v===undefined) return '<td class="num">—</td>';
      if(v==='x') return '<td class="num" style="color:#999;font-style:italic">x</td>';
      return `<td class="num">${{(+v).toFixed(2)}}</td>`;
    }}).join('');
    html+=`<tr class="${{cls}}" data-depth="${{r.depth}}" data-idx="${{i}}" data-parent="${{esc(r.parent_text)}}">
      <td class="id-cell">${{r.id||''}}</td>
      <td class="name-cell">${{indent}}${{togHtml}}${{esc(r.text)}}</td>
      ${{vals}}
    </tr>`;
  }});

  html+='</tbody></table>';
  wrap.innerHTML=html;
  applySearch(searchStr);
}}

function esc(s){{
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}}

// ── Expand / Collapse ────────────────────────
function togRow(idx,e){{
  e&&e.stopPropagation();
  const tbody=document.getElementById('tbl-body');
  if(!tbody)return;
  const allRows=[...tbody.querySelectorAll('tr')];
  const baseRow=allRows.find(r=>+r.dataset.idx===idx);
  if(!baseRow)return;
  const baseDepth=+baseRow.dataset.depth;
  const tog=baseRow.querySelector('.tog');
  const collapsed=tog&&tog.textContent==='+';

  // Find all descendants
  let inBlock=false;
  allRows.forEach(r=>{{
    if(+r.dataset.idx===idx){{inBlock=true;return;}}
    if(!inBlock)return;
    const d=+r.dataset.depth;
    if(d<=baseDepth){{inBlock=false;return;}}
    if(collapsed){{
      // show only direct children; deeper are managed by their own state
      if(d===baseDepth+1) r.classList.remove('hid');
    }}else{{
      r.classList.add('hid');
    }}
  }});
  if(tog) tog.textContent=collapsed?'−':'+';
}}

function expAll(){{
  document.querySelectorAll('#tbl-body tr').forEach(r=>r.classList.remove('hid'));
  document.querySelectorAll('.tog').forEach(t=>t.textContent='−');
}}

function colAll(){{
  const tbody=document.getElementById('tbl-body');
  if(!tbody)return;
  [...tbody.querySelectorAll('tr')].forEach(r=>{{
    if(+r.dataset.depth>0)r.classList.add('hid');
  }});
  document.querySelectorAll('.tog').forEach(t=>t.textContent='+');
}}

function setLevel(maxLv){{
  const tbody=document.getElementById('tbl-body');
  if(!tbody)return;
  [...tbody.querySelectorAll('tr')].forEach(r=>{{
    const d=+r.dataset.depth;
    r.classList.toggle('hid',d>=maxLv);
  }});
  document.querySelectorAll('.tog').forEach(t=>{{
    const row=t.closest('tr');
    if(row){{
      const d=+row.dataset.depth;
      t.textContent=(d<maxLv-1)?'−':'+';
    }}
  }});
}}

// ── Search ───────────────────────────────────
function doSearch(val){{
  searchStr=val.toLowerCase().trim();
  applySearch(searchStr);
}}

function applySearch(q){{
  const tbody=document.getElementById('tbl-body');
  if(!tbody)return;
  [...tbody.querySelectorAll('tr')].forEach(r=>{{
    if(!q){{r.style.display='';return;}}
    const txt=r.querySelector('.name-cell')?.textContent.toLowerCase()||'';
    r.style.display=txt.includes(q)?'':'none';
  }});
}}

// ── Init ─────────────────────────────────────
buildSidebar();
buildCompToggle();
</script>
</body>
</html>"""

    # Inject variables into script
    html = html.replace("%%DATA%%", js_data)
    html = html.replace("%%PERIODS%%", js_periods)
    html = html.replace("%%PERIOD_LABELS%%", js_period_labels)
    html = html.replace("%%COMP_KEYS%%", js_comp_keys)
    html = html.replace("%%COMP_LABELS%%", js_comp_labels)
    html = html.replace("%%SIDEBAR%%", js_sidebar)
    html = html.replace("%%SINGLE_LEVEL%%", js_single_level)

    out_path.write_text(html, encoding="utf-8")
    size_mb = out_path.stat().st_size / (1024 * 1024)
    log.info("HTML saved: %s  (%.1f MB)", out_path, size_mb)


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Build CPI HTML from CSV")
    parser.add_argument("--csv",    default=str(DATA_CSV),  help="Input CSV file")
    parser.add_argument("--out",    default=str(HTML_OUT),  help="Output HTML file")
    parser.add_argument("--params", default="",             help="cpi_params.json path")
    args = parser.parse_args()

    csv_path = Path(args.csv)
    out_path = Path(args.out)

    if not csv_path.exists():
        log.error("CSV not found: %s — run fetch.py first", csv_path)
        sys.exit(1)

    df = pd.read_csv(csv_path, encoding="utf-8-sig", low_memory=False)
    log.info("Loaded %d rows from %s", len(df), csv_path)

    params_path = Path(args.params) if args.params else csv_path.parent / "cpi_params.json"
    build_html(df, params_path if params_path.exists() else None, out_path)


if __name__ == "__main__":
    main()
