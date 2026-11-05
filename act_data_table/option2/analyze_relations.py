# -*- coding: utf-8 -*-
"""테이블 관계 후보 분석 + 서버/인터넷 없이 여는 관계도.

기본 입력: 이 파일 위치/targets
기본 출력: 이 파일 위치/targets_relation_analysis
다른 폴더: python analyze_relations.py --input targets --scope all

지원: CSV, XLSX의 모든 시트. MD/TXT는 참고 문서로만 표시.
단일 컬럼의 이름·값·고유성을 비교하며 업무상 관계를 확정하지 않습니다.
"""
from __future__ import annotations

from time import perf_counter, time

_SCRIPT_STARTED = perf_counter()
_SCRIPT_STARTED_TIMESTAMP = time()

import argparse
import csv
import json
import math
import re
import traceback
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_MAX_ROWS = 5
SAMPLE_MAX_COLUMNS = 10
SAMPLE_MAX_CHARS = 100


def normalize_name(value):
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", str(value)).casefold())


def key_hint(name):
    value = normalize_name(name)
    return bool(re.search(r"(id|code|key)$", value)) or any(
        token in value for token in ("코드", "번호", "사번", "학번", "식별", "키")
    )


def name_similarity(left, right):
    def stem(value):
        return re.sub(r"(id|code|key|코드|번호|키)$", "", normalize_name(value))
    return SequenceMatcher(None, stem(left), stem(right)).ratio()


def text_value(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (datetime, date)):
        result = value.isoformat()
    elif isinstance(value, float) and math.isfinite(value) and value.is_integer():
        result = str(int(value))
    else:
        result = str(value)
    result = unicodedata.normalize("NFKC", result).strip()
    return result if result else None


def numeric_text(value):
    if re.match(r"^[+-]?0\d+$", value):
        return False  # 001 같은 식별자의 선행 0 보존
    try:
        return Decimal(value).is_finite()
    except InvalidOperation:
        return False


def canonical(value, kind):
    if value is None:
        return None
    if kind == "숫자":
        return "N:" + str(Decimal(value).normalize())
    return "T:" + value  # 대소문자는 합치지 않음


def read_text(path):
    raw = path.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    raise ValueError("UTF-8/CP949/UTF-16으로 읽지 못했습니다.")


def read_tables(path):
    if path.suffix.lower() == ".csv":
        # NA 같은 실제 문자열과 선행 0을 임의로 결측/숫자로 바꾸지 않음.
        from io import StringIO
        yield "CSV", pd.read_csv(StringIO(read_text(path)), dtype=str, keep_default_na=False)
    else:
        with pd.ExcelFile(path, engine="openpyxl") as book:
            for name in book.sheet_names:
                yield name, pd.read_excel(book, sheet_name=name, dtype=object, keep_default_na=False)


def scope_for(path, root, mode):
    if mode == "all":
        return "전체"
    cursor = path.parent
    while cursor.is_relative_to(root):
        if (cursor / "README.md").is_file():
            return cursor.relative_to(root).as_posix() or "."
        if cursor == root:
            break
        cursor = cursor.parent
    return path.parent.relative_to(root).as_posix() or "."


def build_profile(table, series, column, index):
    raw_values = [text_value(value) for value in series]
    nonempty = [value for value in raw_values if value is not None]
    distinct = set(nonempty)
    kind = "숫자" if nonempty and all(numeric_text(v) for v in nonempty) else "문자열"
    values = [canonical(value, kind) for value in raw_values]
    values_set = {value for value in values if value is not None}
    counts = Counter(value for value in values if value is not None)
    ratio = len(values_set) / len(nonempty) if nonempty else 0
    hinted = key_hint(column)
    eligible = bool(values_set) and (
        hinted or (kind != "숫자" and len(values_set) >= 2 and ratio >= 0.5)
    )
    return {
        "id": table["id"] + "_C" + str(index + 1),
        "table_id": table["id"], "file": table["file"], "sheet": table["sheet"],
        "scope": table["scope"], "column": str(column), "column_index": index + 1,
        "dtype": str(series.dtype), "kind": kind, "rows": len(series),
        "nonmissing": len(nonempty), "missing": len(series) - len(nonempty),
        "distinct": len(values_set), "unique_pct": round(ratio * 100, 1),
        "missing_pct": round((len(series) - len(nonempty)) / len(series) * 100, 1) if len(series) else 0,
        "key_hint": hinted, "eligible": eligible,
        "examples": list(dict.fromkeys(nonempty))[:5],
        "_values": values, "_set": values_set, "_counts": counts,
        "_display": {canonical(value, kind): value for value in nonempty},
    }


def build_sample(frame):
    """미리보기만 제한하며, 관계 추정에는 전체 원본 행을 사용합니다."""
    shortened = 0

    def display(value):
        nonlocal shortened
        if value is None or pd.isna(value):
            return ""
        text = value.isoformat() if isinstance(value, (datetime, date)) else str(value)
        if len(text) > SAMPLE_MAX_CHARS:
            shortened += 1
            return text[:SAMPLE_MAX_CHARS - 1] + "…"
        return text

    preview = frame.iloc[:SAMPLE_MAX_ROWS, :SAMPLE_MAX_COLUMNS]
    columns = [display(col) for col in preview.columns]
    rows = [[display(value) for value in row]
            for row in preview.itertuples(index=False, name=None)]
    return {"columns": columns, "rows": rows, "shortened_values": shortened}


def public_profile(profile):
    return {key: value for key, value in profile.items() if not key.startswith("_")}


def direction(source, target, shared):
    matches = sum(source["_counts"][v] for v in shared)
    coverage = matches / source["nonmissing"] if source["nonmissing"] else 0
    target_unique = target["distinct"] == target["nonmissing"] and target["nonmissing"] > 0
    # 포함률과 참조 대상의 고유성·값 범위를 함께 비교합니다.
    rank = (target_unique, coverage, target["unique_pct"], target["distinct"])
    return rank, matches, coverage


def infer_relations(profiles, tables, min_coverage):
    by_id = {p["id"]: p for p in profiles}
    table_by_id = {t["id"]: t for t in tables}
    index = defaultdict(list)
    for profile in profiles:
        if profile["eligible"]:
            for value in profile["_set"]:
                index[(profile["scope"], value)].append(profile["id"])
    pairs = set()
    for columns in index.values():
        for i, left in enumerate(columns):
            for right in columns[i + 1:]:
                if by_id[left]["table_id"] != by_id[right]["table_id"]:
                    pairs.add(tuple(sorted((left, right))))
    relations = []
    for left_id, right_id in sorted(pairs):
        left, right = by_id[left_id], by_id[right_id]
        if left["kind"] != right["kind"]:
            continue
        shared = left["_set"] & right["_set"]
        similarity = name_similarity(left["column"], right["column"])
        exact_name = normalize_name(left["column"]) == normalize_name(right["column"])
        rank_left, _, _ = direction(left, right, shared)
        rank_right, _, _ = direction(right, left, shared)
        source, target = (left, right) if rank_left >= rank_right else (right, left)
        _, matched_rows, coverage = direction(source, target, shared)
        if coverage < min_coverage:
            continue
        if not (exact_name or similarity >= 0.35 or (len(shared) >= 3 and coverage >= 0.8)):
            continue
        source_unique = source["distinct"] == source["nonmissing"]
        target_unique = target["distinct"] == target["nonmissing"]
        cardinality = ("1" if source_unique else "N") + ":" + ("1" if target_unique else "N") + " 후보"
        caution = []
        if len(shared) < 3 or min(source["nonmissing"], target["nonmissing"]) < 3:
            caution.append("표본 작음: 값 1~2개의 일치는 관계를 확정하기에 부족")
        if coverage < 1:
            caution.append("참조 대상에서 찾지 못한 값 존재")
        versioned = table_by_id[source["table_id"]]["versioned"] or table_by_id[target["table_id"]]["versioned"]
        if versioned:
            caution.append("버전 혼재: 최종 유효본 필터 확인 필요")
        if not target_unique:
            caution.append("참조 대상 컬럼이 고유하지 않음")
        unmatched = sorted(source["_set"] - target["_set"])
        relations.append({
            "id": "R" + str(len(relations) + 1),
            "from": source["table_id"], "to": target["table_id"],
            "from_column": source["column"], "to_column": target["column"],
            "from_file": source["file"], "to_file": target["file"],
            "from_sheet": source["sheet"], "to_sheet": target["sheet"],
            "scope": source["scope"],
            "from_rows": source["rows"], "to_rows": target["rows"],
            "from_nonmissing": source["nonmissing"], "to_nonmissing": target["nonmissing"],
            "from_distinct": source["distinct"], "to_distinct": target["distinct"],
            "from_unique_pct": source["unique_pct"], "to_unique_pct": target["unique_pct"],
            "from_missing": source["missing"], "to_missing": target["missing"],
            "matched_distinct": len(shared), "matched_rows": matched_rows,
            "row_match_pct": round(coverage * 100, 1),
            "distinct_match_pct": round(len(shared) / source["distinct"] * 100, 1),
            "reverse_distinct_pct": round(len(shared) / target["distinct"] * 100, 1),
            "unmatched_rows": source["nonmissing"] - matched_rows,
            "unmatched_distinct": len(unmatched),
            "matched_examples": [source["_display"][v] for v in sorted(shared)[:10]],
            "unmatched_examples": [source["_display"][v] for v in unmatched[:10]],
            "name_similarity_pct": round(similarity * 100, 1),
            "cardinality": cardinality, "versioned": versioned,
            "status": "포함 관계 후보" if coverage == 1 and target_unique else "부분 일치 후보",
            "cautions": caution,
        })
    return relations


def write_csv(path, records, columns):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow([label for _, label in columns])
        for record in records:
            values = []
            for key, _ in columns:
                value = record.get(key, "")
                if isinstance(value, (list, dict)):
                    value = json.dumps(value, ensure_ascii=False)
                if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
                    value = "'" + value
                values.append(value)
            writer.writerow(values)


HTML = r"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>테이블 관계 후보</title><link rel="icon" href="data:,">
<style>
:root{font-family:"Malgun Gothic",system-ui,sans-serif;color:#26364a;background:#f3f6f8;color-scheme:light}*{box-sizing:border-box}body{margin:0}button,input,select{font:inherit}button,select{border:1px solid #d7e2e8;border-radius:7px;padding:8px 12px;background:white;color:inherit}button{cursor:pointer}button:hover{background:#eaf7f2}button:focus-visible,input:focus-visible,select:focus-visible{outline:3px solid #12a58c;outline-offset:2px}.layout{display:grid;grid-template-columns:235px minmax(0,1fr);min-height:100vh}aside{background:white;border-right:1px solid #dde6eb;padding:25px 16px;overflow:auto}h1{font-size:22px;margin:0 0 10px}h2{font-size:17px;margin:0 0 14px}.sub{color:#718096;font-size:12px;line-height:1.8}.table-choice{display:flex;gap:8px;padding:12px 5px;border-bottom:1px solid #edf1f4;font-size:13px;line-height:1.7}.table-choice small{display:block;color:#7c8b9c;font-size:11px}.main{padding:26px;min-width:0}.heading{display:flex;align-items:center;gap:20px;justify-content:space-between;flex-wrap:wrap}.numbers{font-size:13px;color:#52667c}.numbers strong{font-size:21px;color:#087e6e;margin:0 4px 0 12px}.notice{margin:18px 0;background:#fff3dd;border:1px solid #f1dfb6;color:#795e23;padding:13px 16px;border-radius:8px;font-size:12px;line-height:1.9}.toolbar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin:18px 0;font-size:12px}.workspace{display:grid;grid-template-columns:minmax(0,1fr) 340px;gap:18px}.card{background:white;border:1px solid #dce5eb;border-radius:12px;overflow:hidden}.graph-head{padding:15px 18px;display:flex;justify-content:space-between;gap:12px;font-size:12px;color:#64748b;border-bottom:1px solid #edf1f4}.graph{width:100%;height:610px;display:block;touch-action:none;background:radial-gradient(#dfe7ed 1px,transparent 1px);background-size:18px 18px}.graph .node{cursor:grab}.graph text{user-select:none}.detail{padding:20px;overflow:auto;max-height:710px}.detail p{font-size:13px;line-height:1.9}.detail .path{font-size:11px;color:#718096;overflow-wrap:anywhere}.tag{display:inline-block;padding:3px 8px;background:#e6f4ee;color:#087e6e;border-radius:4px;font-size:11px}.warn{background:#fff0d6;color:#98611a}.metrics{width:100%;border-collapse:collapse;font-size:12px;line-height:1.8}.metrics td,.metrics th{padding:8px 5px;text-align:left;border-bottom:1px solid #edf1f4;vertical-align:top}.metrics th{color:#718096;white-space:nowrap}.relations{display:flex;gap:8px;flex-wrap:wrap;padding:16px;border-top:1px solid #edf1f4}.relations button{font-size:11px;text-align:left}.relations button.active{border-color:#0b8c78;background:#e9f7f0}.relations small{display:block;color:#64748b;margin-top:4px}.section{margin-top:20px}.doc{padding:20px}.doc pre{white-space:pre-wrap;overflow-wrap:anywhere;font:12px/1.9 "Malgun Gothic",sans-serif;color:#4c6076}.rules{font-size:12px;line-height:1.9;color:#64748b;padding:18px}.rules ul{padding-left:18px}.empty{padding:35px;color:#718096;font-size:13px}.legend{font-size:11px;color:#64748b;line-height:1.8;padding:0 18px 15px}.dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px;background:#168977}.orange{background:#d4942c}.foot{font-size:11px;color:#718096;line-height:1.9;margin-top:16px}@media(max-width:1280px){.workspace{grid-template-columns:1fr}.detail{max-height:none}}@media(max-width:700px){.layout{display:block}aside{border-bottom:1px solid #dce5eb;max-height:300px}.main{padding:18px 12px}.graph{height:490px}}
.sample-note{margin:18px 0 10px;padding:12px;background:#fff4df;border:2px solid #d99820;border-radius:8px;color:#745019;font-size:12px;line-height:1.8}.sample-note strong{display:block;font-size:15px;color:#9a5100}.sample-scroll{overflow:auto;max-height:290px;border:1px solid #e2e8ee;border-radius:6px}.sample-grid{border-collapse:collapse;width:max-content;font-size:12px}.sample-grid th,.sample-grid td{padding:8px 10px;border:1px solid #e2e8ee;text-align:left;vertical-align:top;min-width:75px;max-width:230px;white-space:pre-wrap;overflow-wrap:anywhere}.sample-grid th{background:#f1f5f8;position:sticky;top:0;z-index:1}.sample-grid .row-number{min-width:38px;color:#718096;background:#f7f9fb}.profile-heading{font-size:14px;margin:20px 0 10px}
</style></head><body>
<div class="layout"><aside><h2>테이블 목록</h2><p class="sub">체크로 표시할 테이블을 선택하고,<br>관계선 또는 테이블을 클릭하세요.</p><div id="tables"></div><p class="sub" id="source"></p></aside>
<main class="main"><header class="heading"><div><h1>테이블 관계 후보</h1><div class="sub" id="scope"></div></div><div class="numbers" id="numbers"></div></header>
<div class="notice" id="notice"></div>
<div class="toolbar"><label>최소 원시 행 일치율 <select id="coverage"><option value="50">50%</option><option value="80">80%</option><option value="100">100%</option></select></label><label><input id="versions" type="checkbox" checked> 버전 혼재 후보 포함</label><button id="fit">배치 초기화</button></div>
<div class="workspace"><section class="card"><div class="graph-head"><span>화살표: 참조하는 쪽 → 참조 대상 후보</span><span>드래그: 이동 · 휠: 확대/축소</span></div><svg id="graph" class="graph" viewBox="0 0 1100 650" role="img" aria-label="테이블 관계 후보 그래프"></svg><div class="legend"><span class="dot"></span>값 포함 100% 후보　<span class="dot orange"></span>일부 일치 후보　점선: 버전 혼재<br>선의 숫자는 원시 비결측 행의 값 일치율입니다. 업무상 관계의 확률이 아닙니다.</div><div id="relations" class="relations"></div></section>
<section class="card detail" id="detail" aria-live="polite"></section></div>
<details class="card rules section"><summary>탐색 기준과 한계</summary><ul id="rules"></ul></details>
<section class="section" id="documents"></section><div class="foot" id="foot"></div>
</main></div>
<script id="data" type="application/json">__DATA__</script>
<script>
'use strict';
const data=JSON.parse(document.getElementById('data').textContent),$=id=>document.getElementById(id),NS='http://www.w3.org/2000/svg';
const byId=new Map(data.tables.map(t=>[t.id,t])),visible=new Set(byId.keys());let selected=null,positions={},pan={x:0,y:0,k:1},drag=null;
const fmt=n=>Number(n).toLocaleString('ko-KR'),pct=n=>Number(n).toFixed(1)+'%';
function el(tag,text,cls){const e=document.createElement(tag);if(text!==undefined)e.textContent=text;if(cls)e.className=cls;return e;}
function svg(tag,attrs){const e=document.createElementNS(NS,tag);for(const [key,value] of Object.entries(attrs||{}))e.setAttribute(key,String(value));return e;}
function activeEdges(){return data.relations.filter(r=>visible.has(r.from)&&visible.has(r.to)&&r.row_match_pct>=Number($('coverage').value)&&($('versions').checked||!r.versioned));}
function initialPositions(){
 const degrees={};data.tables.forEach(t=>degrees[t.id]=0);data.relations.forEach(r=>{degrees[r.from]++;degrees[r.to]++;});
 const sorted=[...data.tables].sort((a,b)=>degrees[b.id]-degrees[a.id]||a.id.localeCompare(b.id));
 positions={};if(!sorted.length)return;
 const hub=degrees[sorted[0].id]>=2&&sorted.length>2?sorted.shift():null;if(hub)positions[hub.id]={x:550,y:325};
 sorted.forEach((t,i)=>{const a=-Math.PI/2+i*Math.PI*2/Math.max(1,sorted.length);positions[t.id]=sorted.length===1&&!hub?{x:550,y:325}:{x:550+285*Math.cos(a),y:325+235*Math.sin(a)};});
 pan={x:0,y:0,k:1};
}
function endpoint(p,q){const dx=q.x-p.x,dy=q.y-p.y,s=Math.min(109/(Math.abs(dx)||0.0001),43/(Math.abs(dy)||0.0001));return{x:p.x+dx*s,y:p.y+dy*s};}
function paint(){
 const root=$('graph');root.replaceChildren();const defs=svg('defs'),marker=svg('marker',{id:'arrow',viewBox:'0 0 10 10',refX:9,refY:5,markerWidth:7,markerHeight:7,orient:'auto-start-reverse'});marker.append(svg('path',{d:'M0 0 L10 5 L0 10Z',fill:'#668494'}));defs.append(marker);root.append(defs);
 const group=svg('g',{transform:'translate('+pan.x+' '+pan.y+') scale('+pan.k+')'});root.append(group);
 for(const r of activeEdges()){
  const p=endpoint(positions[r.from],positions[r.to]),q=endpoint(positions[r.to],positions[r.from]),active=selected?.type==='edge'&&selected.id===r.id;
  const g=svg('g',{'data-edge':r.id,tabindex:0,role:'button','aria-label':byId.get(r.from).label+' '+r.from_column+' → '+byId.get(r.to).label+' '+r.to_column});
  const line=svg('line',{x1:p.x,y1:p.y,x2:q.x,y2:q.y,stroke:r.row_match_pct===100?'#168977':'#d4942c','stroke-width':active?4:2.5,'marker-end':'url(#arrow)'});if(r.versioned)line.setAttribute('stroke-dasharray','7 5');
  g.append(svg('line',{x1:p.x,y1:p.y,x2:q.x,y2:q.y,stroke:'transparent','stroke-width':18}),line);
  const text=svg('text',{x:(p.x+q.x)/2,y:(p.y+q.y)/2-9,'text-anchor':'middle',fill:'#435b70','font-size':15,'paint-order':'stroke',stroke:'white','stroke-width':6,'stroke-linejoin':'round'});text.textContent=pct(r.row_match_pct);g.append(text);const title=svg('title');title.textContent=r.from_column+' → '+r.to_column+' / '+r.cardinality;g.append(title);
  g.addEventListener('click',()=>showEdge(r));g.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();showEdge(r);}});group.append(g);
 }
 for(const t of data.tables){if(!visible.has(t.id))continue;const p=positions[t.id],g=svg('g',{class:'node',transform:'translate('+(p.x-109)+' '+(p.y-43)+')','data-node':t.id,tabindex:0,role:'button','aria-label':t.label}),active=selected?.type==='table'&&selected.id===t.id;
  g.append(svg('rect',{width:218,height:86,rx:12,fill:active?'#e6f6ee':'white',stroke:active?'#168977':'#bdcfd9','stroke-width':active?3:1.5}));
  const title=svg('text',{x:16,y:31,fill:'#20374b','font-size':18,'font-weight':700});title.textContent=t.label.length>15?t.label.slice(0,14)+'…':t.label;g.append(title);
  const sub=svg('text',{x:16,y:56,fill:'#6a7d90','font-size':13});sub.textContent=t.rows+'행 · '+t.columns+'열'+(t.versioned?' · 버전 혼재':'');g.append(sub);
  const hint=svg('title');hint.textContent=t.file+' / '+t.sheet;g.append(hint);
  g.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();showTable(t);}});group.append(g);
 }
 if(!data.tables.length){const text=svg('text',{x:550,y:325,'text-anchor':'middle',fill:'#718096'});text.textContent='읽을 수 있는 테이블이 없습니다.';root.append(text);}
 renderRelationButtons();
}
function rowsTable(rows){const table=el('table',undefined,'metrics');for(const [key,value] of rows){const tr=el('tr');tr.append(el('th',key),el('td',String(value)));table.append(tr);}return table;}
function showEdge(r){
 selected={type:'edge',id:r.id};const panel=$('detail');panel.replaceChildren();panel.append(el('h2',byId.get(r.from).label+' → '+byId.get(r.to).label));
 panel.append(el('span',r.status,'tag'+(r.versioned||r.row_match_pct<100?' warn':'')));
 panel.append(el('p',r.from_column+' → '+r.to_column+' · '+r.cardinality));
 panel.append(rowsTable([['원시 행 일치율',pct(r.row_match_pct)+' ('+r.matched_rows+'/'+r.from_nonmissing+'행)'],['고유값 일치율',pct(r.distinct_match_pct)],['반대쪽 고유값 포함률',pct(r.reverse_distinct_pct)],['참조 대상 고유성',pct(r.to_unique_pct)],['같은 값 개수',r.matched_distinct+'개'],['미일치 행',r.unmatched_rows+'행'],['일치 예시',r.matched_examples.join(', ')||'없음'],['미일치 예시',r.unmatched_examples.join(', ')||'없음']]));
 for(const caution of r.cautions)panel.append(el('p','⚠ '+caution));
 panel.append(el('p',r.from_file+' ['+r.from_sheet+'] → '+r.to_file+' ['+r.to_sheet+']','path'));paint();
}
function appendSample(panel,t){
 const sample=t.sample,note=el('div',undefined,'sample-note');
 note.append(el('strong','원본 데이터 샘플 · 전체 데이터 아님'));
 const rows=sample?.rows.length||0,cols=sample?.columns.length||0;
 note.append(el('div','전체 '+fmt(t.rows)+'행 × '+fmt(t.columns)+'열 중 상위 '+rows+'행 · 앞 '+cols+'열 표시'));
 note.append(el('div','최대 5행·10열, 셀당 100자. 관계 분석에는 전체 데이터를 사용합니다.'));
 if(sample?.shortened_values)note.append(el('div','긴 값·컬럼명 '+sample.shortened_values+'개를 줄여 표시했습니다. (… 표시)'));
 panel.append(note);
 if(!rows||!cols){panel.append(el('p','표시할 데이터 행이 없습니다.'));return;}
 const wrap=el('div',undefined,'sample-scroll'),table=el('table',undefined,'sample-grid'),thead=el('thead'),head=el('tr');
 head.append(el('th','행','row-number'));sample.columns.forEach(col=>head.append(el('th',col)));thead.append(head);table.append(thead);
 const body=el('tbody');sample.rows.forEach((values,i)=>{const tr=el('tr');tr.append(el('td',String(i+1),'row-number'));values.forEach(value=>tr.append(el('td',value)));body.append(tr);});
 table.append(body);wrap.append(table);panel.append(wrap);
}
function showTable(t){
 selected={type:'table',id:t.id};const panel=$('detail');panel.replaceChildren();panel.append(el('h2',t.label),el('p',t.file+' / '+t.sheet,'path'));
 panel.append(rowsTable([['데이터',t.rows+'행 × '+t.columns+'열'],['비교 범위',t.scope],['버전 컬럼',t.version_columns.join(', ')||'없음']]));
 if(t.versioned)panel.append(el('p','⚠ 원시 행에 여러 버전이 섞여 있습니다. 문서를 확인해 유효 버전을 결정해야 합니다.'));
 appendSample(panel,t);panel.append(el('h3','컬럼 분석','profile-heading'));
 const wrap=el('div');wrap.style.overflowX='auto';const table=el('table',undefined,'metrics');const head=el('tr');['컬럼','형식','고유값','결측','키 후보'].forEach(v=>head.append(el('th',v)));table.append(head);
 data.profiles.filter(p=>p.table_id===t.id).forEach(p=>{const tr=el('tr');[p.column,p.kind,p.distinct,p.missing,p.eligible?'검토':'—'].forEach(v=>tr.append(el('td',String(v))));table.append(tr);});wrap.append(table);panel.append(wrap);paint();
}
function renderRelationButtons(){
 const box=$('relations');box.replaceChildren();for(const r of activeEdges()){const b=el('button',byId.get(r.from).label+' → '+byId.get(r.to).label,selected?.type==='edge'&&selected.id===r.id?'active':'');b.append(el('small',r.from_column+' → '+r.to_column+' · '+pct(r.row_match_pct)));b.addEventListener('click',()=>showEdge(r));box.append(b);}
 if(!activeEdges().length)box.append(el('div','현재 필터에 해당하는 관계 후보가 없습니다.','empty'));
}
function point(e){const p=$('graph').createSVGPoint();p.x=e.clientX;p.y=e.clientY;return p.matrixTransform($('graph').getScreenCTM().inverse());}
$('graph').addEventListener('pointerdown',e=>{if(e.target.closest('[data-edge]'))return;const p=point(e),node=e.target.closest('[data-node]');drag={id:node?.dataset.node,start:p,last:p,moved:false};$('graph').setPointerCapture(e.pointerId);e.preventDefault();});
$('graph').addEventListener('pointermove',e=>{if(!drag)return;const p=point(e),dx=p.x-drag.last.x,dy=p.y-drag.last.y;if(Math.hypot(p.x-drag.start.x,p.y-drag.start.y)>4)drag.moved=true;if(drag.id){positions[drag.id].x+=dx/pan.k;positions[drag.id].y+=dy/pan.k;}else{pan.x+=dx;pan.y+=dy;}drag.last=p;paint();});
$('graph').addEventListener('pointerup',()=>{if(drag?.id&&!drag.moved)showTable(byId.get(drag.id));drag=null;});
$('graph').addEventListener('pointercancel',()=>{drag=null;});
$('graph').addEventListener('wheel',e=>{e.preventDefault();const p=point(e),old=pan.k;pan.k=Math.max(.3,Math.min(3,pan.k*Math.exp(-e.deltaY*.001)));pan.x=p.x-(p.x-pan.x)*pan.k/old;pan.y=p.y-(p.y-pan.y)*pan.k/old;paint();},{passive:false});
data.tables.forEach(t=>{const label=el('label',undefined,'table-choice'),check=el('input');check.type='checkbox';check.checked=true;check.addEventListener('change',()=>{if(check.checked)visible.add(t.id);else visible.delete(t.id);paint();});const text=el('span',t.label);text.append(el('small',t.rows+'행 · '+t.columns+'열'));label.append(check,text);$('tables').append(label);});
$('source').textContent='입력: '+data.meta.input;
$('scope').textContent='비교 범위: '+data.meta.scope_label+' · '+data.meta.generated;
const numbers=$('numbers');for(const [n,label] of [[data.tables.length,'테이블'],[data.profiles.length,'컬럼'],[data.relations.length,'관계 후보']]){numbers.append(el('strong',n),document.createTextNode(label));}
$('notice').textContent='이 그래프는 값·명칭·고유성으로 찾은 관계 후보입니다. 실제 업무 관계나 외래키 제약을 검증한 결과가 아닙니다.'+(data.tables.some(t=>t.versioned)?' 버전 혼재 테이블은 최종본 필터를 자동 적용하지 않았습니다.':'');
data.meta.rules.forEach(rule=>$('rules').append(el('li',rule)));
for(const doc of data.documents){const details=el('details',undefined,'card doc section');details.append(el('summary','참고 문서 원문 · '+doc.file),el('pre',doc.text));$('documents').append(details);}
$('foot').textContent='원본 변경 없음 · 생성 결과는 실행 당시의 스냅샷 · '+data.meta.skipped.length+'개 비테이블 파일은 테이블 분석에서 제외 · CSV 인코딩: UTF-8 BOM';
$('coverage').value=String(Math.max(50,data.meta.minimum_pct));$('coverage').addEventListener('change',paint);$('versions').addEventListener('change',paint);$('fit').addEventListener('click',()=>{initialPositions();paint();});
initialPositions();if(data.relations.length)showEdge(data.relations[0]);else if(data.tables.length)showTable(data.tables[0]);else{paint();$('detail').append(el('p','분석할 테이블이 없습니다.'));}
</script></body></html>"""


def main():
    stage_seconds = {}
    last_checkpoint = _SCRIPT_STARTED

    def checkpoint(name):
        nonlocal last_checkpoint
        now = perf_counter()
        stage_seconds[name] = round(now - last_checkpoint, 6)
        last_checkpoint = now

    checkpoint("모듈 로딩 및 코드 초기화")
    parser = argparse.ArgumentParser(description="관계 후보 자료와 독립 HTML 관계도 생성")
    parser.add_argument("--input", default="targets")
    parser.add_argument("--output", default="targets_relation_analysis")
    parser.add_argument("--scope", choices=("folder", "all"), default="folder")
    parser.add_argument("--min-coverage", type=float, default=0.5)
    args = parser.parse_args()
    if not 0 <= args.min_coverage <= 1:
        parser.error("--min-coverage는 0~1 사이여야 합니다.")
    root = (BASE_DIR / args.input).resolve()
    output = (BASE_DIR / args.output).resolve()
    if not root.is_dir():
        parser.error(f"입력 폴더 없음: {root}")
    if output == root or output.is_relative_to(root):
        parser.error("출력 폴더를 입력 폴더 내부에 두지 마세요.")
    output.mkdir(parents=True, exist_ok=True)
    tables, profiles, documents, skipped, errors = [], [], [], [], []
    checkpoint("경로 확인 및 실행 준비")
    files = sorted((p for p in root.rglob("*") if p.is_file() and not p.is_symlink()
                    and not p.name.startswith("~$")), key=lambda p: p.as_posix().casefold())
    checkpoint("파일 목록 탐색")
    for path in files:
        relative = path.relative_to(root).as_posix()
        if path.suffix.lower() in {".md", ".txt"}:
            try:
                documents.append({"file": relative, "text": read_text(path)})
            except Exception as exc:
                errors.append({"file": relative, "error": str(exc), "traceback": traceback.format_exc()})
        if path.suffix.lower() not in {".csv", ".xlsx"}:
            skipped.append(relative)
            continue
        try:
            for sheet_name, frame in read_tables(path):
                table = {
                    "id": f"T{len(tables)+1:03d}", "file": relative, "sheet": sheet_name,
                    "label": path.name if sheet_name == "CSV" else path.name + " / " + sheet_name,
                    "rows": len(frame), "columns": len(frame.columns),
                    "scope": scope_for(path, root, args.scope),
                    "versioned": False, "version_columns": [],
                }
                columns = [build_profile(table, frame.iloc[:, i], col, i)
                           for i, col in enumerate(frame.columns)]
                for profile in columns:
                    if any(token in normalize_name(profile["column"])
                           for token in ("버전", "version", "revision", "개정")):
                        table["version_columns"].append(profile["column"])
                        if profile["distinct"] > 1:
                            table["versioned"] = True
                table["sample"] = build_sample(frame)
                tables.append(table)
                profiles.extend(columns)
                print(f"[읽음] {relative} / {sheet_name}: {len(frame)}행 × {len(frame.columns)}열")
        except Exception as exc:
            errors.append({"file": relative, "error": str(exc), "traceback": traceback.format_exc()})
            print(f"[읽기 실패] {relative}: {exc}")
    checkpoint("파일 읽기 및 컬럼 분석")
    relations = infer_relations(profiles, tables, args.min_coverage)
    checkpoint("테이블 관계 추정")
    public = [public_profile(p) for p in profiles]
    meta = {
        "input": str(root), "generated": datetime.now().astimezone().isoformat(timespec="seconds"),
        "scope_label": "전체 폴더" if args.scope == "all" else "가장 가까운 README 폴더, 없으면 같은 폴더",
        "minimum_pct": args.min_coverage * 100, "skipped": skipped, "errors": errors,
        "rules": [
            "CSV/XLSX의 모든 시트를 발견해서 분석. 원본 파일·문서는 수정하지 않음.",
            "단일 컬럼만 비교하며 복합키는 탐색하지 않음.",
            "값은 NFKC/앞뒤 공백을 정리하고 대소문자는 보존. 숫자의 선행 0은 합치지 않음.",
            "키 명칭 힌트 또는 고유성이 있는 문자 컬럼만 비교. 일반 숫자 측정 컬럼은 제외.",
            f"공통 값이 있고, 방향 선택 후 비결측 원시 행 일치율 {args.min_coverage*100:.1f}% 이상인 후보만 포함.",
            "컬럼명이 비슷하거나, 3개 이상 고유값이 일치하면서 행 일치율 80% 이상인 후보만 포함.",
            "참조 대상의 고유성·값 포함률을 기준으로 화살표 방향 추정. 실제 PK/FK를 확정하지 않음.",
            "공통 고유값 3개 미만 또는 한쪽 데이터 3행 미만이면 작은 표본으로 표시.",
            "버전 컬럼에 여러 값이 있으면 경고. 어떤 버전이 유효한지 자동 결정하지 않음.",
        ],
    }
    checkpoint("출력 데이터 준비")
    write_csv(output / "tables.csv", tables, [
        ("id", "테이블ID"), ("file", "원본상대경로"), ("sheet", "시트명"),
        ("rows", "행수"), ("columns", "컬럼수"), ("scope", "비교범위"),
        ("versioned", "버전혼재"), ("version_columns", "버전컬럼"),
    ])
    write_csv(output / "column_profiles.csv", public, [
        ("id", "컬럼ID"), ("table_id", "테이블ID"), ("file", "원본상대경로"),
        ("sheet", "시트명"), ("column", "컬럼명"), ("dtype", "자료형"),
        ("kind", "비교자료형"), ("rows", "행수"), ("nonmissing", "비결측수"),
        ("distinct", "고유값수"), ("unique_pct", "고유값비율(%)"),
        ("missing", "결측수"), ("missing_pct", "결측률(%)"),
        ("key_hint", "키명칭힌트"), ("eligible", "관계비교대상"), ("examples", "예시값"),
    ])
    write_csv(output / "relations.csv", relations, [
        ("id", "관계ID"), ("from_file", "참조원본"), ("from_sheet", "참조원본시트"),
        ("from_column", "참조컬럼"), ("to_file", "대상원본"), ("to_sheet", "대상시트"),
        ("to_column", "대상컬럼"), ("scope", "비교범위"), ("cardinality", "관계형태후보"),
        ("from_rows", "참조원본행수"), ("to_rows", "대상행수"),
        ("from_nonmissing", "참조비결측수"), ("to_nonmissing", "대상비결측수"),
        ("from_unique_pct", "참조고유값비율(%)"), ("to_unique_pct", "대상고유값비율(%)"),
        ("matched_distinct", "일치고유값수"), ("matched_rows", "일치행수"),
        ("row_match_pct", "원시행일치율(%)"), ("distinct_match_pct", "고유값일치율(%)"),
        ("reverse_distinct_pct", "반대쪽고유값포함률(%)"),
        ("unmatched_rows", "미일치행수"), ("unmatched_distinct", "미일치고유값수"),
        ("matched_examples", "일치예시"), ("unmatched_examples", "미일치예시"),
        ("status", "후보상태"), ("versioned", "버전혼재"), ("cautions", "검토사항"),
    ])
    checkpoint("CSV 저장")
    data = {"tables": tables, "profiles": public, "relations": relations,
            "documents": documents, "meta": meta}
    embedded = json.dumps(data, ensure_ascii=False).replace("<", r"\u003c").replace("\u2028", r"\u2028").replace("\u2029", r"\u2029")
    (output / "relations.html").write_text(HTML.replace("__DATA__", embedded), encoding="utf-8")
    checkpoint("HTML 생성 및 저장")
    if errors:
        (output / "errors.json").write_text(json.dumps(errors, ensure_ascii=False, indent=2), encoding="utf-8")
    checkpoint("오류 내역 저장")
    print(f"\n[완료] 테이블 {len(tables)}개 / 컬럼 {len(profiles)}개 / 관계 후보 {len(relations)}개 / 읽기 실패 {len(errors)}개")
    print(output / "relations.html")
    for relation in relations:
        print(f"  {relation['from_file']}.{relation['from_column']} → "
              f"{relation['to_file']}.{relation['to_column']}: {relation['row_match_pct']}%, "
              f"미일치 {relation['unmatched_examples']}")
    checkpoint("완료 결과 출력")
    timing = {
        "시작시각": datetime.fromtimestamp(_SCRIPT_STARTED_TIMESTAMP).astimezone().isoformat(timespec="milliseconds"),
        "종료시각": datetime.now().astimezone().isoformat(timespec="milliseconds"),
        "상태": "완료" if not errors else "일부 파일 읽기 실패",
        "입력폴더": str(root),
        "출력폴더": str(output),
        "테이블수": len(tables), "컬럼수": len(profiles),
        "관계후보수": len(relations), "읽기실패수": len(errors),
        "전체시간_초": round(last_checkpoint - _SCRIPT_STARTED, 6),
        "단계별시간_초": stage_seconds,
        "측정범위": "스크립트 코드 시작부터 결과 출력까지. 모듈 로딩 포함. Python 프로세스 기동, 이 JSON 저장, 브라우저 표시 시간 제외.",
    }
    timing_path = output / "latest_run_timing.json"
    timing_path.write_text(json.dumps(timing, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[실행 시간] {timing['전체시간_초']:.6f}초 / 기록: {timing_path}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
