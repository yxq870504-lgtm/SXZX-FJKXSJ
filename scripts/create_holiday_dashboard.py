from pathlib import Path
import os
from openpyxl import load_workbook
from datetime import datetime
from collections import defaultdict
import json

repo_root = Path(os.environ.get('DASHBOARD_REPO_DIR', Path(__file__).resolve().parents[1]))
source_file = Path(os.environ.get('SOURCE_XLSX', repo_root / 'data' / 'holiday_student_analysis.xlsx'))
out_file = repo_root / 'index.html'

GRADE_ORDER = ['初一', '初二', '初三', '高一', '高二', '高三']
STAGE_GRADES = {'初中': ['初一', '初二', '初三'], '高中': ['高一', '高二', '高三']}
SPECIAL_SHORT = {
    '内蒙古自治区': '内蒙古', '广西壮族自治区': '广西', '宁夏回族自治区': '宁夏',
    '新疆维吾尔自治区': '新疆', '西藏自治区': '西藏', '黑龙江省': '黑龙江',
    '香港特别行政区': '香港', '澳门特别行政区': '澳门'
}


def short_province(name):
    if not name:
        return ''
    if name in SPECIAL_SHORT:
        return SPECIAL_SHORT[name]
    s = str(name)
    for suf in ['省', '市', '维吾尔自治区', '壮族自治区', '回族自治区', '自治区', '特别行政区']:
        s = s.replace(suf, '')
    return s[:3]


def norm_province(v):
    s = str(v or '').strip()
    for suf in ['维吾尔自治区', '壮族自治区', '回族自治区', '特别行政区', '自治区', '省', '市']:
        s = s.replace(suf, '')
    return s


def fmt_date(v):
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d')
    return str(v)[:10] if v else ''

wb = load_workbook(source_file, data_only=True)
ws = wb['省份日历_人数底表']
source_lookup = {}
if '0805最新校历来源' in wb.sheetnames:
    sws = wb['0805最新校历来源']
    for r in range(2, sws.max_row + 1):
        p = norm_province(sws.cell(r, 1).value)
        g = sws.cell(r, 3).value
        if p and g:
            source_lookup[(p, g)] = sws.cell(r, 6).value or ''

records = []
for r in range(2, ws.max_row + 1):
    province = ws.cell(r, 2).value
    if not province:
        continue
    item = {'province': province, 'short': short_province(province), 'region': ws.cell(r, 1).value or '', 'grades': {}}
    c = 3
    pkey = norm_province(province)
    for g in GRADE_ORDER:
        source = source_lookup.get((pkey, g), '')
        item['grades'][g] = {
            'holiday': fmt_date(ws.cell(r, c).value),
            'spring': fmt_date(ws.cell(r, c + 1).value),
            'isReal': ws.cell(r, c + 2).value or '否',
            'count': int(ws.cell(r, c + 3).value or 0),
            'pct': float(ws.cell(r, c + 4).value or 0),
            'source': source,
            'official': bool(source and source != '待官方发布后更新'),
        }
        c += 5
    records.append(item)

# Approximate China-map-style layout: visual province tiles arranged west/north/east/south, colored by selected grade count.
MAP_POSITIONS = {
    '新疆': [1, 1], '西藏': [1, 3], '青海': [2, 3], '甘肃': [3, 2], '宁夏': [4, 2], '内蒙古': [5, 1],
    '黑龙江': [9, 1], '吉林': [9, 2], '辽宁': [8, 3], '北京': [7, 3], '天津': [8, 4], '河北': [7, 4],
    '山西': [6, 4], '陕西': [5, 4], '河南': [6, 5], '山东': [8, 5], '江苏': [8, 6], '上海': [9, 7],
    '安徽': [7, 6], '湖北': [6, 6], '四川': [4, 6], '重庆': [5, 7], '贵州': [5, 8], '云南': [4, 9],
    '湖南': [6, 7], '江西': [7, 7], '浙江': [8, 7], '福建': [8, 8], '广东': [7, 9], '广西': [6, 9], '海南': [7, 10]
}

payload = {
    'records': records,
    'stageGrades': STAGE_GRADES,
    'mapPositions': MAP_POSITIONS,
    'generatedAt': datetime.now().strftime('%Y-%m-%d %H:%M'),
}
json_data = json.dumps(payload, ensure_ascii=False)

html = r'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>有道升学-正价学员放假&开学时间分布</title>
<style>
:root{--bg:#f5f7fb;--card:#fff;--ink:#1f2630;--muted:#738096;--blue:#4f72f6;--blue2:#3e66e4;--bar:#75b9f3;--purple:#7b42ef;--line:#e6edf7;--shadow:0 16px 48px rgba(46,73,123,.10)}
*{box-sizing:border-box}body{margin:0;background:var(--bg);font-family:"PingFang SC","Microsoft YaHei",sans-serif;color:var(--ink)}
.shell{max-width:1840px;margin:0 auto;padding:22px 36px 54px}.header{position:relative;min-height:108px;display:flex;align-items:center;justify-content:center}.refresh-note{position:absolute;left:0;top:10px;color:#6f7f99;font-size:18px;font-weight:700}.title{font-size:58px;line-height:1;white-space:nowrap;color:#5975f6;font-weight:900;letter-spacing:-.04em;margin:0}.section-head{margin:28px 0 14px;display:flex;align-items:flex-end;justify-content:space-between;gap:24px}.section-head h2{margin:0;color:#5475f8;font-size:38px;font-weight:900}.section-sub{color:#8a98ad;font-weight:700}.rule{height:4px;background:#5a78f6;border-radius:99px;margin-top:14px}.panel,.kpis{border-radius:24px;background:rgba(255,255,255,.76);border:1px solid #dfe8fb;box-shadow:var(--shadow)}.panel{padding:20px;background:#f7f8fb}.panel-top{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:18px}.pill{display:flex;width:max-content;gap:6px;background:#eef3ff;border:1px solid #dfe8fb;border-radius:999px;padding:5px}.pill button{border:0;background:transparent;padding:10px 30px;border-radius:999px;font-size:18px;font-weight:900;color:#59677d;cursor:pointer}.pill button.active{background:#fff;color:#315be4;box-shadow:0 8px 20px rgba(79,114,246,.15)}.data-note{font-size:15px;color:#738096;font-weight:700}.kpis{padding:20px;margin-bottom:20px}.kpi-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:16px}.kpi{background:#fff;border-radius:16px;min-height:118px;display:flex;flex-direction:column;align-items:center;justify-content:center;border:1px solid #eff3fa}.kpi .label{font-size:20px;font-weight:900;margin-bottom:14px;text-align:center}.kpi .value{font-size:48px;color:#5b7cff;font-weight:300;line-height:1}.kpi .value.small{font-size:30px;font-weight:600}.dist-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}.card{background:#fff;border:1px solid #edf2f7;border-radius:14px;overflow:hidden;box-shadow:0 6px 20px rgba(73,86,112,.05)}.card-head{padding:18px 24px 14px;border-bottom:1px solid #edf2f7;font-size:25px;font-weight:900}.card-body{padding:14px 18px}.bar-chart{height:auto;padding:4px 2px}.bar-row{display:grid;grid-template-columns:104px 1fr;align-items:center;height:30px;gap:10px}.bar-date{font-size:14px;color:#666;text-align:left}.bar-track{height:22px;position:relative;background:transparent;display:flex;justify-content:center}.bar-fill{height:22px;background:#76b9f2;min-width:2px;display:flex;align-items:center;justify-content:center;color:#24506e;font-weight:500;font-size:14px}.bar-fill.purple{background:#874cf3;color:#fff}.map-layout{display:grid;grid-template-columns:minmax(740px,1.1fr) minmax(360px,.9fr);gap:22px;align-items:stretch}.china-map{position:relative;min-height:620px;background:linear-gradient(145deg,#f8fbff,#eef4ff);border:1px solid #dce7fb;border-radius:22px;overflow:hidden}.china-map:before{content:'';position:absolute;inset:34px 48px;background:radial-gradient(circle at 68% 36%,rgba(79,114,246,.12),transparent 28%),radial-gradient(circle at 48% 62%,rgba(117,185,243,.18),transparent 34%);filter:blur(.2px)}.province-tile{position:absolute;width:82px;height:54px;border-radius:14px;border:1px solid rgba(255,255,255,.72);box-shadow:0 8px 22px rgba(63,91,142,.14);display:flex;flex-direction:column;align-items:center;justify-content:center;color:#fff;transform:translate(-50%,-50%)}.province-tile b{font-size:18px;line-height:1}.province-tile span{font-size:12px;margin-top:4px;opacity:.95}.rank-panel{background:#fff;border:1px solid #edf2f7;border-radius:22px;padding:22px;box-shadow:0 6px 20px rgba(73,86,112,.05)}.rank-title{font-size:26px;font-weight:900;color:#315be4;margin-bottom:14px}.rank-chain{font-size:28px;font-weight:900;line-height:1.6;color:#1f2630;word-break:keep-all}.rank-list{margin-top:18px;display:grid;gap:10px}.rank-item{display:grid;grid-template-columns:36px 1fr auto;gap:12px;align-items:center;padding:12px 14px;border-radius:14px;background:#f7f9fe;border:1px solid #edf2f7}.rank-no{width:28px;height:28px;border-radius:50%;background:#eef3ff;color:#315be4;font-weight:900;display:flex;align-items:center;justify-content:center}.rank-name{font-weight:900}.rank-count{color:#5b7cff;font-weight:900}.legend{display:flex;align-items:center;gap:12px;margin-top:16px;color:#738096;font-weight:700}.legend-bar{width:180px;height:12px;border-radius:99px;background:linear-gradient(90deg,#dbeafe,#75b9f3,#315be4)}@media(max-width:1300px){.title{font-size:42px}.kpi-grid{grid-template-columns:repeat(3,1fr)}.map-layout,.dist-grid{grid-template-columns:1fr}.china-map{min-height:560px}}@media(max-width:820px){.shell{padding:18px 16px 36px}.title{font-size:30px;white-space:normal}.header{justify-content:flex-start;padding-top:62px}.section-head{display:block}.kpi-grid{grid-template-columns:repeat(2,1fr)}.pill{width:100%;overflow:auto}.province-tile{width:72px;height:48px}.china-map{min-height:620px}}
</style>
</head>
<body>
<div class="shell">
  <header class="header">
    <div class="refresh-note">每天10点自动刷新官方校历</div>
    <h1 class="title">有道升学-正价学员放假&开学时间分布</h1>
  </header>

  <section class="section-head"><h2>①放假&开学时间分布图</h2><span class="section-sub">按省份 × 年级在班人数加权展示</span></section><div class="rule"></div>
  <main class="main">
    <div class="panel" style="margin-top:18px">
      <div class="panel-top">
        <div class="pill" id="stageTabs"><button data-stage="初中">初中</button><button data-stage="高中" class="active">高中</button></div>
        <div class="data-note" id="stageNote">默认展示高中板块</div>
      </div>
      <div class="kpis"><div class="kpi-grid">
        <div class="kpi"><div class="label">覆盖省份数</div><div class="value" id="kpiProvince">0</div></div>
        <div class="kpi"><div class="label">官方校历发布省份数</div><div class="value" id="kpiOfficial">0</div></div>
        <div class="kpi"><div class="label">最早放假时间</div><div class="value small" id="kpiHMin">-</div></div>
        <div class="kpi"><div class="label">最晚放假时间</div><div class="value small" id="kpiHMax">-</div></div>
        <div class="kpi"><div class="label">最早开学时间</div><div class="value small" id="kpiSMin">-</div></div>
        <div class="kpi"><div class="label">最晚开学时间</div><div class="value small" id="kpiSMax">-</div></div>
      </div></div>
      <div class="dist-grid">
        <div class="card"><div class="card-head">寒假放假时间聚合分布</div><div class="card-body"><div id="holidayBars" class="bar-chart"></div></div></div>
        <div class="card"><div class="card-head">春季开学时间聚合分布</div><div class="card-body"><div id="springBars" class="bar-chart"></div></div></div>
      </div>
    </div>

    <section class="section-head"><h2>②在班学员所在省份分布图</h2><span class="section-sub">按当前学段下所选年级展示</span></section><div class="rule"></div>
    <div class="panel" style="margin-top:18px">
      <div class="panel-top">
        <div class="pill" id="gradeTabs"></div>
        <div class="data-note" id="gradeNote">颜色越深，在班人数越多</div>
      </div>
      <div class="map-layout">
        <div class="china-map" id="chinaMap"></div>
        <div class="rank-panel">
          <div class="rank-title" id="rankTitle">省份Top排名</div>
          <div class="rank-chain" id="rankChain"></div>
          <div class="rank-list" id="rankList"></div>
          <div class="legend"><span>少</span><div class="legend-bar"></div><span>多</span></div>
        </div>
      </div>
    </div>
  </main>
</div>
<script>
const DATA = __DATA__;
let stage = '高中';
let grade = '高一';
const fmtNum = n => (n || 0).toLocaleString('zh-CN');
const fmtPct = n => ((n || 0) * 100).toFixed(2) + '%';
const grades = () => DATA.stageGrades[stage];
function stageRecords(){return DATA.records.filter(r=>grades().some(g=>r.grades[g].count>0 || r.grades[g].holiday || r.grades[g].spring));}
function dist(event){const buckets={}; stageRecords().forEach(rec=>grades().forEach(g=>{const gd=rec.grades[g]; const d=gd[event]; if(!d)return; if(!buckets[d])buckets[d]={count:0,provinceSet:new Set()}; buckets[d].count+=gd.count; buckets[d].provinceSet.add(rec.province)})); const total=Object.values(buckets).reduce((a,b)=>a+b.count,0); let cum=0; return Object.keys(buckets).sort().map(d=>{const cnt=buckets[d].count; cum+=cnt; return {date:d,count:cnt,cumulative:cum,cumulativeShare:total?cum/total:0,provinceCount:buckets[d].provinceSet.size}})}
function dateRange(event){const arr=[];stageRecords().forEach(r=>grades().forEach(g=>{const d=r.grades[g][event]; if(d)arr.push(d)}));arr.sort();return [arr[0]||'-',arr[arr.length-1]||'-']}
function updateKpi(){const ps=new Set();const official=new Set();stageRecords().forEach(r=>{let hasOfficial=false; grades().forEach(g=>{const gd=r.grades[g]; if(gd.count>0 || gd.holiday || gd.spring)ps.add(r.province); if(gd.official)hasOfficial=true;}); if(hasOfficial)official.add(r.province)}); const hr=dateRange('holiday'), sr=dateRange('spring'); document.getElementById('kpiProvince').textContent=ps.size; document.getElementById('kpiOfficial').textContent=official.size; document.getElementById('kpiHMin').textContent=hr[0]; document.getElementById('kpiHMax').textContent=hr[1]; document.getElementById('kpiSMin').textContent=sr[0]; document.getElementById('kpiSMax').textContent=sr[1]; document.getElementById('stageNote').textContent=`当前学段：${stage}｜更新时间：${DATA.generatedAt}`;}
function renderBars(id,event,colorClass=''){const rows=dist(event);document.getElementById(id).innerHTML=rows.map(r=>`<div class="bar-row"><div class="bar-date">${r.date}</div><div class="bar-track"><div class="bar-fill ${colorClass}" style="width:${Math.max(0.8,r.cumulativeShare*100)}%">${fmtPct(r.cumulativeShare)}</div></div></div>`).join('')}
function colorFor(v,max){if(!max)return '#dbeafe'; const t=Math.max(0.08, v/max); const l=88 - Math.round(t*38); return `hsl(224 84% ${l}%)`}
function provinceRows(){return DATA.records.map(r=>({province:r.province,short:r.short,count:r.grades[grade].count,pct:r.grades[grade].pct,source:r.grades[grade].source})).sort((a,b)=>b.count-a.count||a.province.localeCompare(b.province,'zh-CN'))}
function renderMap(){const rows=provinceRows(); const max=Math.max(...rows.map(x=>x.count),0); const map=document.getElementById('chinaMap'); map.innerHTML=rows.map(x=>{const pos=DATA.mapPositions[x.short]||[5,5]; const left=8+pos[0]*8.4; const top=5+pos[1]*8.2; return `<div class="province-tile" title="${x.province}：${fmtNum(x.count)}人" style="left:${left}%;top:${top}%;background:${colorFor(x.count,max)}"><b>${x.short}</b><span>${fmtNum(x.count)}</span></div>`}).join(''); const top5=rows.slice(0,5); document.getElementById('rankTitle').textContent=`${grade} 在班人数省份Top排名`; document.getElementById('rankChain').textContent=top5.map(x=>x.short).join(' > '); document.getElementById('rankList').innerHTML=top5.map((x,i)=>`<div class="rank-item"><div class="rank-no">${i+1}</div><div class="rank-name">${x.short}</div><div class="rank-count">${fmtNum(x.count)}人</div></div>`).join(''); document.getElementById('gradeNote').textContent=`当前年级：${grade}｜颜色越深，在班人数越多`;}
function renderGradeTabs(){const tabs=document.getElementById('gradeTabs'); tabs.innerHTML=grades().map(g=>`<button data-grade="${g}" class="${grade===g?'active':''}">${g}</button>`).join(''); tabs.querySelectorAll('button').forEach(b=>b.onclick=()=>{grade=b.dataset.grade; renderGradeTabs(); renderMap();});}
function render(){updateKpi();renderBars('holidayBars','holiday');renderBars('springBars','spring','purple');renderGradeTabs();renderMap();}
document.querySelectorAll('#stageTabs button').forEach(b=>b.onclick=()=>{document.querySelectorAll('#stageTabs button').forEach(x=>x.classList.remove('active')); b.classList.add('active'); stage=b.dataset.stage; grade=grades()[0]; render();});
render();
</script>
</body>
</html>'''.replace('__DATA__', json_data)

out_file.write_text(html, encoding='utf-8')
print(out_file)
