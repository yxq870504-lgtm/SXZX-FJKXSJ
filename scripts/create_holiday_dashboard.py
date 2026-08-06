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
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js"></script>
<style>
:root{--bg:#f5f7fb;--card:#fff;--ink:#1f2630;--muted:#738096;--blue:#4f72f6;--blue2:#3e66e4;--bar:#75b9f3;--purple:#7b42ef;--line:#e6edf7;--shadow:0 16px 48px rgba(46,73,123,.10)}
*{box-sizing:border-box}body{margin:0;background:var(--bg);font-family:"PingFang SC","Microsoft YaHei",sans-serif;color:var(--ink)}
.shell{max-width:1840px;margin:0 auto;padding:22px 36px 54px}.header{position:relative;min-height:108px;display:flex;align-items:center;justify-content:center}.refresh-note{position:absolute;left:0;top:10px;color:#6f7f99;font-size:18px;font-weight:700}.title{font-size:58px;line-height:1;white-space:nowrap;color:#5975f6;font-weight:900;letter-spacing:-.04em;margin:0}.section-head{margin:28px 0 14px;display:flex;align-items:flex-end;justify-content:space-between;gap:24px}.section-head h2{margin:0;color:#5475f8;font-size:38px;font-weight:900}.section-sub{color:#8a98ad;font-weight:700}.rule{height:4px;background:#5a78f6;border-radius:99px;margin-top:14px}.panel,.kpis{border-radius:24px;background:rgba(255,255,255,.76);border:1px solid #dfe8fb;box-shadow:var(--shadow)}.panel{padding:20px;background:#f7f8fb}.panel-top{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:18px}.pill{display:flex;width:max-content;gap:6px;background:#eef3ff;border:1px solid #dfe8fb;border-radius:999px;padding:5px}.pill button{border:0;background:transparent;padding:10px 30px;border-radius:999px;font-size:18px;font-weight:900;color:#59677d;cursor:pointer}.pill button.active{background:#fff;color:#315be4;box-shadow:0 8px 20px rgba(79,114,246,.15)}.data-note{font-size:15px;color:#738096;font-weight:700}.kpis{padding:20px;margin-bottom:20px}.kpi-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:16px}.kpi{background:#fff;border-radius:16px;min-height:118px;display:flex;flex-direction:column;align-items:center;justify-content:center;border:1px solid #eff3fa}.kpi .label{font-size:20px;font-weight:900;margin-bottom:14px;text-align:center}.kpi .value{font-size:48px;color:#5b7cff;font-weight:300;line-height:1}.kpi .value.small{font-size:30px;font-weight:600}.dist-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}.card{background:#fff;border:1px solid #edf2f7;border-radius:14px;overflow:hidden;box-shadow:0 6px 20px rgba(73,86,112,.05)}.card-head{padding:18px 24px 14px;border-bottom:1px solid #edf2f7;font-size:25px;font-weight:900}.card-body{padding:14px 18px}.bar-chart{height:auto;padding:4px 2px}.bar-row{display:grid;grid-template-columns:104px 1fr;align-items:center;height:30px;gap:10px}.bar-date{font-size:14px;color:#666;text-align:left}.bar-track{height:22px;position:relative;background:transparent;display:flex;justify-content:center}.bar-fill{height:22px;background:#76b9f2;min-width:2px;display:flex;align-items:center;justify-content:center;color:#24506e;font-weight:500;font-size:14px}.bar-fill.purple{background:#874cf3;color:#fff}.map-wrap{background:#1d2430;border:1px solid #dce7fb;border-radius:22px;overflow:hidden;box-shadow:0 10px 34px rgba(31,38,48,.16)}.china-map{height:680px;width:100%;background:#1d2430}.top5-copy{margin-top:16px;padding:18px 22px;border-radius:18px;background:#fff;border:1px solid #edf2f7;color:#1f2630;font-size:24px;font-weight:900;line-height:1.55;box-shadow:0 6px 20px rgba(73,86,112,.05)}.top5-copy span{color:#315be4}.map-fallback{height:680px;display:flex;align-items:center;justify-content:center;color:#dbeafe;font-size:20px;font-weight:900;background:#1d2430}.legend{display:flex;align-items:center;gap:12px;margin-top:12px;color:#738096;font-weight:700}.legend-bar{width:180px;height:12px;border-radius:99px;background:linear-gradient(90deg,#16466b,#157a66,#85a71e)}@media(max-width:1300px){.title{font-size:42px}.kpi-grid{grid-template-columns:repeat(3,1fr)}.dist-grid{grid-template-columns:1fr}.china-map{height:600px}}@media(max-width:820px){.shell{padding:18px 16px 36px}.title{font-size:30px;white-space:normal}.header{justify-content:flex-start;padding-top:62px}.section-head{display:block}.kpi-grid{grid-template-columns:repeat(2,1fr)}.pill{width:100%;overflow:auto}.china-map{height:560px}}
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
      <div class="map-wrap"><div class="china-map" id="chinaMap"></div></div>
      <div class="top5-copy" id="rankCopy">在班学员所在省份top5：</div>
      <div class="legend"><span>人数少</span><div class="legend-bar"></div><span>人数多</span></div>
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
let mapChart = null;
let chinaMapReady = null;
const NAME_ALIAS = {'北京':'北京市','天津':'天津市','上海':'上海市','重庆':'重庆市','河北':'河北省','山西':'山西省','辽宁':'辽宁省','吉林':'吉林省','黑龙江':'黑龙江省','江苏':'江苏省','浙江':'浙江省','安徽':'安徽省','福建':'福建省','江西':'江西省','山东':'山东省','河南':'河南省','湖北':'湖北省','湖南':'湖南省','广东':'广东省','海南':'海南省','四川':'四川省','贵州':'贵州省','云南':'云南省','陕西':'陕西省','甘肃':'甘肃省','青海':'青海省','台湾':'台湾省','内蒙古':'内蒙古自治区','广西':'广西壮族自治区','西藏':'西藏自治区','宁夏':'宁夏回族自治区','新疆':'新疆维吾尔自治区','香港':'香港特别行政区','澳门':'澳门特别行政区'};
function provinceRows(){return DATA.records.map(r=>({name:NAME_ALIAS[r.short]||r.province,province:r.province,short:r.short,count:r.grades[grade].count,pct:r.grades[grade].pct,source:r.grades[grade].source})).sort((a,b)=>b.count-a.count||a.province.localeCompare(b.province,'zh-CN'))}
async function ensureChinaMap(){
  if(chinaMapReady) return chinaMapReady;
  chinaMapReady = fetch('https://geo.datav.aliyun.com/areas_v3/bound/100000_full.json').then(r=>r.json()).then(geo=>{echarts.registerMap('china', geo); return true;}).catch(()=>false);
  return chinaMapReady;
}
async function renderMap(){
  const rows=provinceRows(); const max=Math.max(...rows.map(x=>x.count),0); const top5=rows.slice(0,5);
  document.getElementById('rankCopy').innerHTML=`<span>${grade}</span>在班学员所在省份top5：${top5.map(x=>x.short).join(' > ')}`;
  document.getElementById('gradeNote').textContent=`当前年级：${grade}｜鼠标悬停省份可查看在班人数与占比`;
  const ok = await ensureChinaMap(); const el=document.getElementById('chinaMap');
  if(!ok || !window.echarts){el.innerHTML='<div class="map-fallback">地图资源加载失败，请刷新页面重试</div>'; return;}
  if(!mapChart) mapChart = echarts.init(el);
  mapChart.setOption({
    backgroundColor:'#1d2430',
    tooltip:{trigger:'item',backgroundColor:'#6a4ce6',borderColor:'#6a4ce6',borderWidth:0,textStyle:{color:'#fff',fontSize:15,fontWeight:800},formatter:p=>{const d=p.data||{};return `${d.short||p.name}<br/>在班人数 <b style="color:#ffe66d">${fmtNum(d.count||0)}</b><br/>在班占比 <b style="color:#ffe66d">${fmtPct(d.pct||0)}</b>`}},
    visualMap:{min:0,max:max||1,left:28,bottom:24,text:['人数多','人数少'],textStyle:{color:'#c8d3e5',fontWeight:800},calculable:true,inRange:{color:['#16466b','#157a66','#85a71e']}},
    series:[{type:'map',map:'china',roam:false,zoom:1.12,top:22,bottom:28,label:{show:true,color:'#e7eef8',fontSize:13,fontWeight:900},emphasis:{label:{color:'#fff',fontWeight:900},itemStyle:{areaColor:'#ffffff',borderColor:'#ffffff',borderWidth:1.8}},itemStyle:{borderColor:'#8fb3c9',borderWidth:1.2,areaColor:'#16466b'},data:rows}]
  }, true);
  setTimeout(()=>mapChart && mapChart.resize(), 0);
}
window.addEventListener('resize',()=>mapChart&&mapChart.resize());
function renderGradeTabs(){const tabs=document.getElementById('gradeTabs'); tabs.innerHTML=grades().map(g=>`<button data-grade="${g}" class="${grade===g?'active':''}">${g}</button>`).join(''); tabs.querySelectorAll('button').forEach(b=>b.onclick=()=>{grade=b.dataset.grade; renderGradeTabs(); renderMap();});}
function render(){updateKpi();renderBars('holidayBars','holiday');renderBars('springBars','spring','purple');renderGradeTabs();renderMap();}
document.querySelectorAll('#stageTabs button').forEach(b=>b.onclick=()=>{document.querySelectorAll('#stageTabs button').forEach(x=>x.classList.remove('active')); b.classList.add('active'); stage=b.dataset.stage; grade=grades()[0]; render();});
render();
</script>
</body>
</html>'''.replace('__DATA__', json_data)

out_file.write_text(html, encoding='utf-8')
print(out_file)
