const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');

const htmlPath = path.resolve(process.argv[2] || 'index.html');
const sourceDir = path.resolve(process.argv[3] || '.culture-safe-map-build/source');
const html = fs.readFileSync(htmlPath, 'utf8');
const lock = JSON.parse(fs.readFileSync(path.join(__dirname, 'template_lock.json'), 'utf8'));
const scripts = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(x => x[1]);
for (let i = 0; i < scripts.length; i++) new vm.Script(scripts[i], {filename: `inline-${i}.js`});
const oldStyles = [...html.matchAll(/<style([^>]*)>([\s\S]*?)<\/style>/g)].filter(x => !x[1].includes('v30-hwpx-download')).map(x => x[2]);
const styleHash = crypto.createHash('sha256').update(JSON.stringify(oldStyles)).digest('hex');
assert.equal(styleHash, lock.approvedStyleSha256, '승인 CSS 전체 보존');
const staticMarkup = html.replace(/<script[^>]*>[\s\S]*?<\/script>/g,'');
const ids = [...staticMarkup.matchAll(/\bid="([^"]+)"/g)].map(x => x[1]);
assert.equal(ids.length, new Set(ids).size, '중복 HTML ID 없음');
assert.equal(ids.filter(x => x === 'noticeHwpxDownloadBtn').length, 1);
const main = scripts.find(x => x.includes('const NOTICE_AUTOMATION='));
assert(main, '공문 데이터 바인딩 영역');
assert(main.includes("fetch(url,{cache:'no-store'})"));
const cut = main.slice(0, main.lastIndexOf('(async()=>'));
const elements = new Map();
function element(id) {
  if (!elements.has(id)) {
    const classes = new Set();
    const parent = {clientWidth: 640, querySelector(){return null}, appendChild(){}};
    elements.set(id, {
      id, value: '', hidden: false, disabled: false, innerHTML: '', textContent: '', open: true,
      style: {setProperty(k,v){this[k]=v},removeProperty(k){delete this[k]}}, parentElement: parent,
      clientWidth: 640, offsetWidth: 300, offsetHeight: 220, children: [],
      classList: {add(...x){x.forEach(y=>classes.add(y))},remove(...x){x.forEach(y=>classes.delete(y))},contains(x){return classes.has(x)},toggle(x){if(classes.has(x)){classes.delete(x);return false}classes.add(x);return true}},
      setAttribute(k,v){this[k]=v}, getAttribute(){return ''}, addEventListener(){},
      querySelector(){return null}, querySelectorAll(){return []}, closest(){return null}, appendChild(x){this.children.push(x)},
      getBoundingClientRect(){return {left:0,top:0,width:640,height:180}}, scrollTo(){}, click(){}, remove(){},
    });
  }
  return elements.get(id);
}
const csv = {
  'disease.csv': fs.readFileSync(path.join(sourceDir,'disease_df.csv'),'utf8'),
  'ili_region.csv': fs.readFileSync(path.join(sourceDir,'ili_region_df_260923.csv'),'utf8'),
  'performance_master.csv': fs.readFileSync(path.join(sourceDir,'performance_master_df.csv'),'utf8'),
};
const requested = [];
let downloadedBlob;
const workspace = element('__workspace');
const ctx = {
  console, Intl, Date, Map, Set, URLSearchParams, Number, String, Math, Array, JSON,
  Uint8Array, ArrayBuffer, Blob, Promise, setTimeout, clearTimeout, setImmediate, clearImmediate,
  atob: x => Buffer.from(x,'base64').toString('binary'),
  URL: {createObjectURL(blob){downloadedBlob=blob;return 'blob:test-hwpx'},revokeObjectURL(){}},
  innerWidth: 1440, addEventListener(){},
  fetch: async (url, options) => {
    assert.equal(options.cache,'no-store');
    assert(String(url).startsWith('https://raw.githubusercontent.com/Yu-Hyemin/culture_safe_map/main/output/'));
    const name = Object.keys(csv).find(x => String(url).includes(x));
    assert(name,'승인 CSV만 조회'); requested.push(name);
    return {ok:true,status:200,text:async()=>csv[name]};
  },
  document: {
    getElementById: element, body: {...element('__body'),className:''},
    querySelector: x => x === '.workspace' ? workspace : null, querySelectorAll: ()=>[],
    createElement: type => element(`__created_${type}_${elements.size}`),
  },
  getComputedStyle:()=>({columnGap:'14px',paddingLeft:'14px',paddingRight:'14px'}),
  requestAnimationFrame: fn=>fn(),cancelAnimationFrame(){},
};
ctx.window=ctx; ctx.self=ctx;
vm.createContext(ctx);

(async()=>{
  const zipScript = scripts.find(x => x.includes('JSZip v'));
  assert(zipScript,'번들 JSZip');
  vm.runInContext(zipScript,ctx,{timeout:30000});
  vm.runInContext(cut,ctx,{timeout:30000});
  await vm.runInContext('(async()=>{await loadRemoteCsvData();rebuildRemoteDerivedData()})()',ctx,{timeout:30000});
  assert.deepEqual(requested.sort(),Object.keys(csv).sort());
  const counts = vm.runInContext('({performance:PERFORMANCE_DATA.length,disease:DISEASE_DATA.length,ili:ILI_REGION_DATA.length,data:DATA.length})',ctx);
  assert(counts.performance>0 && counts.disease>0 && counts.ili>0);
  assert.equal(counts.data,counts.performance);
  vm.runInContext('state.authoritySeatFilter=0',ctx);
  const unfiltered = vm.runInContext('monitoringFilterRows(DATA).length',ctx);
  const threshold = vm.runInContext('Math.max(...DATA.map(x=>Number(x.seats)||0))',ctx);
  ctx.threshold=threshold;
  vm.runInContext('state.authoritySeatFilter=threshold',ctx);
  const filtered = vm.runInContext('monitoringFilterRows(DATA)',ctx);
  assert(filtered.every(x=>Number(x.seats)>=threshold));
  if (threshold>0) assert(filtered.length<unfiltered,'좌석 필터 실제 변경');
  vm.runInContext('state.authoritySeatFilter=0;state.role="authority";authorityPane="notice"',ctx);
  const eligible = vm.runInContext("NOTICE_EVENTS.find(ev=>{const r=NOTICE_AUTOMATION.events[ev.id];return r&&['우선 대비','대비 필요'].includes(r.status)&&r.rows.some(x=>['우선 대비','대비 필요'].includes(x.status))})",ctx);
  if (eligible) {
    element('authNoticeEvent').value=eligible.id;
    vm.runInContext('renderNoticePanel();createAuthorityNoticeDraft()',ctx);
    assert.equal(element('authNoticeDraftBtn').disabled,false);
    assert.equal(element('noticeDocEventName').textContent,eligible.name);
    assert(element('authNoticeDraftText').value.includes(eligible.name));
    assert.equal(element('authNoticeDraftWrap').hidden,false);
    assert.equal(element('noticeHwpxDownloadBtn').disabled,false);
    const selectedMarkup=vm.runInContext(`mapEventVenueHTML(NOTICE_EVENT_BY_ID.get(${JSON.stringify(eligible.id)}),true)`,ctx);
    assert(selectedMarkup.includes('data-notice-event-id='));
    const tip=element('mapTooltip');
    let handle;
    tip.querySelectorAll=selector=>selector==='[data-notice-event-id]'?[{dataset:{noticeEventId:eligible.id},addEventListener(t,fn){if(t==='click')handle=fn}}]:[];
    vm.runInContext(`pinMapTooltip(${JSON.stringify(selectedMarkup)},{originalEvent:{clientX:40,clientY:40}})`,ctx);
    assert(handle,'공문 지도 공연 선택 연결');
    handle({stopPropagation(){}});
    assert.equal(vm.runInContext('authorityPane',ctx),'notice');
    assert.equal(element('noticeDocEventName').textContent,eligible.name);
    await vm.runInContext('downloadSelectedNoticeHwpx()',ctx,{timeout:30000});
    assert(downloadedBlob,'실제 브라우저 함수로 HWPX 생성');
    const zip=await ctx.JSZip.loadAsync(new Uint8Array(await downloadedBlob.arrayBuffer()));
    assert.equal(await zip.file('mimetype').async('string'),'application/hwp+zip');
    const names=Object.keys(zip.files).filter(x=>/^Contents\/section\d+\.xml$/.test(x));
    assert(names.length>0);
    const allSections=(await Promise.all(names.map(x=>zip.file(x).async('string')))).join('\n');
    assert(!/◆[A-Z0-9_]+◆/.test(allSections),'모든 공문 키 치환');
    assert(allSections.includes(eligible.name.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')),'선택 공연이 한글에 바인딩');
    assert(allSections.includes('행사별 감염병 판단 및 예방관리 요약'));
    assert(allSections.includes('군중행사 감염병 완화요인 회신 체크리스트'));
  }
  assert.equal(vm.runInContext(`noticeHwpxXml('<a & "b">')`,ctx),'&lt;a &amp; &quot;b&quot;&gt;');
  console.log(JSON.stringify({ok:true,counts,approvedCssPreserved:true,seatFilterChanged:filtered.length<unfiltered,selectedNoticeTested:!!eligible,hwpxGenerated:!!downloadedBlob},null,2));
})().catch(error=>{console.error(error);process.exitCode=1});
