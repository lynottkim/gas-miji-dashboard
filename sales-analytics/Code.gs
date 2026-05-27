// ============================================================
//  미지세계 매출 분석 대시보드  —  Code.gs  v6.2
// ============================================================
// ★★ 배포 전환 — 아래 한 줄만 바꾸면 됩니다 ★★
var DEPLOY_MODE = 'sena';    // 'sena' | 'multi'
// var DEPLOY_MODE = 'multi'; // ← 멀티브랜드 배포 시 위 줄과 교체

// ── 공통 설정 (양쪽 배포 동일) ──────────────────────────────
var VERSION = 'v6.2';

// ── SENA 배포 설정 ───────────────────────────────────────────
var CONFIG_SENA = {
  MODE:           'sena',
  VERSION:        VERSION,
  SPREADSHEET_ID: '1qRyacofsX9QP5aUdioKVG5-6VTp_WUfvqozePkK4PmM',
  SHEET_PREFIX:   'salesrecord',
  CATEGORY_SHEET: 'CATEGORY_MAP',
  DASHBOARD_TITLE:'SENA 영업 대시보드',
  BRAND_COL:      17,
  DATA_START_ROW: 3,
  COL: { DATE:2, SALE_NO:3, TX_TYPE:4, GROUP:5, SKU:6, ITEM_NAME:7, QTY:8,
         NOTE:9, UNIT_PRC:10, SUPPLY:11, VAT:12, TOTAL:13, CUSTOMER:14 }
};

// ── 멀티브랜드 배포 설정 ────────────────────────────────────
var CONFIG_MULTI = {
  MODE:           'multi',
  VERSION:        VERSION,
  SPREADSHEET_ID: '1Mm5Tm-UZrs78t7gYDHiNbCoZVJRdSIBn8lfvCw-qpK4',
  SHEET_PREFIX:   'salesrecord',
  CATEGORY_SHEET: 'CATEGORY_MAP',
  DASHBOARD_TITLE:'멀티브랜드 영업 대시보드',
  BRAND_COL:      17,
  DATA_START_ROW: 3,
  COL: { DATE:2, SALE_NO:3, TX_TYPE:4, GROUP:5, SKU:6, ITEM_NAME:7, QTY:8,
         NOTE:9, UNIT_PRC:10, SUPPLY:11, VAT:12, TOTAL:13, CUSTOMER:14 }
};

// ── 활성 CONFIG (DEPLOY_MODE에 따라 자동 선택) ──────────────
var CONFIG = (DEPLOY_MODE === 'multi') ? CONFIG_MULTI : CONFIG_SENA;

// ── 브랜드 분류 (multi 모드) ───────────────────────────────
// Q열 품목코드 앞 2자리로 브랜드 결정
// 빈값·SK·SENA → SK(SENA), KM→KLIM, TT→TOURATECH, BB→BARKBUSTERS, SB→SCHUBERTH
// SC→배송비(브랜드 분류 없음), 그 외→MISC
var BRAND_MAP = {
  'SK': 'SENA', 'KM': 'KLIM', 'TT': 'TOURATECH',
  'BB': 'BARKBUSTERS', 'SB': 'SCHUBERTH', 'SC': 'SC'
};

function getBrandFromCode_(code) {
  if (!code || code.trim() === '') return 'SENA';
  var s = code.trim().toUpperCase();
  if (s.startsWith('SENA')) return 'SENA';
  var prefix = s.slice(0, 2);
  return BRAND_MAP[prefix] || 'MISC';
}

// ── 진입점 ─────────────────────────────────────────────────
function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  var file  = (page === 'help') ? 'help' : 'index';
  var title = (page === 'help')
    ? CONFIG.DASHBOARD_TITLE + ' — 사용 가이드'
    : CONFIG.DASHBOARD_TITLE;
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSS_() { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function getScriptUrl() { return ScriptApp.getService().getUrl(); }

// ── 연결 테스트 ────────────────────────────────────────────
function testConnection() {
  try {
    var ss=getSS_(); var tabs=ss.getSheets().map(function(s){return s.getName();});
    var re=new RegExp('^'+CONFIG.SHEET_PREFIX+'\\d{4}$');
    return {ok:true,spreadsheetName:ss.getName(),allTabs:tabs,
            dataTabs:tabs.filter(function(n){return re.test(n);})};
  } catch(e){return {ok:false,error:e.message};}
}

// ── 메타데이터 ─────────────────────────────────────────────
// 최적화: 추출일=P1셀, 거래처/유형=현재연도 2컬럼만
function getMetadata() {
  try {
    var ss=getSS_();
    var re=new RegExp('^'+CONFIG.SHEET_PREFIX+'(\\d{4})$');
    var years=[];
    ss.getSheets().forEach(function(s){var m=s.getName().match(re);if(m)years.push(m[1]);});
    years.sort().reverse();
    if(!years.length) return {ok:false,error:'데이터 탭 없음'};

    // ① 추출일: 각 탭 P1(16열)
    var yearMeta={};
    years.forEach(function(yr){
      var sh=ss.getSheetByName(CONFIG.SHEET_PREFIX+yr); if(!sh) return;
      var v=String(sh.getRange(1,16,1,1).getValue()).trim();
      yearMeta[yr]={extractDate:/^D\d{2}\.\d{2}\.\d{2}/.test(v)?v:''};
    });

    // ② 거래처·거래유형: 현재연도 D·N 2컬럼만
    var SKIP={'거래처명':1,'거래유형명':1,'':1};
    var customers={},txTypes={};
    var baseSh=ss.getSheetByName(CONFIG.SHEET_PREFIX+years[0]);
    if(baseSh){
      var lastRow=baseSh.getLastRow();
      var dataRows=lastRow-CONFIG.DATA_START_ROW+1;
      if(dataRows>0){
        baseSh.getRange(CONFIG.DATA_START_ROW,CONFIG.COL.TX_TYPE,dataRows,1).getValues()
          .forEach(function(r){var t=String(r[0]).trim();if(!SKIP[t])txTypes[t]=1;});
        baseSh.getRange(CONFIG.DATA_START_ROW,CONFIG.COL.CUSTOMER,dataRows,1).getValues()
          .forEach(function(r){var c=String(r[0]).trim();if(!SKIP[c])customers[c]=1;});
      }
    }

    var catData=loadCategoryMap_(ss);
    var userEmail='';
    try{ userEmail=Session.getActiveUser().getEmail()||''; }catch(e){}
    return {
      ok:true, years:years, yearMeta:yearMeta,
      customers:Object.keys(customers).sort(),
      txTypes:Object.keys(txTypes).sort(),
      categoryMap:catData.map, primaryCats:catData.primaries,
      tags:catData.tags, skuMeta:catData.skuMeta,
      mode:           CONFIG.MODE,
      dashboardTitle: CONFIG.DASHBOARD_TITLE,
      version:        CONFIG.VERSION||'',
      userEmail:      userEmail,
      scriptUrl:      ScriptApp.getService().getUrl()
    };
  } catch(e){return {ok:false,error:e.message+'\n'+e.stack};}
}

// ── 추가 연도 프리셋 로드 ──────────────────────────────────
function loadYearPreset(years){
  try{
    var ss=getSS_(); var SKIP={'거래처명':1,'거래유형명':1,'':1};
    var customers={},txTypes={};
    years.forEach(function(yr){
      var sh=ss.getSheetByName(CONFIG.SHEET_PREFIX+yr); if(!sh) return;
      var lastRow=sh.getLastRow(); var dataRows=lastRow-CONFIG.DATA_START_ROW+1;
      if(dataRows<=0) return;
      sh.getRange(CONFIG.DATA_START_ROW,CONFIG.COL.TX_TYPE,dataRows,1).getValues()
        .forEach(function(r){var t=String(r[0]).trim();if(!SKIP[t])txTypes[t]=1;});
      sh.getRange(CONFIG.DATA_START_ROW,CONFIG.COL.CUSTOMER,dataRows,1).getValues()
        .forEach(function(r){var c=String(r[0]).trim();if(!SKIP[c])customers[c]=1;});
    });
    return {ok:true,customers:Object.keys(customers).sort(),txTypes:Object.keys(txTypes).sort()};
  } catch(e){return {ok:false,error:e.message};}
}

// ── 데이터 조회 & 집계 ─────────────────────────────────────
function getFilteredData(filters) {
  try {
    var ss=getSS_(); var catData=loadCategoryMap_(ss);
    var skuMeta=catData.skuMeta;
    var rows=loadRows_(filters,skuMeta);

    var effectiveGroupBy = filters.groupBy || 'sku';

    // 타임컬럼 정의: 선택 연도 × 기간으로 자동 결정
    // 1개년=월별, 2개년=월별(가로스크롤), 3~4개년=분기별, 5개년+=연간
    var timeColsResult = buildTimeCols_(
      filters.years||[], filters.months||[],
      filters.dateFrom||'', filters.dateTo||''
    );
    var timeCols      = timeColsResult.cols;       // [{key,label,months:[]}]
    var timeSpanLabel = timeColsResult.spanLabel;  // 표시용 기간 문자열
    var flatMonths    = timeColsResult.flatMonths; // 집계 대상 월 목록

    // 메인 집계: 비-SKU 뷰는 그룹 행에도 타임컬럼 합계를 담아 아코디언과 그리드 일치
    var agg        = (effectiveGroupBy!=='sku')
                       ? aggregate_(rows, effectiveGroupBy, timeCols)
                       : aggregate_(rows, effectiveGroupBy);
    var monthly    = aggregateMonthly_(rows);
    var txBreak    = aggregateByField_(rows,'txType');
    var catBreak   = aggregateBySkuMeta_(rows,skuMeta,'primaryCat');
    var skuAgg     = aggregate_(rows,'sku');
    var skuMonthly = aggregateSkuMonthlyByTimeCols_(rows, timeCols);
    var brandBreak = (CONFIG.MODE==='multi') ? aggregateByField_(rows,'brand') : null;
    var familyBreak= aggregateByField_(rows,'familyTag');
    // 교차 데이터: SKU뷰 외 모든 뷰(시리즈/거래처/월/거래유형)는 아코디언용 cross 생성
    var crossData = null;
    if(effectiveGroupBy!=='sku'){
      crossData = aggregateCross_(rows, effectiveGroupBy, timeCols);
    }

    // 연도별 breakdown: 그룹 행의 연도별 판매수량·매출액·평균단가 이중 표시용
    // yearBreak[groupKey] = [{year, saleQty, supply, avgPrice}, ...]
    var yearBreak = null;
    var selectedYears = (filters.years||[]).slice().sort();
    if(selectedYears.length>=2 && effectiveGroupBy!=='sku'){
      yearBreak = aggregateYearBreak_(rows, effectiveGroupBy, selectedYears);
    }

    // ── 미분류 SKU 감지 ──────────────────────────────────────
    var mappedSkuSet = skuMeta;
    var unmappedMap  = {};
    rows.forEach(function(r){
      if(!mappedSkuSet[r.sku]){
        if(!unmappedMap[r.sku]) unmappedMap[r.sku]={
          sku:r.sku, itemName:r.itemName, brand:r.brand||'', supply:0, qty:0
        };
        unmappedMap[r.sku].supply+=r.supply;
        unmappedMap[r.sku].qty   +=r.qty;
      }
    });
    var unmappedSkus=Object.keys(unmappedMap)
      .map(function(k){return unmappedMap[k];})
      .sort(function(a,b){return b.supply-a.supply;});

    var ssUrl='https://docs.google.com/spreadsheets/d/'+CONFIG.SPREADSHEET_ID+'/edit';

    // ── 디버그 정보 (개발 편의용) ────────────────────────────
    var debugInfo={
      totalRows: rows.length,
      skuMetaCount: Object.keys(skuMeta).length,
      filtersApplied: {
        brands: filters.brands||[],
        primaryCat: filters.primaryCat||'',
        years: filters.years||[]
      },
      sampleSkus: rows.slice(0,3).map(function(r){return {sku:r.sku,brand:r.brand};}),
      sampleMapKeys: Object.keys(skuMeta).slice(0,3)
    };

    return {ok:true, data:agg, totalRows:rows.length,
            monthly:monthly, txBreak:txBreak, catBreak:catBreak,
            skuAgg:skuAgg, skuMonthly:skuMonthly,
            crossData:crossData,
            timeCols:timeCols, timeSpanLabel:timeSpanLabel,
            yearBreak:yearBreak, selectedYears:selectedYears,
            brandBreak:brandBreak, familyBreak:familyBreak,
            unmappedSkus:unmappedSkus, ssUrl:ssUrl,
            debugInfo:debugInfo};
  } catch(e){return {ok:false,error:e.message+'\n'+e.stack};}
}

// ── 전년도 월별 ────────────────────────────────────────────
function getPrevYearMonthly(filters) {
  try {
    var ss=getSS_(); var skuMeta=loadCategoryMap_(ss).skuMeta;
    var pf=JSON.parse(JSON.stringify(filters));
    pf.years=filters.years.map(function(y){return String(Number(y)-1);});
    var rows=loadRows_(pf,skuMeta);
    return {ok:true,monthly:aggregateMonthly_(rows)};
  } catch(e){return {ok:false,error:e.message};}
}

// ── 발주 예측 ─────────────────────────────────────────────
function getOrderForecast(filters) {
  try {
    var ss=getSS_(); var skuMeta=loadCategoryMap_(ss).skuMeta;
    var ff=JSON.parse(JSON.stringify(filters)); ff.years=[];
    var rows=loadRows_(ff,skuMeta);
    if(!rows.length) return {ok:true,data:[],baseDate:''};

    var maxDate=rows.reduce(function(mx,r){return r.date>mx?r.date:mx;},'2000-01-01');
    var baseDate=new Date(maxDate);
    var SAFETY=1.2;

    // SKU별 월별 맵
    var skuMap={};
    rows.forEach(function(r){
      if(!skuMap[r.sku]) skuMap[r.sku]={months:{},itemName:r.itemName};
      var k=r.date.slice(0,7);
      skuMap[r.sku].months[k]=(skuMap[r.sku].months[k]||0)+r.qty;
    });

    var W={3:[.50,.30,.20],6:[.25,.20,.18,.15,.12,.10],
           12:[.18,.14,.12,.10,.09,.08,.07,.06,.05,.05,.04,.02]};

    function calc(mmap,nMonths,seaMons){
      var wts=W[nMonths]||W[3]; var pts=[];
      for(var i=0;i<nMonths;i++){
        var d=new Date(baseDate); d.setMonth(d.getMonth()-i);
        var mon=d.getMonth()+1;
        if(seaMons.length&&seaMons.indexOf(mon)===-1) continue;
        var k=d.getFullYear()+'-'+(mon<10?'0':'')+mon;
        pts.push({qty:mmap[k]||0,w:wts[pts.length]||(1/nMonths)});
      }
      if(!pts.length) return null;
      var sw=0,swq=0;
      pts.forEach(function(p){sw+=p.w;swq+=p.qty*p.w;});
      var wavg=sw>0?swq/sw:0;
      // 선형 트렌드
      var n=pts.length,sx=0,sy=0,sxy=0,sx2=0;
      pts.forEach(function(p,i){sx+=i;sy+=p.qty;sxy+=i*p.qty;sx2+=i*i;});
      var slope=n>1?(n*sxy-sx*sy)/(n*sx2-sx*sx):0;
      var tr=wavg>0?Math.max(-.5,Math.min(.5,slope/wavg)):0;
      return {wavg:Math.round(wavg),trend:Math.round(tr*1000)/10,
              adj:Math.max(0,Math.round(wavg*(1+tr))),count:pts.length};
    }

    // 현재월 기준 3개월 이상 판매 없는 SKU 제외 (단종 또는 판매 저조)
    var now=new Date();
    var curMk=now.getFullYear()+'-'+(now.getMonth()<9?'0':'')+(now.getMonth()+1);
    var threeMonthsAgo=new Date(now.getFullYear(),now.getMonth()-3,1);
    var cutoffMk=threeMonthsAgo.getFullYear()+'-'+(threeMonthsAgo.getMonth()<9?'0':'')+(threeMonthsAgo.getMonth()+1);

    var result=Object.keys(skuMap).map(function(sku){
      var meta=skuMeta[sku]||{};
      var seaMons=parseSeasonalMonths_(meta.seasonal||'');
      var mmap=skuMap[sku].months;
      var allKeys=Object.keys(mmap).sort();
      var lastSaleMk=allKeys.length?allKeys[allKeys.length-1]:'';
      // 3개월 이상 미판매 → 발주예측 불필요, null 반환으로 제외
      if(lastSaleMk && lastSaleMk < cutoffMk) return null;
      var dataMonths=allKeys.length;
      var r3 =dataMonths>=1?calc(mmap,3,seaMons):null;
      var r6 =dataMonths>=3?calc(mmap,6,seaMons):null;
      var r12=dataMonths>=6?calc(mmap,12,seaMons):null;
      var best=r12||r6||r3;
      var recommend=null,recBasis='';
      if(best){
        recommend=Math.ceil(best.adj*SAFETY);
        recBasis=(r12?'12개월':r6?'6개월':'3개월')+
          (best.trend>0?' ↑+'+best.trend+'%':best.trend<0?' ↓'+best.trend+'%':' →')+
          ' ×'+SAFETY;
      }
      return {sku:sku,itemName:skuMap[sku].itemName||'',
              primaryCat:meta.primaryCat||'',
              seasonal:meta.seasonal?'🌨 시즌('+meta.seasonal+')':'',
              dataMonths:dataMonths,firstMonth:allKeys[0]||'',
              avg3:r3?r3.wavg:null,trend3:r3?r3.trend:null,
              avg6:r6?r6.wavg:null,trend6:r6?r6.trend:null,
              avg12:r12?r12.wavg:null,trend12:r12?r12.trend:null,
              fc1:best?best.adj:null,fc3:best?best.adj*3:null,
              recommend:recommend,recBasis:recBasis};
    }).filter(function(r){return r!==null;});

    if(filters.skus&&filters.skus.length)
      result=result.filter(function(r){return filters.skus.indexOf(r.sku)!==-1;});
    if(filters.primaryCat)
      result=result.filter(function(r){return r.primaryCat===filters.primaryCat;});
    result.sort(function(a,b){return b.dataMonths-a.dataMonths;});
    return {ok:true,data:result,baseDate:maxDate};
  } catch(e){return {ok:false,error:e.message+'\n'+e.stack};}
}

function parseSeasonalMonths_(s){
  if(!s) return [];
  var p=s.split('-'); if(p.length!==2) return [];
  var from=parseInt(p[0]),to=parseInt(p[1]),m=[];
  if(from<=to){for(var i=from;i<=to;i++)m.push(i);}
  else{for(var i=from;i<=12;i++)m.push(i);for(var i=1;i<=to;i++)m.push(i);}
  return m;
}

// series(tags[2]) 추출 — MAP의 series 컬럼을 쓰므로 단순 passthrough
// tags 배열에서 Family 추출이 필요할 때의 fallback용
function getFamilyTag_(tags){
  if(!tags||!tags.length) return '';
  var SKIP={'All Weather':1,'Summer':1,'Transition':1,'Winter':1,'Spring':1,
            'Fall':1,'n/a':1,'KLIM':1,'Woman':1,'Goretex':1,'Mesh':1,
            'ADV':1,'Touring':1,'Off-road':1,'Lifestyle':1,'Street':1};
  for(var i=2;i<tags.length;i++){
    if(tags[i]&&!SKIP[tags[i]]) return tags[i];
  }
  // fallback: Style(index 1)
  if(tags[1]&&tags[1]!=='KLIM'&&tags[1]!=='n/a') return tags[1];
  return '';
}

// ── 직전 N개월 키 목록 (yyyy-MM, 오름차순) ────────────────
function buildRecentMonths_(n){
  var out=[]; var now=new Date();
  for(var i=n-1;i>=0;i--){
    var d=new Date(now.getFullYear(), now.getMonth()-i, 1);
    var mk=d.getFullYear()+'-'+((d.getMonth()+1)<10?'0':'')+(d.getMonth()+1);
    out.push(mk);
  }
  return out;
}

// ── 타임컬럼 정의 생성 ───────────────────────────────────────
// 선택 연도 × 기간 프리셋(월 배열) → {cols, unit, spanLabel, flatMonths}
// cols: [{key, label, months:[yyyy-MM,...]}]
// unit: 'month' | 'quarter' | 'year'
// spanLabel: 표시용 기간 문자열 (예: "2023-01 ~ 2024-12 · 월별")
// flatMonths: 모든 cols의 months를 평탄화한 배열 (crossData 필터용)
function buildTimeCols_(years, selMonths, dateFrom, dateTo){
  // 유효 연도가 없으면 최근 12개월 기본값
  if(!years||!years.length){
    var recent=buildRecentMonths_(12);
    var cols=recent.map(function(mk){
      return {key:mk, label:mk.slice(5)+'월', months:[mk]};
    });
    return {cols:cols, unit:'month', spanLabel:'최근 12개월', flatMonths:recent};
  }

  var sortedYears=years.slice().sort();
  var nYears=sortedYears.length;

  // 유효 월 필터 (기간 프리셋 없으면 전월 포함)
  var activeMons=selMonths&&selMonths.length?selMonths:
    [1,2,3,4,5,6,7,8,9,10,11,12];

  // 날짜 직접 입력이 있으면 해당 범위로 덮어씀
  if(dateFrom&&dateTo){
    var dfY=parseInt(dateFrom.slice(0,4)), dfM=parseInt(dateFrom.slice(5,7));
    var dtY=parseInt(dateTo.slice(0,4)), dtM=parseInt(dateTo.slice(5,7));
    var rangeMonths=[];
    for(var y=dfY;y<=dtY;y++){
      var mStart=(y===dfY?dfM:1), mEnd=(y===dtY?dtM:12);
      for(var m=mStart;m<=mEnd;m++){
        rangeMonths.push(y+'-'+(m<10?'0':'')+m);
      }
    }
    var cols=rangeMonths.map(function(mk){
      return {key:mk, label:mk.slice(2,4)+'-'+mk.slice(5)+'월', months:[mk]};
    });
    var span=dateFrom+' ~ '+dateTo+' · 월별';
    return {cols:cols, unit:'month', spanLabel:span, flatMonths:rangeMonths};
  }

  // 단위 결정: 연도 수 기준
  // 1개년: 월별 / 2개년: 월별(가로 스크롤) / 3~4개년: 분기별 / 5개년+: 연간
  var unit= nYears<=2?'month': nYears<=4?'quarter':'year';

  var cols=[], flatMonths=[];
  var minY=parseInt(sortedYears[0]), maxY=parseInt(sortedYears[sortedYears.length-1]);

  if(unit==='month'){
    for(var y=minY;y<=maxY;y++){
      var yr=String(y);
      if(sortedYears.indexOf(yr)===-1) continue;
      for(var mi=0;mi<activeMons.length;mi++){
        var m=activeMons[mi];
        var mk=y+'-'+(m<10?'0':'')+m;
        // 항상 YYMM 형태 (1개년: "2505", 2개년: "2505" — 분기 "24Q4"와 일관성)
        var lbl=yr.slice(2)+(m<10?'0':'')+m;
        cols.push({key:mk, label:lbl, months:[mk]});
        flatMonths.push(mk);
      }
    }
  } else if(unit==='quarter'){
    // 분기별: 활성 월을 분기로 그룹핑
    var qDef=[[1,2,3],[4,5,6],[7,8,9],[10,11,12]];
    for(var y=minY;y<=maxY;y++){
      var yr=String(y);
      if(sortedYears.indexOf(yr)===-1) continue;
      for(var q=0;q<4;q++){
        var qMons=qDef[q].filter(function(m){return activeMons.indexOf(m)!==-1;});
        if(!qMons.length) continue;
        var qMonKeys=qMons.map(function(m){return y+'-'+(m<10?'0':'')+m;});
        cols.push({key:yr+'Q'+(q+1), label:yr.slice(2)+'Q'+(q+1), months:qMonKeys});
        qMonKeys.forEach(function(mk){flatMonths.push(mk);});
      }
    }
  } else {
    // 연간: 연도별 칸
    for(var yi=0;yi<sortedYears.length;yi++){
      var y=parseInt(sortedYears[yi]);
      var yMons=activeMons.map(function(m){return y+'-'+(m<10?'0':'')+m;});
      cols.push({key:String(y), label:String(y)+'년', months:yMons});
      yMons.forEach(function(mk){flatMonths.push(mk);});
    }
  }

  // spanLabel 생성
  var unitStr=unit==='month'?'월별':unit==='quarter'?'분기별':'연간';
  var periodStr=selMonths&&selMonths.length?
    selMonths.length===12?'전체':selMonths[0]+'~'+selMonths[selMonths.length-1]+'월':'전체';
  var spanLabel=sortedYears.join(', ')+'년'+(periodStr!=='전체'?' · '+periodStr:'')+' · '+unitStr;

  return {cols:cols, unit:unit, spanLabel:spanLabel, flatMonths:flatMonths};
}

// ── 시리즈 키 결정 ─────────────────────────────────────────
// ① MAP에 series값 있으면 그것 → ② 없으면 품번 앞자리(- 앞, 또는 KLIM 4자리)
// → ③ 그것도 없으면 SKU 그대로(단독 시리즈)
function resolveSeriesKey_(sku, meta){
  if(meta && meta.series) return meta.series;
  if(!sku) return sku;
  // KM-3610-... → 3610  (KLIM 품번)
  var km=sku.match(/^KM-?(\d{4})/);
  if(km) return km[1];
  // 일반: 첫 '-' 앞부분 (50S-10 → 50S)
  if(sku.indexOf('-')!==-1){
    var head=sku.split('-')[0];
    if(head) return head;
  }
  return sku; // 단독
}

// ── 행 로드 ───────────────────────────────────────────────
function loadRows_(filters,skuMeta){
  var ss=getSS_();
  var re=new RegExp('^'+CONFIG.SHEET_PREFIX+'(\\d{4})$');
  var allYears=ss.getSheets()
    .map(function(s){return s.getName();}).filter(function(n){return re.test(n);})
    .map(function(n){return n.replace(CONFIG.SHEET_PREFIX,'');}).sort().reverse();
  var targetYears=(filters.years&&filters.years.length)?filters.years:allYears;

  // 태그·주카테고리 → SKU 집합 변환
  var tagSkus=null;
  if(filters.tags&&filters.tags.length){
    tagSkus=[];
    Object.keys(skuMeta).forEach(function(sku){
      if(filters.tags.some(function(t){return skuMeta[sku].tags&&skuMeta[sku].tags.indexOf(t)!==-1;}))
        tagSkus.push(sku);
    });
  }
  var primarySkus=null;
  if(filters.primaryCat){
    primarySkus=[];
    Object.keys(skuMeta).forEach(function(sku){
      if(skuMeta[sku].primaryCat===filters.primaryCat) primarySkus.push(sku);
    });
  }

  var extractPat=/^D\d{2}\.\d{2}\.\d{2}/;
  var isMulti=(CONFIG.MODE==='multi');
  // multi 모드: 읽을 컬럼 수를 Q열(17)까지 확장
  var readCols=isMulti?Math.max(17,14):14;
  var rows=[];
  targetYears.forEach(function(yr){
    var sh=ss.getSheetByName(CONFIG.SHEET_PREFIX+yr); if(!sh) return;
    var lastRow=sh.getLastRow(); if(lastRow<CONFIG.DATA_START_ROW) return;
    var vals=sh.getRange(CONFIG.DATA_START_ROW,1,lastRow-CONFIG.DATA_START_ROW+1,readCols).getValues();
    var emptyCount=0;
    for(var i=0;i<vals.length;i++){
      var r=vals[i]; var dateVal=r[CONFIG.COL.DATE-1];
      if(!dateVal){if(++emptyCount>=5)break;continue;} emptyCount=0;
      if(extractPat.test(String(dateVal).trim())) continue;
      var d=(dateVal instanceof Date)?dateVal:new Date(dateVal);
      if(isNaN(d.getTime())) continue;
      if(filters.dateFrom&&d<new Date(filters.dateFrom)) continue;
      if(filters.dateTo&&d>new Date(filters.dateTo+'T23:59:59')) continue;
      if(filters.months&&filters.months.length&&filters.months.indexOf(d.getMonth()+1)===-1) continue;
      var sku=String(r[CONFIG.COL.SKU-1]).trim();
      var customer=String(r[CONFIG.COL.CUSTOMER-1]).trim();
      var txType=String(r[CONFIG.COL.TX_TYPE-1]).trim();
      var saleNo=String(r[CONFIG.COL.SALE_NO-1]).trim();

      // multi 모드: 브랜드 분류
      var brand='';
      var seriesId='';
      if(isMulti){
        var internalCode=r.length>=CONFIG.BRAND_COL?String(r[CONFIG.BRAND_COL-1]).trim():'';
        brand=getBrandFromCode_(internalCode);
        // 브랜드 필터: 코드(KM)와 이름(KLIM) 모두 허용
        if(filters.brands&&filters.brands.length){
          var brandMatch=filters.brands.some(function(b){
            // 코드 형태(KM,TT,BB,SB,MISC)이면 이름으로 변환해서 비교
            var nameFromCode={KM:'KLIM',TT:'TOURATECH',BB:'BARKBUSTERS',SB:'SCHUBERTH',MISC:'MISC'};
            var bName=nameFromCode[b]||b;
            return brand===bName||brand===b;
          });
          if(!brandMatch) continue;
        }
        if(brand==='KLIM'&&sku.indexOf('-')!==-1) seriesId=sku.split('-')[0];
      }

      if(filters.skus&&filters.skus.length&&filters.skus.indexOf(sku)===-1) continue;
      if(filters.customers&&filters.customers.length&&filters.customers.indexOf(customer)===-1) continue;
      if(filters.txTypes&&filters.txTypes.length&&filters.txTypes.indexOf(txType)===-1) continue;
      if(tagSkus!==null&&tagSkus.indexOf(sku)===-1) continue;
      if(primarySkus!==null&&primarySkus.indexOf(sku)===-1) continue;
      var meta=skuMeta[sku]||null;
      // 시리즈 묶음 키: ① MAP series값 → ② 품번 앞자리 폴백 → ③ SKU 단독
      var seriesKey=resolveSeriesKey_(sku, meta);
      var seriesLabel=(meta&&meta.series)?meta.series:seriesKey;
      rows.push({
        date:Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd'),
        month:d.getMonth()+1,year:d.getFullYear(),saleNo:saleNo,txType:txType,
        sku:sku,itemName:String(r[CONFIG.COL.ITEM_NAME-1]).trim(),
        qty:Number(r[CONFIG.COL.QTY-1])||0,supply:Number(r[CONFIG.COL.SUPPLY-1])||0,
        customer:customer,
        brand:brand,seriesId:seriesId,
        // CATEGORY_MAP에서 가져온 분류 정보
        primaryCat: meta?meta.primaryCat:'',
        familyTag:  meta?(meta.series||getFamilyTag_(meta.tags)):'',
        seriesKey: seriesKey,       // 시리즈뷰 그룹 키
        seriesLabel: seriesLabel    // 시리즈뷰 표시명
      });
    }
  });
  return rows;
}

// ── 집계 함수들 ────────────────────────────────────────────
function makeGroup_(key,r,nMonths){
  return {key:key,sku:r.sku,itemName:r.itemName,
          shipQty:0,retQty:0,saleQty:0,supply:0,saleNos:{},
          catSet:{},skuSet:{},
          itemNames:[],  // 시리즈뷰 대표규격명 계산용
          monthly:[]};
}
function finalizeGroup_(g){
  var tx=Object.keys(g.saleNos).length;
  g.txCount=tx;
  g.retRate=g.shipQty>0?Math.round(Math.abs(g.retQty)/g.shipQty*1000)/10:0;
  g.avgPrice=g.saleQty>0?Math.round(g.supply/g.saleQty):0;
  g.skuCount=Object.keys(g.skuSet).length;
  // 대표규격명: 하위 SKU itemName들의 공통 앞부분
  g.repName=g.itemNames.length?commonPrefix_(g.itemNames):'';
  var cs=g.catSet;
  g.catSummary=Object.keys(cs).sort(function(a,b){return cs[b]-cs[a];})
    .slice(0,4).map(function(k){return k+'×'+cs[k];}).join(', ');
  delete g.saleNos; delete g.catSet; delete g.skuSet; delete g.itemNames; return g;
}
function addRow_(g,r,timeIdx){
  if(r.qty>=0)g.shipQty+=r.qty;else g.retQty+=r.qty;
  g.saleQty+=r.qty; g.supply+=r.supply;
  if(r.saleNo)g.saleNos[r.saleNo]=1;
  if(r.primaryCat) g.catSet[r.primaryCat]=(g.catSet[r.primaryCat]||0)+1;
  // 대표규격명용: SKU당 itemName 하나만 수집 (skuSet으로 중복 방지)
  if(r.itemName && !g.skuSet[r.sku] && g.itemNames.length<50)
    g.itemNames.push(r.itemName);
  // 타임컬럼별 판매수량 누적 (timeCols 기반)
  if(timeIdx){
    var mk=r.year+'-'+(r.month<10?'0':'')+r.month;
    var ci=timeIdx[mk];
    if(ci!==undefined){
      if(!g.monthly.length){ for(var i=0;i<timeIdx._n;i++) g.monthly.push(0); }
      g.monthly[ci]+=r.qty;
    }
  }
}

// timeCols → month→colIndex 역매핑 생성
function buildTimeIdx_(timeCols){
  if(!timeCols||!timeCols.length) return null;
  var idx={_n:timeCols.length};
  timeCols.forEach(function(col,ci){
    col.months.forEach(function(mk){ idx[mk]=ci; });
  });
  return idx;
}

function aggregate_(rows,groupBy,timeCols){
  var timeIdx=buildTimeIdx_(timeCols);
  var groups={};
  rows.forEach(function(r){
    var key=groupBy==='customer' ? r.customer
           :groupBy==='month'   ? r.year+'-'+(r.month<10?'0':'')+r.month
           :groupBy==='txtype'  ? r.txType
           :groupBy==='brand'   ? r.brand
           :groupBy==='series'  ? r.seriesKey
           :groupBy==='family'  ? (r.familyTag||r.sku)
           :r.sku;
    if(!groups[key]){
      groups[key]=makeGroup_(key,r);
      if(groupBy==='series') groups[key].key=r.seriesLabel||key;
    }
    addRow_(groups[key],r,timeIdx);
    if(groupBy==='series') groups[key].skuSet[r.sku]=1;
  });
  return Object.keys(groups).map(function(k){return finalizeGroup_(groups[k]);})
    .sort(function(a,b){return b.supply-a.supply;});
}

// 통합 교차 집계: { groupKey: [ {sku,itemName,...,avgPrice, monthly[], pctInGroup}, ... ] }
// 모든 비-SKU 뷰(시리즈/거래처/월/거래유형)의 아코디언에서 하위 SKU 목록 표시용.
// 월별 판매수량과 그룹 내 비중까지 포함하여 SKU탭과 동일 그리드를 만들 수 있게 함.
function aggregateCross_(rows,groupBy,timeCols){
  function groupKeyOf(r){
    return groupBy==='customer'?r.customer
         :groupBy==='month'?r.year+'-'+(r.month<10?'0':'')+r.month
         :groupBy==='txtype'?r.txType
         :groupBy==='series'?(r.seriesLabel||r.seriesKey)
         :r.sku;
  }
  var timeIdx=buildTimeIdx_(timeCols);
  var nCols=(timeCols||[]).length;
  var cross={};
  rows.forEach(function(r){
    var gk=groupKeyOf(r);
    if(!cross[gk])cross[gk]={};
    if(!cross[gk][r.sku])cross[gk][r.sku]={sku:r.sku,itemName:r.itemName,
      shipQty:0,retQty:0,saleQty:0,supply:0,
      monthly:new Array(nCols).fill(0)};
    var g=cross[gk][r.sku];
    if(r.qty>=0)g.shipQty+=r.qty;else g.retQty+=r.qty;
    g.saleQty+=r.qty;g.supply+=r.supply;
    if(timeIdx){
      var mk=r.year+'-'+(r.month<10?'0':'')+r.month;
      var ci=timeIdx[mk];
      if(ci!==undefined) g.monthly[ci]+=r.qty;
    }
  });
  var result={};
  Object.keys(cross).forEach(function(gk){
    var arr=Object.keys(cross[gk]).map(function(sk){
      var g=cross[gk][sk];
      g.retRate=g.shipQty>0?Math.round(Math.abs(g.retQty)/g.shipQty*1000)/10:0;
      g.avgPrice=g.saleQty>0?Math.round(g.supply/g.saleQty):0;
      return g;
    }).sort(function(a,b){return b.supply-a.supply;});
    // 그룹 내 비중 (매출 기준)
    var gtot=arr.reduce(function(s,x){return s+x.supply;},0);
    arr.forEach(function(x){ x.pctInGroup=gtot>0?Math.round(x.supply/gtot*1000)/10:0; });
    result[gk]=arr;
  });
  return result;
}

// 여러 itemName의 공통 앞부분 추출 (색/사이즈 제거 → 대표규격명)
// 예: ["Krios Pro SM Black", "Krios Pro MD White", "Krios Pro LG Red"] → "Krios Pro"
function commonPrefix_(names){
  if(!names||!names.length) return '';
  if(names.length===1) return names[0];
  // 컬러/사이즈 키워드 목록
  var STOP=/\b(SM|MD|LG|XL|XXL|XS|2XL|3XL|S|M|L|[0-9]+|Black|White|Red|Blue|Green|Gray|Grey|Orange|Yellow|Purple|Pink|Brown|Silver|Gold|Carbon|Camo|Hi.?Vis|Charger|Electric|Matte|Gloss|ECE|DOT|EU|US|TU|One|Size|Petrol|Lime|Teal|Navy)\b/i;
  // 토큰 단위로 자름
  function tokenize(s){ return (s||'').split(/\s+/); }
  var toks=names.map(tokenize);
  var ref=toks[0]; var common=[];
  for(var i=0;i<ref.length;i++){
    var t=ref[i];
    if(STOP.test(t)) break;
    var allMatch=toks.every(function(ts){return ts[i]===t;});
    if(!allMatch) break;
    common.push(t);
  }
  return common.join(' ');
}

// 연도별 breakdown: 그룹 행의 연도별 판매수량·매출액·평균단가
// 반환: { groupKey: [{year, saleQty, supply, avgPrice}, ...] }
// effectiveGroupBy에 따라 groupKey를 결정
function aggregateYearBreak_(rows, groupBy, years){
  function groupKeyOf(r){
    return groupBy==='customer'?r.customer
         :groupBy==='month'?r.year+'-'+(r.month<10?'0':'')+r.month
         :groupBy==='txtype'?r.txType
         :groupBy==='series'?(r.seriesLabel||r.seriesKey)
         :r.sku;
  }
  // {groupKey: {year: {saleQty, supply}}}
  var acc={};
  rows.forEach(function(r){
    var gk=groupKeyOf(r);
    var yr=String(r.year);
    if(!acc[gk]) acc[gk]={};
    if(!acc[gk][yr]) acc[gk][yr]={saleQty:0, supply:0};
    acc[gk][yr].saleQty+=r.qty;
    acc[gk][yr].supply+=r.supply;
  });
  // 연도 순서대로 배열로 변환 + avgPrice 계산
  var result={};
  Object.keys(acc).forEach(function(gk){
    result[gk]=years.map(function(yr){
      var d=acc[gk][yr]||{saleQty:0,supply:0};
      return {
        year:yr,
        saleQty:d.saleQty,
        supply:d.supply,
        avgPrice:d.saleQty>0?Math.round(d.supply/d.saleQty):0
      };
    });
  });
  // 전체 합계 행용 ('_TOTAL_' 키)
  var totalByYear={};
  rows.forEach(function(r){
    var yr=String(r.year);
    if(!totalByYear[yr]) totalByYear[yr]={saleQty:0,supply:0};
    totalByYear[yr].saleQty+=r.qty;
    totalByYear[yr].supply+=r.supply;
  });
  result['_TOTAL_']=years.map(function(yr){
    var d=totalByYear[yr]||{saleQty:0,supply:0};
    return {
      year:yr,
      saleQty:d.saleQty,
      supply:d.supply,
      avgPrice:d.saleQty>0?Math.round(d.supply/d.saleQty):0
    };
  });
  return result;
}

function aggregateMonthly_(rows){
  var m={};
  rows.forEach(function(r){
    var k=r.year+'-'+(r.month<10?'0':'')+r.month;
    if(!m[k])m[k]={key:k,supply:0,saleQty:0,shipQty:0,retQty:0};
    m[k].supply+=r.supply;m[k].saleQty+=r.qty;
    if(r.qty>=0)m[k].shipQty+=r.qty;else m[k].retQty+=r.qty;
  });
  return Object.keys(m).sort().map(function(k){return m[k];});
}

function aggregateByField_(rows,field){
  var m={};
  rows.forEach(function(r){
    var k=r[field]||'기타';
    if(!m[k])m[k]={key:k,supply:0,saleQty:0};
    m[k].supply+=r.supply;m[k].saleQty+=r.qty;
  });
  return Object.keys(m).map(function(k){return m[k];}).sort(function(a,b){return b.supply-a.supply;});
}

function aggregateBySkuMeta_(rows,skuMeta,field){
  var m={};
  rows.forEach(function(r){
    var meta=skuMeta[r.sku]||{};
    var k=meta[field]||'기타';
    if(!m[k])m[k]={key:k,supply:0,saleQty:0};
    m[k].supply+=r.supply;m[k].saleQty+=r.qty;
  });
  return Object.keys(m).map(function(k){return m[k];}).sort(function(a,b){return b.supply-a.supply;});
}

function buildDefaultCatData_(){
  var map={'배달헤드셋':{label:'배달헤드셋',skus:['K10-01','SK10-01','SMH5-FM-BP-10']},
           '모터사이클헤드셋':{label:'모터사이클헤드셋',skus:['50S-10','60S-01','SPIDER-ST1-10']},
           '모터사이클헬멧':{label:'모터사이클헬멧',skus:['PHANTOM-MB00L3','SURGE-MBGB00M3']},
           '자전거':{label:'자전거',skus:['BIKOM20-02','M1EVO-MB00L10','S1-MB00M']},
           '스노우스포츠':{label:'스노우스포츠',skus:['SNOWTALK2-01','LTS2-MB00L']},
           '헤드셋액세서리':{label:'헤드셋액세서리',skus:['K10-SP01','SC-A0354']},
           '헬멧액세서리':{label:'헬멧액세서리',skus:['SC-A0372','SC-A0373']},
           '소모품':{label:'소모품',skus:['NSP-SC-WMIC-01']},
           '배송비':{label:'배송비',skus:['SC-STF010','SC-SVC010']},
           '서비스파츠':{label:'서비스파츠',skus:['SC010']}};
  var skuMeta={};
  Object.keys(map).forEach(function(cat){
    map[cat].skus.forEach(function(sku){
      skuMeta[sku]={primaryCat:cat,tags:[cat],seasonal:'',itemName:sku};
    });
  });
  return {map:map,primaries:Object.keys(map),tags:Object.keys(map),skuMeta:skuMeta};
}

// SKU × 월 매트릭스
// 반환: { months:['2026-01',...], skus:{ SKU:{itemName, monthly:[qty,...], total} } }
// SKU별 타임컬럼 집계 (SKU뷰 메인 행의 타임컬럼 수량)
function aggregateSkuMonthlyByTimeCols_(rows, timeCols){
  var timeIdx=buildTimeIdx_(timeCols);
  var nCols=(timeCols||[]).length;
  var skuMap={};
  rows.forEach(function(r){
    if(!skuMap[r.sku]) skuMap[r.sku]={itemName:r.itemName,supply:0,
      monthly:new Array(nCols).fill(0)};
    skuMap[r.sku].supply+=r.supply;
    if(timeIdx){
      var mk=r.year+'-'+(r.month<10?'0':'')+r.month;
      var ci=timeIdx[mk];
      if(ci!==undefined) skuMap[r.sku].monthly[ci]+=r.qty;
    }
  });
  // supply 상위 50개 SKU만 전송
  var skuList=Object.keys(skuMap).sort(function(a,b){
    return skuMap[b].supply-skuMap[a].supply;
  }).slice(0,50);
  var result={};
  skuList.forEach(function(sku){
    result[sku]={itemName:skuMap[sku].itemName, supply:skuMap[sku].supply,
                 monthly:skuMap[sku].monthly};
  });
  return {cols:timeCols, skus:result};
}

// 반품 인사이트: SKU × 거래처 반품 집계
// 반환: 반품수량 내림차순 상위 15건 { sku, itemName, customer, retQty, shipQty, retRate, supply }
function aggregateReturnInsight_(rows){
  var m={};
  rows.forEach(function(r){
    if(r.qty>=0) return; // 반품 행만
    var k=r.sku+'||'+r.customer;
    if(!m[k]) m[k]={sku:r.sku,itemName:r.itemName,customer:r.customer,
                    retQty:0,supply:0};
    m[k].retQty+=r.qty; // 음수 누적
    m[k].supply+=r.supply;
  });
  // 출고수량도 같이 계산
  rows.forEach(function(r){
    if(r.qty<0) return;
    var k=r.sku+'||'+r.customer;
    if(!m[k]) return;
    if(!m[k].shipQty) m[k].shipQty=0;
    m[k].shipQty+=r.qty;
  });
  return Object.keys(m).map(function(k){
    var g=m[k];
    g.retRate=g.shipQty>0?Math.round(Math.abs(g.retQty)/g.shipQty*1000)/10:100;
    return g;
  }).filter(function(g){ return g.retQty<0; }) // 반품 있는 것만
   .sort(function(a,b){ return b.retRate - a.retRate; }) // 반품율 내림차순
   .slice(0,5);
}
// ── loadCategoryMap_ 멀티브랜드 확장 ──────────────────────
// multi 모드에서 CATEGORY_MAP 탭 없으면 계층그룹 자동 파싱
function loadCategoryMap_(ss){
  var sh=ss.getSheetByName(CONFIG.CATEGORY_SHEET);
  if(!sh){
    return CONFIG.MODE==='multi'
      ? buildMultiCatData_(ss)
      : buildDefaultCatData_();
  }
  var data=sh.getDataRange().getValues();
  var map={},skuMeta={},primSet={},tagSet={};
  for(var i=1;i<data.length;i++){
    var row=data[i];
    var cat=String(row[0]).trim(),sku=String(row[1]).trim();
    var itemName=String(row[2]).trim(),primary=String(row[3]).trim();
    var tagsRaw=String(row[4]).trim(),seasonal=String(row[5]).trim();
    if(!cat||!sku||cat==='카테고리명') continue;
    if(!map[cat]) map[cat]={label:cat,skus:[]};
    if(map[cat].skus.indexOf(sku)===-1) map[cat].skus.push(sku);
    var tags=tagsRaw?tagsRaw.split('|').map(function(t){return t.trim();}):[];
    var brand=cat.indexOf('/')!==-1?cat.split('/')[0]:(tags[0]||'');
    var series=row.length>6?String(row[6]).trim():'';
    // series 없으면 tags에서 Family 추출 (하위 호환)
    if(!series&&brand==='KLIM') series=getFamilyTag_(tags);
    skuMeta[sku]={primaryCat:primary,tags:tags,seasonal:seasonal,
                  itemName:itemName,brand:brand,series:series};
    if(primary) primSet[primary]=1;
    tags.forEach(function(t){if(t)tagSet[t]=1;});
  }
  if(!Object.keys(map).length){
    return CONFIG.MODE==='multi'
      ? buildMultiCatData_(ss)
      : buildDefaultCatData_();
  }
  return {map:map,primaries:Object.keys(primSet).sort(),
          tags:Object.keys(tagSet).sort(),skuMeta:skuMeta};
}

// ── 멀티브랜드 카테고리 자동 파싱 ─────────────────────────
// 품목등록 시트 없이 판매현황 데이터에서 계층그룹(E열=GROUP)을 읽어 자동 분류
// KLIM: Category 필드 추출, TOURATECH: 2번째 필드, SCHUBERTH: 2번째, BARKBUSTERS: 한글 첫 단어
function buildMultiCatData_(ss){
  // 가장 최근 salesrecord 탭에서 E열(GROUP)·F열(SKU)·G열(ITEM_NAME)·Q열(코드) 샘플링
  var re=new RegExp('^'+CONFIG.SHEET_PREFIX+'\\d{4}$');
  var sheets=ss.getSheets().filter(function(s){return re.test(s.getName());});
  if(!sheets.length) return buildDefaultMulti_();
  var sh=sheets.sort(function(a,b){return b.getName().localeCompare(a.getName(),'ko');})[0];
  var lastRow=sh.getLastRow();
  var dataRows=Math.min(lastRow-CONFIG.DATA_START_ROW+1,5000);
  if(dataRows<=0) return buildDefaultMulti_();

  var vals=sh.getRange(CONFIG.DATA_START_ROW,1,dataRows,17).getValues();
  var map={},skuMeta={},primSet={},tagSet={};

  vals.forEach(function(r){
    var grp  = String(r[CONFIG.COL.GROUP    -1]).trim();
    var sku  = String(r[CONFIG.COL.SKU      -1]).trim();
    var name = String(r[CONFIG.COL.ITEM_NAME-1]).trim();
    var code = r.length>=17?String(r[16]).trim():'';
    if(!sku||!grp||grp==='품목계층그룹명') return;

    var brand=getBrandFromCode_(code);
    if(brand==='SC') return; // 배송비는 별도

    var parsed=parseMultiGroup_(grp,brand,sku,name);
    var primary=parsed.primary;
    var tags   =parsed.tags;
    var cat    =brand+'/'+primary;

    if(!map[cat]) map[cat]={label:cat,skus:[]};
    if(map[cat].skus.indexOf(sku)===-1) map[cat].skus.push(sku);

    if(!skuMeta[sku]){
      skuMeta[sku]={primaryCat:primary,tags:tags,seasonal:'',
                    itemName:name,brand:brand};
      if(primary) primSet[primary]=1;
      tags.forEach(function(t){if(t)tagSet[t]=1;});
    }
  });

  if(!Object.keys(map).length) return buildDefaultMulti_();
  return {map:map,primaries:Object.keys(primSet).sort(),
          tags:Object.keys(tagSet).sort(),skuMeta:skuMeta};
}

// 계층그룹 파싱 — 브랜드별 규칙
function parseMultiGroup_(grp,brand,sku,name){
  var parts=grp.split('ㆍ').map(function(p){return p.trim();});
  var primary='기타';
  var tags=[];

  if(brand==='KLIM'){
    // KlimㆍCategoryㆍJacketㆍStyleㆍADVㆍFamilyㆍBajaS4ㆍSeasonalㆍSummer...
    // 인덱스:  0       1         2       3       4      5       6       7          8
    var catIdx=parts.indexOf('Category');
    if(catIdx!==-1 && parts[catIdx+1]) primary=parts[catIdx+1]; // Jacket/Pants/Glove...
    // Style 태그
    var styleIdx=parts.indexOf('Style');
    if(styleIdx!==-1 && parts[styleIdx+1] && parts[styleIdx+1]!=='n/a') tags.push(parts[styleIdx+1]);
    // Family(시리즈) 태그
    var famIdx=parts.indexOf('Family');
    if(famIdx!==-1 && parts[famIdx+1] && parts[famIdx+1]!=='n/a') tags.push(parts[famIdx+1]);
    // Seasonal 태그
    var seaIdx=parts.indexOf('Seasonal');
    if(seaIdx!==-1 && parts[seaIdx+1] && parts[seaIdx+1]!=='n/a') tags.push(parts[seaIdx+1]);
    // Specially Specified 태그 (Goretex, Mesh 등)
    var ssIdx=parts.indexOf('Specially Specified');
    if(ssIdx!==-1 && parts[ssIdx+1] && parts[ssIdx+1]!=='n/a') tags.push(parts[ssIdx+1]);

  } else if(brand==='TOURATECH'){
    // TouratechㆍRiding GearㆍHelmetㆍAventuro Carbon2
    if(parts.length>=3) primary=parts[1]||'기타';
    if(parts.length>=4) tags.push(parts[2]);
    if(parts.length>=5) tags.push(parts[3]);

  } else if(brand==='SCHUBERTH'){
    // SchuberthㆍHelmetㆍC5 SeriesㆍC5
    if(parts.length>=2) primary=parts[1]||'기타';
    if(parts.length>=3 && parts[2]) tags.push(parts[2]);
    if(parts.length>=4 && parts[3]) tags.push(parts[3]);

  } else if(brand==='BARKBUSTERS'){
    // 핸드가드ㆍBarkbusters / 프레임킷ㆍBarkbusters / 스페어파츠ㆍBarkbusters
    if(parts[0] && parts[0]!=='Barkbusters') primary=parts[0];
    else primary='기타';

  } else {
    primary=parts[0]||'기타';
  }

  // 중복 제거
  var seen={}; tags=tags.filter(function(t){ if(!t||seen[t]) return false; seen[t]=1; return true; });
  return {primary:primary, tags:tags};
}

function buildDefaultMulti_(){
  return {map:{},primaries:[],tags:{},skuMeta:{}};
}
