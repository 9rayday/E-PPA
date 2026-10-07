/* E-PPA app.js — 전력 데이터 파싱 + KMA/NASA MCP 보정 GHI + 분석 네비게이션 */

'use strict';

let parsedData = null; // { monthly, type }
var _region = 'sudo'; // 'sudo' | 'nonsudo'
var _distSource = 'actual'; // 'actual'(3개년 실측 분포, 기본값) | 'forecast'(GHI 모델 기반 이론치)

/* NASA POWER API 2023-2025 월별 GHI 3개년 평균 (kWh/m²/day, 월1~12)
   수도권: 서울 37.57°N 126.98°E
   비수도권: 대구 35.87°N 128.60°E */
var GHI_REGION = {
  sudo:    [2.29, 3.22, 4.25, 4.86, 5.32, 5.69, 4.64, 4.93, 4.01, 3.19, 2.56, 2.06],
  nonsudo: [2.82, 3.34, 4.31, 5.05, 5.57, 5.48, 5.00, 5.29, 3.93, 3.35, 3.02, 2.54]
};

function selectRegion(r) {
  _region = r;
  document.getElementById('btn-sudo').classList.toggle('active', r === 'sudo');
  document.getElementById('btn-nonsudo').classList.toggle('active', r === 'nonsudo');
}

function getGHI() {
  var vals = GHI_REGION[_region], ghi = {};
  for (var m = 1; m <= 12; m++) ghi[m] = vals[m - 1];
  return ghi;
}

/* 3개년 실측 기반 지역별 월별 발전량 분포(%, 합계=100)
   수도권: 서울/인천/경기 3개 지역 평균, 비수도권: 나머지 13개 시도 평균
   NASA POWER GHI는 연간 총 발전시간(발전량) 추론에 사용하고,
   월별 배분은 실측 발전량 통계 기반의 이 분포를 사용한다.
   (GHI 모델은 장마철(6~8월) 발전 손실을 과소평가하고 만추~겨울철을 과소평가하는 경향이 있어 보정 효과가 있음) */
var MONTHLY_DIST_REGION = {
  sudo:    [5.70, 6.93, 9.50, 9.80, 10.80, 10.67, 8.63, 9.47, 8.17, 7.67, 6.77, 5.83],
  nonsudo: [5.88, 6.55, 9.14, 9.70, 10.62, 10.37, 9.05, 10.03, 7.97, 7.55, 7.12, 6.04]
};

function selectDistSource(s) {
  _distSource = s;
  document.getElementById('btn-dist-actual').classList.toggle('active', s === 'actual');
  document.getElementById('btn-dist-forecast').classList.toggle('active', s === 'forecast');
}

/* 실측 분포(MONTHLY_DIST_REGION) 대신 "예측값"을 고르면, NASA POWER GHI(일사량)
   모델 자체가 함의하는 월별 비중(GHI[월]×해당월 일수, 연간 합으로 정규화)을 씀.
   실측 분포처럼 연도와 무관한 고정 비율표를 쓰기 위해 평년(365일) 기준으로 계산 —
   윤년 2월 하루 차이는 비중에 0.1%p 미만 영향이라 연도별로 다시 계산할 필요는 없음. */
function getGHIDist() {
  var ghi = getGHI(), days = [31,28,31,30,31,30,31,31,30,31,30,31];
  var raw = {}, total = 0;
  for (var m = 1; m <= 12; m++) { raw[m] = ghi[m] * days[m - 1]; total += raw[m]; }
  var dist = {};
  for (var m2 = 1; m2 <= 12; m2++) dist[m2] = total > 0 ? raw[m2] / total : 1 / 12;
  return dist;
}

function getMonthlyDist() {
  if (_distSource === 'forecast') return getGHIDist();
  var vals = MONTHLY_DIST_REGION[_region], dist = {};
  for (var m = 1; m <= 12; m++) dist[m] = vals[m - 1] / 100;
  return dist;
}

/* ── 초기화 ── */
document.addEventListener('DOMContentLoaded', function () {
  setupUpload();
  document.getElementById('btn-cta').addEventListener('click', runAnalysis);
});

/* ── 파일 업로드 설정 ── */
function setupUpload() {
  var box   = document.getElementById('upload-box');
  var input = document.getElementById('file-input');

  box.addEventListener('click', function (e) {
    if (e.target !== input) input.click();
  });
  input.addEventListener('change', function (e) {
    if (e.target.files[0]) onFile(e.target.files[0]);
  });
  box.addEventListener('dragover', function (e) {
    e.preventDefault(); box.classList.add('drag-over');
  });
  box.addEventListener('dragleave', function () {
    box.classList.remove('drag-over');
  });
  box.addEventListener('drop', function (e) {
    e.preventDefault();
    box.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]);
  });
}

/* ── 파일 읽기 ── */
function onFile(file) {
  var reader = new FileReader();
  reader.onload = function (e) {
    try {
      var wb   = XLSX.read(e.target.result, { type: 'array' });
      var ws   = wb.Sheets[wb.SheetNames[0]];
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      var type    = autoDetect(rows);
      var monthly = type === 'ami' ? parseAMI(rows) : parseMonthly(rows);

      if (!monthly || monthly.length === 0) {
        alert('데이터를 읽을 수 없습니다.\nKEPCO AMI 형식 또는 월별(년도|월|사용량|청구금액) 형식을 확인하세요.');
        return;
      }

      parsedData = { monthly: monthly, type: type };
      onLoaded(file.name, monthly);
    } catch (err) {
      alert('파일 파싱 오류: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ── 형식 자동 감지 ── */
function autoDetect(rows) {
  for (var r = 0; r < Math.min(rows.length, 10); r++) {
    var row = rows[r].filter(function (c) { return c !== '' && c !== null; });
    if (row.length > 20) return 'ami';
  }
  return 'monthly';
}

/* ── 계절·시간대 구분 (한전 전기요금표 2026.6.1 시행 기준) ──
   토요일은 최대부하→중간부하 재분류. 공휴일 재분류(→경부하)는 별도 휴일 목록이
   필요해 미반영 — 평일/토요일 구분까지만 실측 반영, 나머지는 추정치 유지. */
function getSeason(month) {
  if (month >= 6 && month <= 8) return 'summer';
  if ([3, 4, 5, 9, 10].indexOf(month) >= 0) return 'spring_fall';
  return 'winter';
}
function getTouBucket(hour, season) {
  if (hour >= 22 || hour < 8) return 'lo';                 /* 경부하 22~08시, 전 계절 동일 */
  if (season === 'winter') {
    if ((hour >= 8 && hour < 9) || (hour >= 12 && hour < 16) || (hour >= 19 && hour < 22)) return 'mid';
    return 'hi';                                            /* 09~12, 16~19시 */
  }
  if ((hour >= 8 && hour < 15) || (hour >= 21 && hour < 22)) return 'mid';
  return 'hi';                                              /* 15~21시 */
}

/* ── KEPCO AMI 파싱 (15분/30분/60분 등 간격 자동감지 + 일합계) ──
   열의 "개수"가 아니라 헤더의 시간 라벨(00:15~24:00 또는 01:00~24:00 등, 합계)로
   각 열이 몇 번째 구간인지 확정 판별. 예전엔 앞에서부터 96개(15분 간격)를 구간값으로
   고정 가정했는데, 특정 날짜에 구간 하나라도 빈 칸(결측)이면 뒤 구간들이 통째로 밀리면서
   합계열(그날 총사용량, 수만 단위)까지 구간값 후보에 섞여 최댓값으로 잘못 잡히는 문제가
   있었음(요금적용전력이 터무니없이 크게 나오는 원인). 라벨 간 최소 간격으로 구간 길이
   (stepMin)를 자동감지해 15분/60분 등 다른 간격의 파일도 지원. */
function parseAMI(rows) {
  var header   = rows[0] || [];
  var timeCols = [];  /* {col, totalMin} — 헤더에서 인식된 시각 라벨들 */
  var sumColIdx = -1;
  for (var h = 1; h < header.length; h++) {
    var label = String(header[h]).trim();
    if (label.indexOf('합계') >= 0) { sumColIdx = h; continue; }
    var tm = label.match(/^(\d{1,2}):(\d{2})$/);
    if (tm) timeCols.push({ col: h, totalMin: parseInt(tm[1], 10) * 60 + parseInt(tm[2], 10) });
  }

  /* 라벨 간 최소 간격(분)을 구간 길이로 자동감지 (15분 KEPCO 표준 AMI, 60분 시간단위 등) */
  var stepMin = 15, slotCount = 96, colInterval = {}, useHeaderMap = false;
  if (timeCols.length >= 4) {
    var mins = timeCols.map(function (t) { return t.totalMin; }).sort(function (a, b) { return a - b; });
    var minGap = Infinity;
    for (var i = 1; i < mins.length; i++) {
      var g = mins[i] - mins[i - 1];
      if (g > 0 && g < minGap) minGap = g;
    }
    var candStep = isFinite(minGap) && minGap > 0 ? minGap : 15;
    var candSlot = Math.round(1440 / candStep);
    var candMap  = {};
    timeCols.forEach(function (t) {
      var idx = Math.round(t.totalMin / candStep) - 1;
      if (idx >= 0 && idx < candSlot) candMap[t.col] = idx;
    });
    /* 인식된 구간열이 전체 구간의 80% 이상 커버해야 신뢰(그렇지 않으면 옛 방식 폴백) */
    if (Object.keys(candMap).length >= candSlot * 0.8) {
      useHeaderMap = true; stepMin = candStep; slotCount = candSlot; colInterval = candMap;
    }
  }
  var intervalToKw = 60 / stepMin; /* 구간 에너지(kWh) → 순시 kW 환산 배수(15분=4배, 60분=1배 등) */

  /* ── 1차 패스: 날짜별 원본(raw) 값 수집 — 단위(kWh/Wh) 판정은 파일 전체를 보고
     한 번만 내려야 함. 예전엔 "하루 합계 5만 초과=Wh"로 날짜별 판정했는데, 대형
     사업장(예: 5MW급, 평일 하루 7~9만kWh)은 평일만 Wh로 오판되어 1000으로 나뉘고
     주말은 정상 유지되는 식으로 날짜마다 들쭉날쭉하게 깨지는 문제가 있었음. */
  var seenDates  = {};
  var dayRecords = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    if (!row[0]) continue;

    var dateStr   = String(row[0]).trim();
    var dateMatch = dateStr.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
    if (!dateMatch) continue;

    var year  = parseInt(dateMatch[1]);
    var month = parseInt(dateMatch[2]);
    var day   = parseInt(dateMatch[3]);
    if (year < 2010 || year > 2035 || month < 1 || month > 12) continue;

    var dateKey = year + '-' + month + '-' + day;
    if (seenDates[dateKey]) continue;
    seenDates[dateKey] = true;

    var dayKwh, dayMaxInterval, intervalVals; /* intervalVals: 길이 slotCount, 결측 구간은 undefined */

    if (useHeaderMap) {
      intervalVals = new Array(slotCount);
      var sumVal = null;
      for (var c = 1; c < row.length; c++) {
        var v = parseFloat(row[c]);
        if (isNaN(v) || v < 0) continue;
        if (c === sumColIdx) { sumVal = v; continue; }
        if (colInterval.hasOwnProperty(c)) intervalVals[colInterval[c]] = v;
      }
      var validVals = intervalVals.filter(function (x) { return x !== undefined; });
      if (validVals.length === 0 && sumVal === null) continue;
      var intervalSum = validVals.reduce(function (a, b) { return a + b; }, 0);
      dayMaxInterval = validVals.length ? Math.max.apply(null, validVals) : 0;
      /* 합계열이 있고 구간합과 5% 이내로 맞으면 합계열 신뢰, 아니면 구간합 사용(결측 보정) */
      dayKwh = (sumVal !== null && intervalSum > 0 && Math.abs(sumVal - intervalSum) / intervalSum < 0.05)
        ? sumVal : intervalSum;
    } else {
      /* 헤더에서 시간 라벨을 못 읽은 파일 — 기존 개수 기반 추정으로 폴백(15분 96구간 가정) */
      var nums = [];
      for (var c2 = 1; c2 < row.length; c2++) {
        var v2 = parseFloat(row[c2]);
        if (!isNaN(v2) && v2 >= 0) nums.push(v2);
      }
      if (nums.length === 0) continue;
      var sum96   = nums.slice(0, 96).reduce(function (a, b) { return a + b; }, 0);
      var lastVal = nums[nums.length - 1];
      if (nums.length > 96 && sum96 > 0 && Math.abs(lastVal - sum96) / sum96 < 0.05) {
        dayKwh = lastVal;
      } else {
        dayKwh = nums.reduce(function (a, b) { return a + b; }, 0);
      }
      intervalVals = nums.slice(0, 96);
      dayMaxInterval = intervalVals.length ? Math.max.apply(null, intervalVals) : 0;
    }

    dayRecords.push({
      year: year, month: month, day: day,
      dayKwh: dayKwh, dayMaxInterval: dayMaxInterval, intervalVals: intervalVals
    });
  }

  /* 파일 전체에서 "구간(1개 15분/1시간 등) 값"의 최댓값으로 Wh/kWh를 1회만 판정.
     대형 사업장이라도 구간 하나의 값이 10만kWh를 넘는 경우는 사실상 없어(예: 1시간
     구간이면 순간 100MW급) 날짜별로 흔들리지 않는 안전한 임계값으로 쓸 수 있음. */
  var globalMaxInterval = 0;
  dayRecords.forEach(function (d) {
    if (d.dayMaxInterval > globalMaxInterval) globalMaxInterval = d.dayMaxInterval;
  });
  var wattUnit = globalMaxInterval > 100000;

  /* ── 2차 패스: 월별 집계(단위 보정은 파일 전체에 동일하게 적용) ── */
  var monthly = {};
  dayRecords.forEach(function (d) {
    var dayKwh = d.dayKwh, dayMaxInterval = d.dayMaxInterval, intervalVals = d.intervalVals;
    if (wattUnit) { dayKwh /= 1000; dayMaxInterval /= 1000; }

    var key = d.year + '-' + d.month;
    if (!monthly[key]) monthly[key] = { year: d.year, month: d.month, kwh: 0, amount: 0, maxInterval: 0, lo: 0, mid: 0, hi: 0 };
    monthly[key].kwh += dayKwh;
    /* 요금적용전력(순시 최대수요, kW) = 구간 최대 에너지(kWh) × intervalToKw */
    monthly[key].maxInterval = Math.max(monthly[key].maxInterval, dayMaxInterval);

    /* 시간대별(경부하/중간부하/최대부하) 실측 집계 — intervalVals[ti]는 항상 정확히 ti번째 구간 */
    var dow    = new Date(d.year, d.month - 1, d.day).getDay(); /* 0=일 ~ 6=토 */
    var season = getSeason(d.month);
    for (var ti = 0; ti < intervalVals.length; ti++) {
      var val = intervalVals[ti];
      if (val === undefined || val === null) continue;
      var hour   = Math.floor(ti * stepMin / 60);
      var bucket = getTouBucket(hour, season);
      if (dow === 6 && bucket === 'hi') bucket = 'mid'; /* 토요일: 최대부하→중간부하 */
      monthly[key][bucket] += wattUnit ? val / 1000 : val;
    }
  });

  return Object.values(monthly)
    .sort(function (a, b) { return a.year !== b.year ? a.year - b.year : a.month - b.month; })
    .map(function (m) {
      return {
        year: m.year, month: m.month, kwh: Math.round(m.kwh), amount: Math.round(m.amount),
        demandKw: Math.round(m.maxInterval * intervalToKw),
        lo: Math.round(m.lo), mid: Math.round(m.mid), hi: Math.round(m.hi)
      };
    });
}

/* ── 월별 청구 데이터 파싱 (년도|월|사용량|청구금액) ── */
function parseMonthly(rows) {
  var monthly = [];

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var nums = [];

    for (var c = 0; c < row.length; c++) {
      var raw = String(row[c]).replace(/,/g, '').trim();
      var n   = parseFloat(raw);
      if (!isNaN(n)) nums.push(n);
    }

    if (nums.length < 2) continue;

    var year = null, month = null, kwh = null, amount = null;

    for (var i = 0; i < nums.length; i++) {
      var n = nums[i];
      if (n >= 2010 && n <= 2035 && year   === null) { year   = Math.round(n); continue; }
      if (n >= 1    && n <= 12   && month  === null) { month  = Math.round(n); continue; }
      if (n >= 10   && n < 1e7   && kwh    === null) { kwh    = Math.round(n); continue; }
      if (n >= 100  && n < 1e9   && amount === null) { amount = Math.round(n); continue; }
    }

    if (year && month && kwh && kwh > 10) {
      monthly.push({ year: year, month: month, kwh: kwh, amount: amount || 0 });
    }
  }

  return monthly.sort(function (a, b) { return a.year !== b.year ? a.year - b.year : a.month - b.month; });
}

/* ── 파일 로드 완료 처리 ── */
function onLoaded(filename, monthly) {
  var box = document.getElementById('upload-box');
  box.classList.add('has-file');
  document.getElementById('upload-label').textContent = filename;

  var totalKwh = monthly.reduce(function (s, m) { return s + m.kwh; }, 0);
  document.getElementById('upload-hint').textContent =
    monthly.length + '개월 인식 · 총 ' + (totalKwh / 1000).toFixed(0) + 'MWh';

  /* 월 누락 경고 */
  var warnEl = document.getElementById('upload-warn');
  if (!warnEl) {
    warnEl = document.createElement('div');
    warnEl.id = 'upload-warn';
    warnEl.style.cssText = 'font-size:11px;margin-top:8px;line-height:1.7;padding:8px 10px;border-radius:6px;display:none';
    box.appendChild(warnEl);
  }
  if (monthly.length < 12) {
    var presentMonths = monthly.map(function(m){ return m.month; });
    var missingNums = [];
    for (var i = 1; i <= 12; i++) {
      if (presentMonths.indexOf(i) < 0) missingNums.push(i + '월');
    }
    var missingStr = missingNums.length ? missingNums.join(', ') + ' 누락' : (12 - monthly.length) + '개월 누락';
    warnEl.style.cssText += ';background:rgba(218,119,86,.08);border:1px solid rgba(218,119,86,.25);color:#da7756;display:block';
    warnEl.textContent = '⚠ ' + missingStr + ' — 데이터가 부족해 연간 분석값이 낮게 나올 수 있습니다';
  } else {
    warnEl.style.display = 'none';
    warnEl.textContent = '';
  }

  /* 설정 패널 표시 */
  var cfg = document.getElementById('cfg');
  cfg.style.display = 'flex';
}

/* ── 연간 태양광 발전량 (kWh) — NASA POWER GHI 기반 ── */
function calcAnnualSolar(cap, ghi, year) {
  var PR = 0.82, total = 0;
  for (var m = 1; m <= 12; m++) {
    var days = new Date(year, m, 0).getDate();
    total += ghi[m] * days;
  }
  return cap * PR * total;
}

/* ── 월별 태양광 발전량 (kWh) — 연간 총량 × 실측 월별 분포 ── */
function calcSolar(cap, ghi, dist, month, year) {
  return calcAnnualSolar(cap, ghi, year) * dist[month];
}

/* ── 메인 분석 실행 ── */
function runAnalysis() {
  if (!parsedData) {
    alert('전력 데이터 파일을 먼저 업로드하세요.');
    return;
  }

  var cap = parseFloat(document.getElementById('inp-capacity').value);
  var ppa = parseFloat(document.getElementById('inp-ppa').value);

  if (!cap || cap <= 0) { alert('태양광 설치 용량을 입력하세요.'); return; }
  if (!ppa || ppa <= 0) { alert('PPA 단가를 입력하세요.'); return; }

  var tariffPlan = document.getElementById('inp-tariff-plan').value || '';
  var ghi     = getGHI();
  var dist    = getMonthlyDist();
  var monthly = parsedData.monthly;

  var enriched = monthly.map(function (m) {
    var gen      = calcSolar(cap, ghi, dist, m.month, m.year);
    var eff      = Math.min(gen, m.kwh);
    var selfRate = m.kwh > 0 ? eff / m.kwh : 0;
    return {
      year: m.year, month: m.month,
      kwh: m.kwh, amount: m.amount,
      gen: Math.round(gen),
      eff: Math.round(eff),
      selfRate: selfRate,
      demandKw: m.demandKw,
      lo: m.lo, mid: m.mid, hi: m.hi
    };
  });

  localStorage.setItem('eppa_results', JSON.stringify({
    monthly:     enriched,
    params:      { region: _region, cap: cap, ppa: ppa, tariffPlan: tariffPlan, pr: 0.82 },
    ghi:         ghi,
    generatedAt: new Date().toISOString()
  }));

  window.location.href = 'summary.html';
}
