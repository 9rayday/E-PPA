/* E-PPA app.js — 전력 데이터 파싱(AMI 시간별 일 단위 보존) + NASA POWER GHI 기반 발전 분포 + 분석 네비게이션
   v2 변경점
   · AMI를 월 합계가 아니라 "일별 24시간 kWh"로 보존(days) → summary에서 일 단위 시간별 매칭·시간대별 요금 계산
   · 시간대(경/중/최대) 구분·공휴일·요금 계산은 summary.html에서 요금제/공휴일과 함께 계산(여기서는 하지 않음)
   · 롤링 12개월(예: 8/6~다음해 8/5) 파일도 일 단위로 처리 — 부분월/13개 버킷 문제 방지 */

'use strict';

let parsedData = null; // { monthly, days, type }
var _region = 'sudo'; // 'sudo' | 'nonsudo'
var _distSource = 'actual'; // 'actual'(3개년 실측 분포, 기본값) | 'forecast'(GHI 모델 기반 이론치)

/* NASA POWER API 2023-2025 월별 GHI 3개년 평균 (kWh/m²/day, 월1~12)
   수도권: 서울 37.57°N 126.98°E
   비수도권: 대구 35.87°N 128.60°E */
var GHI_REGION = {
  sudo:    [2.29, 3.22, 4.25, 4.86, 5.32, 5.69, 4.64, 4.93, 4.01, 3.19, 2.56, 2.06],
  nonsudo: [2.82, 3.34, 4.31, 5.05, 5.57, 5.48, 5.00, 5.29, 3.93, 3.35, 3.02, 2.54]
};
var PR_DEFAULT = 0.82;
var DAYS_NONLEAP = [31,28,31,30,31,30,31,31,30,31,30,31];

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
  var ghi = getGHI(), days = DAYS_NONLEAP;
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

/* GHI 모델 기준 연평균 발전시간(h/day) = PR × Σ(GHI×일수) / 365
   (summary 초기값으로 사용 — 예전엔 자가소비량으로 역산해 초과발전이 있으면 값이 낮게 잡히던 문제 수정) */
function getGhiHours() {
  var ghi = getGHI(), s = 0;
  for (var m = 1; m <= 12; m++) s += ghi[m] * DAYS_NONLEAP[m - 1];
  return PR_DEFAULT * s / 365;
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
      var parsed  = type === 'ami' ? parseAMI(rows) : { monthly: parseMonthly(rows), days: null };
      var monthly = parsed.monthly;

      if (!monthly || monthly.length === 0) {
        alert('데이터를 읽을 수 없습니다.\nKEPCO AMI 형식 또는 월별(년도|월|사용량|청구금액) 형식을 확인하세요.');
        return;
      }

      parsedData = { monthly: monthly, days: parsed.days, type: type };
      onLoaded(file.name, monthly, parsed.days);
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

/* ── KEPCO AMI 파싱 (15분/30분/60분 등 간격 자동감지 + 일합계) ──
   열의 "개수"가 아니라 헤더의 시간 라벨(00:15~24:00 또는 01:00~24:00 등, 합계)로
   각 열이 몇 번째 구간인지 확정 판별. 결측 구간이 있어도 뒤 구간이 밀리지 않음.
   라벨 간 최소 간격으로 구간 길이(stepMin)를 자동감지해 15분/60분 등 다른 간격도 지원.
   반환: { monthly:[달력월 집계], days:[{y,m,d,w,h:[24시간 kWh]}] } */
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
  var slotsPerHour = Math.max(1, Math.round(60 / stepMin));

  /* ── 1차 패스: 날짜별 원본(raw) 값 수집 — 단위(kWh/Wh) 판정은 파일 전체를 보고 한 번만 ── */
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

  /* 파일 전체에서 "구간(1개 15분/1시간 등) 값"의 최댓값으로 Wh/kWh를 1회만 판정 */
  var globalMaxInterval = 0;
  dayRecords.forEach(function (d) {
    if (d.dayMaxInterval > globalMaxInterval) globalMaxInterval = d.dayMaxInterval;
  });
  var wattUnit = globalMaxInterval > 100000;

  /* 구간값 → 시간별(0~23시) kWh. 결측 구간이 있으면 그 시간의 존재 구간 평균으로 보정,
     시간 전체가 비면 그날 다른 시간 평균으로 채운 뒤 일합계에 맞춰 스케일 */
  function toHourly(intervalVals, dayKwh) {
    var hrs = new Array(24), have = new Array(24), cnt, hh, k, s;
    for (hh = 0; hh < 24; hh++) {
      s = 0; cnt = 0;
      for (k = 0; k < slotsPerHour; k++) {
        var idx = hh * slotsPerHour + k;
        var val = intervalVals[idx];
        if (val !== undefined && val !== null && !isNaN(val)) { s += val; cnt++; }
      }
      if (cnt > 0) { hrs[hh] = s * (slotsPerHour / cnt); have[hh] = true; }
      else { hrs[hh] = 0; have[hh] = false; }
    }
    var okSum = 0, okN = 0;
    for (hh = 0; hh < 24; hh++) if (have[hh]) { okSum += hrs[hh]; okN++; }
    if (okN === 0) return null;
    var fill = okSum / okN;
    for (hh = 0; hh < 24; hh++) if (!have[hh]) hrs[hh] = fill;
    var tot = 0;
    for (hh = 0; hh < 24; hh++) tot += hrs[hh];
    var scale = (tot > 0 && dayKwh > 0) ? dayKwh / tot : 1;
    for (hh = 0; hh < 24; hh++) hrs[hh] = Math.round(hrs[hh] * scale * 100) / 100;
    return hrs;
  }

  /* ── 2차 패스: 월별 집계 + 일별 시간 배열(단위 보정은 파일 전체에 동일하게 적용) ── */
  var monthly = {};
  var days = [];
  dayRecords.forEach(function (d) {
    var dayKwh = d.dayKwh, dayMaxInterval = d.dayMaxInterval, intervalVals = d.intervalVals;
    var unitDiv = wattUnit ? 1000 : 1;
    dayKwh /= unitDiv; dayMaxInterval /= unitDiv;
    var vals = intervalVals.map(function (x) {
      return (x === undefined || x === null) ? x : x / unitDiv;
    });

    var hourly = toHourly(vals, dayKwh);
    if (!hourly) return;

    var key = d.year + '-' + d.month;
    if (!monthly[key]) monthly[key] = {
      year: d.year, month: d.month, kwh: 0, amount: 0, maxInterval: 0, dayCount: 0
    };
    monthly[key].kwh += dayKwh;
    monthly[key].dayCount++;
    /* 요금적용전력(순시 최대수요, kW) = 구간 최대 에너지(kWh) × intervalToKw */
    monthly[key].maxInterval = Math.max(monthly[key].maxInterval, dayMaxInterval);

    var dow = new Date(d.year, d.month - 1, d.day).getDay(); /* 0=일 ~ 6=토 */
    days.push({ y: d.year, m: d.month, d: d.day, w: dow, h: hourly });
  });

  days.sort(function (a, b) {
    return a.y !== b.y ? a.y - b.y : (a.m !== b.m ? a.m - b.m : a.d - b.d);
  });

  var monthlyArr = Object.values(monthly)
    .sort(function (a, b) { return a.year !== b.year ? a.year - b.year : a.month - b.month; })
    .map(function (m) {
      var dim = new Date(m.year, m.month, 0).getDate();
      return {
        year: m.year, month: m.month, kwh: Math.round(m.kwh), amount: Math.round(m.amount),
        demandKw: Math.round(m.maxInterval * intervalToKw),
        dayCount: m.dayCount, dim: dim, partial: m.dayCount < dim
      };
    });

  return { monthly: monthlyArr, days: days };
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
function onLoaded(filename, monthly, days) {
  var box = document.getElementById('upload-box');
  box.classList.add('has-file');
  document.getElementById('upload-label').textContent = filename;

  var totalKwh = monthly.reduce(function (s, m) { return s + m.kwh; }, 0);
  var nDays = days ? days.length : 0;
  document.getElementById('upload-hint').textContent =
    (days ? nDays + '일(' + (nDays / 30.4).toFixed(1) + '개월)' : monthly.length + '개월') +
    ' 인식 · 총 ' + (totalKwh / 1000).toFixed(0) + 'MWh';

  var warnEl = document.getElementById('upload-warn');
  if (!warnEl) {
    warnEl = document.createElement('div');
    warnEl.id = 'upload-warn';
    warnEl.style.cssText = 'font-size:11px;margin-top:8px;line-height:1.7;padding:8px 10px;border-radius:6px;display:none';
    box.appendChild(warnEl);
  }

  var warnStyle = ';background:rgba(218,119,86,.08);border:1px solid rgba(218,119,86,.25);color:#da7756;display:block';
  var infoStyle = ';background:rgba(93,168,122,.08);border:1px solid rgba(93,168,122,.25);color:#5da87a;display:block';

  if (days) {
    /* AMI: 일 단위 커버리지로 판단 (롤링 12개월 파일은 달력월 13개로 보이지만 365일이면 정상) */
    if (nDays < 350) {
      warnEl.style.cssText = 'font-size:11px;margin-top:8px;line-height:1.7;padding:8px 10px;border-radius:6px' + warnStyle;
      warnEl.textContent = '⚠ ' + nDays + '일치 데이터 — 1년(365일)보다 짧아 연간 분석값이 낮게 나올 수 있습니다';
    } else if (monthly.length > 12) {
      warnEl.style.cssText = 'font-size:11px;margin-top:8px;line-height:1.7;padding:8px 10px;border-radius:6px' + infoStyle;
      warnEl.textContent = '✓ ' + nDays + '일 · 달력월 ' + monthly.length + '개(부분월 포함)로 걸쳐 있는 연속 데이터 — 일 단위로 정확히 계산합니다';
    } else {
      warnEl.style.display = 'none';
      warnEl.textContent = '';
    }
  } else if (monthly.length < 12) {
    var presentMonths = monthly.map(function(m){ return m.month; });
    var missingNums = [];
    for (var i = 1; i <= 12; i++) {
      if (presentMonths.indexOf(i) < 0) missingNums.push(i + '월');
    }
    var missingStr = missingNums.length ? missingNums.join(', ') + ' 누락' : (12 - monthly.length) + '개월 누락';
    warnEl.style.cssText = 'font-size:11px;margin-top:8px;line-height:1.7;padding:8px 10px;border-radius:6px' + warnStyle;
    warnEl.textContent = '⚠ ' + missingStr + ' — 데이터가 부족해 연간 분석값이 낮게 나올 수 있습니다';
  } else {
    warnEl.style.display = 'none';
    warnEl.textContent = '';
  }

  /* 설정 패널 표시 */
  var cfg = document.getElementById('cfg');
  cfg.style.display = 'flex';
}

/* ── 메인 분석 실행 ──
   발전량·자가소비·초과발전·요금은 모두 summary.html에서 계산한다(설비용량·발전시간·정산방식·요금제를
   화면에서 바꿀 때마다 다시 계산해야 하므로). 여기서는 입력값과 지역 분포만 넘긴다. */
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
  var distObj = getMonthlyDist();
  var dist    = [];
  for (var m = 1; m <= 12; m++) dist.push(distObj[m]);

  var payload = {
    monthly:     parsedData.monthly,
    days:        parsedData.days || null,
    params:      { region: _region, cap: cap, ppa: ppa, tariffPlan: tariffPlan, pr: PR_DEFAULT,
                   distSource: _distSource, dist: dist, ghiHours: getGhiHours() },
    ghi:         ghi,
    generatedAt: new Date().toISOString()
  };

  try {
    localStorage.setItem('eppa_results', JSON.stringify(payload));
  } catch (e) {
    /* 저장 용량 초과 등 — 일별 데이터를 빼고 월 단위로라도 진행 */
    try {
      payload.days = null;
      localStorage.setItem('eppa_results', JSON.stringify(payload));
      alert('브라우저 저장 용량 문제로 일별 데이터를 저장하지 못해 월 단위로 분석합니다.');
    } catch (e2) {
      alert('분석 데이터를 저장할 수 없습니다: ' + e2.message);
      return;
    }
  }

  window.location.href = 'summary.html';
}
