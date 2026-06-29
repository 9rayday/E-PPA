/* E-PPA app.js — 전력 데이터 파싱 + NASA API + 분석 네비게이션 */

'use strict';

let parsedData = null; // { monthly, type }

/* ── 초기화 ── */
document.addEventListener('DOMContentLoaded', function () {
  setupUpload();
  document.getElementById('btn-location').addEventListener('click', useMyLocation);
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

/* ── KEPCO AMI 파싱 (15분 간격 96열 + 일합계) ── */
function parseAMI(rows) {
  var monthly = {};

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    if (!row[0]) continue;

    var dateStr   = String(row[0]).trim();
    var dateMatch = dateStr.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
    if (!dateMatch) continue;

    var year  = parseInt(dateMatch[1]);
    var month = parseInt(dateMatch[2]);
    if (year < 2010 || year > 2035 || month < 1 || month > 12) continue;

    var nums = [];
    for (var c = 1; c < row.length; c++) {
      var v = parseFloat(row[c]);
      if (!isNaN(v) && v >= 0) nums.push(v);
    }
    if (nums.length === 0) continue;

    /* 일합계 결정: 마지막 값이 앞 96개의 합에 근접하면 합계열로 간주 */
    var sum96   = nums.slice(0, 96).reduce(function (a, b) { return a + b; }, 0);
    var lastVal = nums[nums.length - 1];
    var dayKwh;

    if (nums.length > 96 && sum96 > 0 && Math.abs(lastVal - sum96) / sum96 < 0.05) {
      dayKwh = lastVal;
    } else {
      dayKwh = nums.reduce(function (a, b) { return a + b; }, 0);
    }

    /* Wh → kWh 변환: 일 사용량이 50,000 초과 시 단위 Wh 가정 */
    if (dayKwh > 50000) dayKwh /= 1000;

    var key = year + '-' + month;
    if (!monthly[key]) monthly[key] = { year: year, month: month, kwh: 0, amount: 0 };
    monthly[key].kwh += dayKwh;
  }

  return Object.values(monthly)
    .sort(function (a, b) { return a.year !== b.year ? a.year - b.year : a.month - b.month; })
    .map(function (m) { return { year: m.year, month: m.month, kwh: Math.round(m.kwh), amount: Math.round(m.amount) }; });
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

  /* 평균 단가 자동 산출 */
  var totalAmount = monthly.reduce(function (s, m) { return s + m.amount; }, 0);
  if (totalKwh > 0 && totalAmount > 0) {
    var avgUnit = Math.round(totalAmount / totalKwh);
    document.getElementById('inp-tariff').value = avgUnit;
    document.getElementById('tariff-auto').textContent = '자동';
  }

  /* 설정 패널 표시 */
  var cfg = document.getElementById('cfg');
  cfg.style.display = 'flex';
}

/* ── 현재 위치 사용 ── */
function useMyLocation() {
  var btn = document.getElementById('btn-location');
  btn.textContent = '위치 확인 중...';
  btn.disabled = true;

  if (!navigator.geolocation) {
    alert('이 브라우저는 위치 정보를 지원하지 않습니다.');
    btn.textContent = '현재 위치 사용';
    btn.disabled = false;
    return;
  }

  navigator.geolocation.getCurrentPosition(
    function (pos) {
      document.getElementById('inp-lat').value = pos.coords.latitude.toFixed(4);
      document.getElementById('inp-lng').value = pos.coords.longitude.toFixed(4);
      document.getElementById('loc-auto').style.display = 'inline';
      btn.textContent = '위치 적용됨 ✓';
      btn.disabled = false;
    },
    function (err) {
      alert('위치 정보 오류: ' + err.message);
      btn.textContent = '현재 위치 사용';
      btn.disabled = false;
    }
  );
}

/* ── NASA POWER API — 월별 GHI 평균 ── */
function fetchGHI(lat, lng, startYear, endYear) {
  var start = startYear + '0101';
  var end   = endYear   + '1231';
  var url = 'https://power.larc.nasa.gov/api/temporal/daily/point'
    + '?parameters=ALLSKY_SFC_SW_DWN&community=RE'
    + '&longitude=' + lng + '&latitude=' + lat
    + '&start=' + start + '&end=' + end + '&format=JSON';

  return fetch(url)
    .then(function (r) {
      if (!r.ok) throw new Error('NASA API 오류 ' + r.status);
      return r.json();
    })
    .then(function (data) {
      var daily  = data.properties.parameter.ALLSKY_SFC_SW_DWN;
      var byMonth = {};

      Object.keys(daily).forEach(function (dateStr) {
        var v     = daily[dateStr];
        if (v <= 2) return; /* GHI 유효값 필터 (PV_P.js 동일 로직) */
        var month = parseInt(dateStr.slice(4, 6));
        if (!byMonth[month]) byMonth[month] = { sum: 0, cnt: 0 };
        byMonth[month].sum += v;
        byMonth[month].cnt++;
      });

      var ghi = {};
      for (var m = 1; m <= 12; m++) {
        ghi[m] = byMonth[m] ? byMonth[m].sum / byMonth[m].cnt : 3.5;
      }
      return ghi;
    });
}

/* ── 월별 태양광 발전량 (kWh) ── */
function calcSolar(cap, ghi, month, year) {
  var PR   = 0.82;
  var days = new Date(year, month, 0).getDate();
  return cap * ghi[month] * PR * days;
}

/* ── 로딩 상태 ── */
function setLoading(on) {
  document.getElementById('btn-cta').disabled     = on;
  document.getElementById('btn-label').style.display   = on ? 'none'   : 'inline';
  document.getElementById('btn-spinner').style.display = on ? 'inline' : 'none';
  document.getElementById('btn-arrow').style.display   = on ? 'none'   : 'inline';
}

/* ── 메인 분석 실행 ── */
function runAnalysis() {
  if (!parsedData) {
    alert('전력 데이터 파일을 먼저 업로드하세요.');
    return;
  }

  var lat = parseFloat(document.getElementById('inp-lat').value);
  var lng = parseFloat(document.getElementById('inp-lng').value);
  var cap = parseFloat(document.getElementById('inp-capacity').value);
  var ppa = parseFloat(document.getElementById('inp-ppa').value);

  if (isNaN(lat) || isNaN(lng)) { alert('위치 좌표를 확인하세요.'); return; }
  if (!cap || cap <= 0)         { alert('태양광 설치 용량을 입력하세요.'); return; }
  if (!ppa || ppa <= 0)         { alert('PPA 단가를 입력하세요.'); return; }

  var tariff = parseFloat(document.getElementById('inp-tariff').value) || 0;

  setLoading(true);

  var monthly    = parsedData.monthly;
  var years      = monthly.map(function (m) { return m.year; });
  var startYear  = Math.min.apply(null, years);
  var endYear    = Math.max.apply(null, years);

  fetchGHI(lat, lng, startYear, endYear)
    .then(function (ghi) {
      var enriched = monthly.map(function (m) {
        var gen      = calcSolar(cap, ghi, m.month, m.year);
        var eff      = Math.min(gen, m.kwh);
        var selfRate = m.kwh > 0 ? eff / m.kwh : 0;
        return {
          year: m.year, month: m.month,
          kwh: m.kwh, amount: m.amount,
          gen: Math.round(gen),
          eff: Math.round(eff),
          selfRate: selfRate
        };
      });

      localStorage.setItem('eppa_results', JSON.stringify({
        monthly:     enriched,
        params:      { lat: lat, lng: lng, cap: cap, ppa: ppa, tariff: tariff, pr: 0.82 },
        ghi:         ghi,
        generatedAt: new Date().toISOString()
      }));

      window.location.href = 'summary.html';
    })
    .catch(function (err) {
      alert('NASA API 오류: ' + err.message + '\n\n인터넷 연결을 확인하거나 잠시 후 다시 시도하세요.');
      setLoading(false);
    });
}
