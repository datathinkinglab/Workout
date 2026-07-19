/**
 * 오늘도 생존운동 — Google Apps Script 백엔드
 *
 * 이 스크립트는 Google Sheet에 바인딩(확장 프로그램 > Apps Script)해서 사용합니다.
 * 웹앱으로 배포하면 doGet()이 화면을 반환하고,
 * 화면에서 google.script.run 으로 아래 서버 함수를 호출해 시트에 기록합니다.
 */

var SHEETS = {
  SETTINGS: '사용자설정',
  WORKOUT: '운동기록',
  MEAL: '식사기록',
  HEALTH: '건강기록',
  PLAN: '운동계획'
};

var HEADERS = {};
HEADERS[SHEETS.SETTINGS] = ['사용자명', '시작체중', '목표체중', '키', '시작일', '주차'];
HEADERS[SHEETS.WORKOUT] = ['기록ID', '날짜', '운동종류', '목표시간', '실제시간', '목표횟수', '실제횟수', '세트', '강도', '완료여부', '통증부위', '운동후상태', '메모', '기록시각'];
HEADERS[SHEETS.MEAL] = ['기록ID', '날짜', '시간', '식사구분', '음식명', '식사량', '채소', '단백질', '포만감', '야식', '음료', '사진URL', '메모'];
HEADERS[SHEETS.HEALTH] = ['기록ID', '날짜', '체중', '최고혈압', '최저혈압', '맥박', '걸음수', '수면시간', '컨디션', '허리둘레', '메모'];
HEADERS[SHEETS.PLAN] = ['주차', '요일', '운동종류', '목표시간', '목표횟수', '세트', '설명'];

var DAY_KO = ['일', '월', '화', '수', '목', '금', '토'];
var TOTAL_WEEKS = 12;

// ---------------------------------------------------------------- 웹앱 진입점

function doGet() {
  ensureSheets_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('오늘도 생존운동')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------- 공통 유틸

function getSs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('스프레드시트를 찾을 수 없습니다. 시트에 바인딩하거나 스크립트 속성 SPREADSHEET_ID를 설정하세요.');
  return SpreadsheetApp.openById(id);
}

function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Seoul';
}

function fmtDate_(d) {
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function fmtTime_(d) {
  return Utilities.formatDate(d, tz_(), 'HH:mm');
}

function todayStr_() {
  return fmtDate_(new Date());
}

function parseDate_(s) {
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function dayOfWeekKo_(dateStr) {
  var d = parseDate_(dateStr);
  return d ? DAY_KO[d.getDay()] : '';
}

/** 셀 값을 화면에서 쓰기 좋은 문자열/숫자로 정규화한다. */
function normalizeCell_(header, value) {
  if (value === null || value === undefined || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (header === '시간' || header === '기록시각') return fmtTime_(value);
    return fmtDate_(value);
  }
  return value;
}

function getSheet_(name) {
  ensureSheets_();
  return getSs_().getSheetByName(name);
}

/** 시트를 [{헤더: 값}] 배열로 읽는다. */
function readRecords_(name) {
  var sh = getSheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(String);
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    var empty = true;
    for (var j = 0; j < headers.length; j++) {
      var v = normalizeCell_(headers[j], values[i][j]);
      if (v !== '') empty = false;
      row[headers[j]] = v;
    }
    if (!empty) {
      row._row = i + 1;
      out.push(row);
    }
  }
  return out;
}

function newId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 1000);
}

function appendRecord_(sheetName, obj) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_(sheetName);
    var headers = HEADERS[sheetName];
    var row = headers.map(function (h) {
      return obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
    });
    sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- 시트 초기화

function ensureSheets_() {
  var ss = getSs_();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });

  var settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (settings.getLastRow() < 2) {
    settings.appendRow(['사용자', '', '', '', todayStr_(), 1]);
  }

  var plan = ss.getSheetByName(SHEETS.PLAN);
  if (plan.getLastRow() < 2) {
    var rows = buildDefaultPlan_();
    plan.getRange(2, 1, rows.length, HEADERS[SHEETS.PLAN].length).setValues(rows);
  }
}

/**
 * 12주 기본 운동계획.
 * - 걷기: 주차가 올라갈수록 10분 → 30분으로 서서히 증가 (일요일은 휴식)
 * - 근력: 월/수/금, 3주차부터 2세트·종목 추가
 * - 스트레칭: 매일 5분
 */
function buildDefaultPlan_() {
  var walkMinutes = [10, 12, 15, 18, 20, 22, 25, 25, 28, 28, 30, 30];
  var days = ['월', '화', '수', '목', '금', '토', '일'];
  var rows = [];
  for (var w = 1; w <= TOTAL_WEEKS; w++) {
    days.forEach(function (d) {
      if (d !== '일') {
        rows.push([w, d, '평지 걷기', walkMinutes[w - 1], '', '', '대화가 가능한 속도로']);
      }
      rows.push([w, d, '스트레칭', 5, '', '', '천천히 숨을 쉬면서']);
      if (d === '월' || d === '수' || d === '금') {
        var sets = w < 3 ? 1 : 2;
        var chairReps = Math.min(15, 8 + Math.floor((w - 1) / 2));
        rows.push([w, d, '의자 앉았다 일어나기', '', chairReps, sets, '천천히, 무릎이 아프면 중단']);
        rows.push([w, d, '뒤꿈치 들기', '', 10, sets, '벽이나 의자를 잡고']);
        if (w >= 3) rows.push([w, d, '앉아서 무릎 펴기', '', 10, 1, '좌우 각 10회']);
        if (w >= 5) rows.push([w, d, '벽 밀기', '', 10, 2, '팔꿈치를 천천히 굽혔다 펴기']);
      }
    });
  }
  return rows;
}

// ---------------------------------------------------------------- 설정/주차

function getSettings_() {
  var rows = readRecords_(SHEETS.SETTINGS);
  var s = rows[0] || {};
  return {
    사용자명: s['사용자명'] || '사용자',
    시작체중: s['시작체중'] || '',
    목표체중: s['목표체중'] || '',
    키: s['키'] || '',
    시작일: s['시작일'] || todayStr_()
  };
}

function saveSettings(settings) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_(SHEETS.SETTINGS);
    var start = settings.시작일 && parseDate_(settings.시작일) ? settings.시작일 : todayStr_();
    sh.getRange(2, 1, 1, 6).setValues([[
      settings.사용자명 || '사용자',
      settings.시작체중 || '',
      settings.목표체중 || '',
      settings.키 || '',
      start,
      ''
    ]]);
  } finally {
    lock.releaseLock();
  }
  return getAppData(todayStr_());
}

function getWeekInfo_(settings, dateStr) {
  var start = parseDate_(settings.시작일) || parseDate_(todayStr_());
  var cur = parseDate_(dateStr);
  var diffDays = Math.floor((cur - start) / 86400000);
  if (diffDays < 0) diffDays = 0;
  var weekNum = Math.floor(diffDays / 7) + 1;
  return {
    dayNum: diffDays + 1,
    weekNum: Math.min(weekNum, TOTAL_WEEKS),
    totalWeeks: TOTAL_WEEKS
  };
}

function getPlanFor_(weekNum, dayKo, planRows) {
  var rows = planRows || readRecords_(SHEETS.PLAN);
  return rows.filter(function (r) {
    return Number(r['주차']) === weekNum && String(r['요일']) === dayKo;
  });
}

// ---------------------------------------------------------------- 조회

/** 화면 초기화/새로고침용 데이터 한 번에 반환 */
function getAppData(dateStr) {
  ensureSheets_();
  var date = dateStr && parseDate_(dateStr) ? dateStr : todayStr_();
  var settings = getSettings_();
  var week = getWeekInfo_(settings, date);
  var dayKo = dayOfWeekKo_(date);

  var allWorkouts = readRecords_(SHEETS.WORKOUT);
  var allMeals = readRecords_(SHEETS.MEAL);
  var allHealth = readRecords_(SHEETS.HEALTH);

  var todayWorkouts = allWorkouts.filter(function (r) { return r['날짜'] === date; });
  var todayMeals = allMeals.filter(function (r) { return r['날짜'] === date; });

  var planRows = readRecords_(SHEETS.PLAN);
  var plan = getPlanFor_(week.weekNum, dayKo, planRows).map(function (p) {
    var done = todayWorkouts.some(function (w) {
      return w['운동종류'] === p['운동종류'] && w['완료여부'] && w['완료여부'] !== '미실시';
    });
    return {
      운동종류: p['운동종류'],
      목표시간: p['목표시간'],
      목표횟수: p['목표횟수'],
      세트: p['세트'],
      설명: p['설명'],
      완료: done
    };
  });

  var mealSlots = ['아침', '점심', '저녁'];
  var mealDone = {};
  mealSlots.forEach(function (s) {
    mealDone[s] = todayMeals.some(function (m) { return m['식사구분'] === s; });
  });

  var totalItems = plan.length + mealSlots.length;
  var doneItems = plan.filter(function (p) { return p.완료; }).length +
    mealSlots.filter(function (s) { return mealDone[s]; }).length;

  return {
    today: date,
    dayKo: dayKo,
    settings: settings,
    week: week,
    plan: plan,
    todayWorkouts: todayWorkouts,
    todayMeals: todayMeals,
    mealDone: mealDone,
    achievement: totalItems ? Math.round((doneItems / totalItems) * 100) : 0,
    doneItems: doneItems,
    totalItems: totalItems,
    streak: computeStreak_(allWorkouts, allMeals, allHealth, date),
    weekWorkoutDays: computeWeekWorkoutDays_(allWorkouts, date),
    recentHealth: allHealth.slice(-30).reverse(),
    planExerciseNames: uniqueExerciseNames_(planRows),
    advice: computeAdvice_(settings, date, allWorkouts, allMeals, planRows)
  };
}

function uniqueExerciseNames_(planRows) {
  var names = {};
  (planRows || readRecords_(SHEETS.PLAN)).forEach(function (r) {
    if (r['운동종류']) names[r['운동종류']] = true;
  });
  return Object.keys(names);
}

// ---------------------------------------------------------------- 코칭/자동 조정

/**
 * 최근 7일의 통증·실천율을 보고 안내 메시지를 만든다.
 * 진단이 아니라 격려와 "줄여보기" 제안만 한다. 운동을 강제로 금지하지 않는다.
 */
function computeAdvice_(settings, dateStr, allWorkouts, allMeals, planRows) {
  var end = parseDate_(dateStr);
  var start = new Date(end.getTime() - 6 * 86400000);
  var startDay = parseDate_(settings.시작일);

  var painCount = 0;
  allWorkouts.forEach(function (r) {
    var d = parseDate_(r['날짜']);
    if (!d || d < start || d > end) return;
    if (r['통증부위'] && r['통증부위'] !== '없음') painCount++;
  });

  var sum = 0, n = 0;
  for (var d = new Date(start.getTime()); d <= end; d = new Date(d.getTime() + 86400000)) {
    if (startDay && d < startDay) continue;
    sum += computeDayAchievement_(settings, fmtDate_(d), allWorkouts, allMeals, planRows);
    n++;
  }
  var rate = n ? Math.round(sum / n) : 0;

  var week = getWeekInfo_(settings, dateStr);
  var adjusted = nextWeekAdjusted_(week.weekNum, planRows);

  if (painCount >= 2) {
    return {
      type: 'easier', canAdjust: !adjusted, adjusted: adjusted,
      message: '최근 7일 동안 통증 기록이 ' + painCount + '회 있었어요. 다음 주 운동량을 조금 줄여보는 건 어떨까요? 통증이 계속되면 의료기관 상담을 고려하세요.'
    };
  }
  if (n >= 3 && rate < 40) {
    return {
      type: 'easier', canAdjust: !adjusted, adjusted: adjusted,
      message: '최근 실천율이 ' + rate + '%예요. 괜찮아요 — 목표를 잠깐 낮추면 다시 이어가기가 훨씬 쉬워져요.'
    };
  }
  if (n >= 3 && rate >= 80 && painCount === 0) {
    return {
      type: 'praise', canAdjust: false, adjusted: false,
      message: '최근 실천율 ' + rate + '% · 통증 없이 잘 진행 중이에요. 지금 페이스 그대로! 💪'
    };
  }
  return null;
}

function nextWeekAdjusted_(weekNum, planRows) {
  var target = Math.min(weekNum + 1, TOTAL_WEEKS);
  return (planRows || readRecords_(SHEETS.PLAN)).some(function (r) {
    return Number(r['주차']) === target && String(r['설명']).indexOf('자동 조정') !== -1;
  });
}

/**
 * 다음 주 운동계획의 목표시간·목표횟수를 20% 줄인다 (최소 5).
 * 이미 조정된 주차는 다시 줄이지 않는다. 늘리는 방향의 자동 조정은 하지 않는다.
 */
function applyEasierNextWeek() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var settings = getSettings_();
    var week = getWeekInfo_(settings, todayStr_());
    var target = Math.min(week.weekNum + 1, TOTAL_WEEKS);
    if (!nextWeekAdjusted_(week.weekNum, null)) {
      var sh = getSheet_(SHEETS.PLAN);
      var values = sh.getDataRange().getValues();
      for (var i = 1; i < values.length; i++) {
        if (Number(values[i][0]) !== target) continue;
        var t = parseFloat(values[i][3]);
        if (!isNaN(t) && t > 0) values[i][3] = Math.max(5, Math.round(t * 0.8));
        var reps = parseFloat(values[i][4]);
        if (!isNaN(reps) && reps > 0) values[i][4] = Math.max(5, Math.round(reps * 0.8));
        values[i][6] = (values[i][6] ? values[i][6] + ' ' : '') + '(자동 조정됨)';
      }
      sh.getRange(1, 1, values.length, values[0].length).setValues(values);
    }
  } finally {
    lock.releaseLock();
  }
  return getAppData(todayStr_());
}

/** 기록이 하나라도 있는 날을 실천일로 보고 연속 일수를 계산 */
function computeStreak_(workouts, meals, health, todayStr) {
  var dates = {};
  [workouts, meals, health].forEach(function (list) {
    list.forEach(function (r) {
      if (r['날짜']) dates[r['날짜']] = true;
    });
  });
  var d = parseDate_(todayStr);
  if (!dates[fmtDate_(d)]) d = new Date(d.getTime() - 86400000);
  var streak = 0;
  while (dates[fmtDate_(d)]) {
    streak++;
    d = new Date(d.getTime() - 86400000);
  }
  return streak;
}

function mondayOf_(dateStr) {
  var d = parseDate_(dateStr);
  var shift = (d.getDay() + 6) % 7;
  return new Date(d.getTime() - shift * 86400000);
}

function computeWeekWorkoutDays_(workouts, dateStr) {
  var mon = mondayOf_(dateStr);
  var end = parseDate_(dateStr);
  var days = {};
  workouts.forEach(function (r) {
    if (!r['날짜'] || (r['완료여부'] && r['완료여부'] === '미실시')) return;
    var d = parseDate_(r['날짜']);
    if (d && d >= mon && d <= end) days[r['날짜']] = true;
  });
  return Object.keys(days).length;
}

// ---------------------------------------------------------------- 저장

function saveWorkout(rec) {
  var date = rec.날짜 && parseDate_(rec.날짜) ? rec.날짜 : todayStr_();
  appendRecord_(SHEETS.WORKOUT, {
    기록ID: newId_('W'),
    날짜: date,
    운동종류: rec.운동종류 || '',
    목표시간: rec.목표시간 || '',
    실제시간: rec.실제시간 || '',
    목표횟수: rec.목표횟수 || '',
    실제횟수: rec.실제횟수 || '',
    세트: rec.세트 || '',
    강도: rec.강도 || '',
    완료여부: rec.완료여부 || '완료',
    통증부위: rec.통증부위 || '없음',
    운동후상태: rec.운동후상태 || '',
    메모: rec.메모 || '',
    기록시각: fmtTime_(new Date())
  });
  return getAppData(date);
}

/** 식사 사진 저장용 Drive 폴더 (없으면 생성, ID는 스크립트 속성에 캐시) */
function getPhotoFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 폴더 삭제됨 → 재생성 */ }
  }
  var name = '오늘도 생존운동 사진';
  var it = DriveApp.getFoldersByName(name);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function saveMeal(rec) {
  var date = rec.날짜 && parseDate_(rec.날짜) ? rec.날짜 : todayStr_();

  var photoUrl = rec.사진URL || '';
  if (rec.사진Base64) {
    try {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(rec.사진Base64),
        rec.사진Mime || 'image/jpeg',
        '식사_' + date + '_' + (rec.식사구분 || '') + '_' + Utilities.formatDate(new Date(), tz_(), 'HHmmss') + '.jpg'
      );
      photoUrl = getPhotoFolder_().createFile(blob).getUrl();
    } catch (e) {
      photoUrl = ''; // 사진 저장 실패해도 식사 기록은 계속 저장
    }
  }

  appendRecord_(SHEETS.MEAL, {
    기록ID: newId_('M'),
    날짜: date,
    시간: rec.시간 || fmtTime_(new Date()),
    식사구분: rec.식사구분 || '',
    음식명: rec.음식명 || '',
    식사량: rec.식사량 || '보통',
    채소: rec.채소 || '없음',
    단백질: rec.단백질 || '없음',
    포만감: rec.포만감 || '적당',
    야식: rec.야식 || '아니오',
    음료: rec.음료 || '아니오',
    사진URL: photoUrl,
    메모: rec.메모 || ''
  });
  return getAppData(date);
}

/** 지난 식사 불러오기: 최근 기록에서 (식사구분+음식명) 중복 없이 최대 12개 */
function getRecentMealPresets() {
  var meals = readRecords_(SHEETS.MEAL);
  var seen = {};
  var out = [];
  for (var i = meals.length - 1; i >= 0 && out.length < 12; i--) {
    var m = meals[i];
    if (!m['음식명']) continue;
    var key = m['식사구분'] + '|' + m['음식명'];
    if (seen[key]) continue;
    seen[key] = true;
    out.push({
      식사구분: m['식사구분'],
      음식명: m['음식명'],
      식사량: m['식사량'],
      채소: m['채소'],
      단백질: m['단백질'],
      포만감: m['포만감'],
      야식: m['야식'],
      음료: m['음료'],
      날짜: m['날짜']
    });
  }
  return out;
}

function saveHealth(rec) {
  var date = rec.날짜 && parseDate_(rec.날짜) ? rec.날짜 : todayStr_();
  appendRecord_(SHEETS.HEALTH, {
    기록ID: newId_('H'),
    날짜: date,
    체중: rec.체중 || '',
    최고혈압: rec.최고혈압 || '',
    최저혈압: rec.최저혈압 || '',
    맥박: rec.맥박 || '',
    걸음수: rec.걸음수 || '',
    수면시간: rec.수면시간 || '',
    컨디션: rec.컨디션 || '',
    허리둘레: rec.허리둘레 || '',
    메모: rec.메모 || ''
  });
  return getAppData(date);
}

/**
 * 홈 화면 체크리스트 빠른 체크/해제.
 * 체크: '빠른 체크' 메모의 완료 기록을 추가 (실제시간·횟수는 목표값으로 기록)
 * 해제: 같은 날짜·종류의 '빠른 체크' 기록만 삭제 (직접 입력한 기록은 보호)
 */
function quickToggleExercise(dateStr, exerciseName, done) {
  var date = dateStr && parseDate_(dateStr) ? dateStr : todayStr_();
  if (done) {
    var plan = getPlanFor_(getWeekInfo_(getSettings_(), date).weekNum, dayOfWeekKo_(date));
    var item = null;
    for (var i = 0; i < plan.length; i++) {
      if (plan[i]['운동종류'] === exerciseName) { item = plan[i]; break; }
    }
    appendRecord_(SHEETS.WORKOUT, {
      기록ID: newId_('W'),
      날짜: date,
      운동종류: exerciseName,
      목표시간: item ? item['목표시간'] : '',
      실제시간: item ? item['목표시간'] : '',
      목표횟수: item ? item['목표횟수'] : '',
      실제횟수: item ? item['목표횟수'] : '',
      세트: item ? item['세트'] : '',
      강도: '',
      완료여부: '완료',
      통증부위: '없음',
      운동후상태: '',
      메모: '빠른 체크',
      기록시각: fmtTime_(new Date())
    });
  } else {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = getSheet_(SHEETS.WORKOUT);
      var rows = readRecords_(SHEETS.WORKOUT).filter(function (r) {
        return r['날짜'] === date && r['운동종류'] === exerciseName && r['메모'] === '빠른 체크';
      });
      for (var j = rows.length - 1; j >= 0; j--) {
        sh.deleteRow(rows[j]._row);
      }
    } finally {
      lock.releaseLock();
    }
  }
  return getAppData(date);
}

// ---------------------------------------------------------------- 주간 리포트

/** mode: 'week'(기본, 월요일부터) 또는 'month'(1일부터) */
function getReportData(mode) {
  ensureSheets_();
  mode = mode === 'month' ? 'month' : 'week';
  var today = todayStr_();
  var settings = getSettings_();
  var allWorkouts = readRecords_(SHEETS.WORKOUT);
  var allMeals = readRecords_(SHEETS.MEAL);
  var allHealth = readRecords_(SHEETS.HEALTH);
  var planRows = readRecords_(SHEETS.PLAN);

  var mon = mondayOf_(today);
  var todayD = parseDate_(today);
  var periodStart = mode === 'month'
    ? new Date(todayD.getFullYear(), todayD.getMonth(), 1, 12, 0, 0)
    : mon;

  // ---- 최근 12주 축 (월요일 기준)
  var weekStarts = [];
  for (var i = 11; i >= 0; i--) {
    weekStarts.push(new Date(mon.getTime() - i * 7 * 86400000));
  }
  var weekLabels = weekStarts.map(function (d) {
    return (d.getMonth() + 1) + '/' + d.getDate();
  });

  function weekIndexOf(dateStr) {
    var d = parseDate_(dateStr);
    if (!d) return -1;
    for (var k = weekStarts.length - 1; k >= 0; k--) {
      if (d >= weekStarts[k]) {
        return d < new Date(weekStarts[k].getTime() + 7 * 86400000) ? k : -1;
      }
    }
    return -1;
  }

  // ---- 체중 추이 (주별 평균)
  var weightSum = new Array(12).fill(0);
  var weightCnt = new Array(12).fill(0);
  allHealth.forEach(function (r) {
    var w = parseFloat(r['체중']);
    if (isNaN(w)) return;
    var idx = weekIndexOf(r['날짜']);
    if (idx >= 0) { weightSum[idx] += w; weightCnt[idx]++; }
  });
  var weightSeries = weightSum.map(function (s, k) {
    return weightCnt[k] ? Math.round((s / weightCnt[k]) * 10) / 10 : null;
  });

  // ---- 주별 걷기 시간
  var walkSeries = new Array(12).fill(0);
  allWorkouts.forEach(function (r) {
    if (String(r['운동종류']).indexOf('걷') === -1) return;
    if (r['완료여부'] === '미실시') return;
    var minutes = parseFloat(r['실제시간']) || 0;
    var idx = weekIndexOf(r['날짜']);
    if (idx >= 0) walkSeries[idx] += minutes;
  });

  // ---- 기간(주/월) 일별 달성률 → 평균
  var dayAchievements = [];
  for (var d = new Date(periodStart.getTime()); d <= todayD; d = new Date(d.getTime() + 86400000)) {
    dayAchievements.push(computeDayAchievement_(settings, fmtDate_(d), allWorkouts, allMeals, planRows));
  }
  var periodAchievement = dayAchievements.length
    ? Math.round(dayAchievements.reduce(function (a, b) { return a + b; }, 0) / dayAchievements.length)
    : 0;

  // ---- 기간 식사 기록률 (아침/점심/저녁, 경과일 기준)
  var elapsedDays = dayAchievements.length || 1;
  var mealRate = {};
  ['아침', '점심', '저녁'].forEach(function (slot) {
    var days = {};
    allMeals.forEach(function (m) {
      if (m['식사구분'] !== slot) return;
      var md = parseDate_(m['날짜']);
      if (md && md >= periodStart && md <= todayD) days[m['날짜']] = true;
    });
    mealRate[slot] = Math.round((Object.keys(days).length / elapsedDays) * 100);
  });

  // ---- 혈압 추이 (최근 30개)
  var bpRecords = allHealth.filter(function (r) {
    return parseFloat(r['최고혈압']) && parseFloat(r['최저혈압']);
  }).slice(-30);
  var bp = {
    labels: bpRecords.map(function (r) {
      var d2 = parseDate_(r['날짜']);
      return d2 ? (d2.getMonth() + 1) + '/' + d2.getDate() : r['날짜'];
    }),
    sys: bpRecords.map(function (r) { return parseFloat(r['최고혈압']); }),
    dia: bpRecords.map(function (r) { return parseFloat(r['최저혈압']); })
  };

  // ---- 기간 요약 수치
  var weekMeals = allMeals.filter(function (m) {
    var md = parseDate_(m['날짜']);
    return md && md >= periodStart && md <= todayD;
  });
  var strengthCount = 0;
  var periodWalk = 0;
  allWorkouts.forEach(function (r) {
    var d3 = parseDate_(r['날짜']);
    if (!d3 || d3 < periodStart || d3 > todayD) return;
    if (r['완료여부'] === '미실시') return;
    if (String(r['운동종류']).indexOf('걷') !== -1) {
      periodWalk += parseFloat(r['실제시간']) || 0;
      return;
    }
    if (r['운동종류'] === '스트레칭') return;
    strengthCount++;
  });
  var weekBp = allHealth.filter(function (r) {
    var d4 = parseDate_(r['날짜']);
    return d4 && d4 >= periodStart && d4 <= todayD && parseFloat(r['최고혈압']);
  });
  var avgSys = weekBp.length ? Math.round(weekBp.reduce(function (a, r) { return a + parseFloat(r['최고혈압']); }, 0) / weekBp.length) : null;
  var avgDia = weekBp.length ? Math.round(weekBp.reduce(function (a, r) { return a + parseFloat(r['최저혈압'] || 0); }, 0) / weekBp.length) : null;

  var weights = allHealth.map(function (r) { return parseFloat(r['체중']); }).filter(function (w) { return !isNaN(w); });
  var latestWeight = weights.length ? weights[weights.length - 1] : null;
  var startWeight = parseFloat(settings.시작체중);
  var weightChange = (latestWeight !== null && !isNaN(startWeight))
    ? Math.round((latestWeight - startWeight) * 10) / 10
    : null;

  var mealRateAll = Math.round((weekMeals.filter(function (m) {
    return ['아침', '점심', '저녁'].indexOf(m['식사구분']) !== -1;
  }).length / (elapsedDays * 3)) * 100);

  var activeDays = {};
  allWorkouts.forEach(function (r) {
    if (!r['날짜'] || r['완료여부'] === '미실시') return;
    var d5 = parseDate_(r['날짜']);
    if (d5 && d5 >= periodStart && d5 <= todayD) activeDays[r['날짜']] = true;
  });

  return {
    mode: mode,
    weekLabels: weekLabels,
    weightSeries: weightSeries,
    walkSeries: walkSeries,
    achievement: periodAchievement,
    mealRate: mealRate,
    bp: bp,
    summary: {
      운동일수: Object.keys(activeDays).length,
      걷기시간: Math.round(periodWalk),
      근력횟수: strengthCount,
      식사기록률: Math.min(mealRateAll, 100),
      체중변화: weightChange,
      평균혈압: avgSys ? avgSys + '/' + avgDia : null
    },
    streak: computeStreak_(allWorkouts, allMeals, allHealth, today),
    latestWeight: latestWeight,
    settings: settings
  };
}

function computeDayAchievement_(settings, dateStr, allWorkouts, allMeals, planRows) {
  var week = getWeekInfo_(settings, dateStr);
  var plan = getPlanFor_(week.weekNum, dayOfWeekKo_(dateStr), planRows);
  var dayWorkouts = allWorkouts.filter(function (r) { return r['날짜'] === dateStr; });
  var dayMeals = allMeals.filter(function (r) { return r['날짜'] === dateStr; });
  var done = 0;
  plan.forEach(function (p) {
    var ok = dayWorkouts.some(function (w) {
      return w['운동종류'] === p['운동종류'] && w['완료여부'] && w['완료여부'] !== '미실시';
    });
    if (ok) done++;
  });
  ['아침', '점심', '저녁'].forEach(function (slot) {
    if (dayMeals.some(function (m) { return m['식사구분'] === slot; })) done++;
  });
  var total = plan.length + 3;
  return total ? Math.round((done / total) * 100) : 0;
}
