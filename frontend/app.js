/* AI-дизайнер презентаций — рабочий интерфейс.
   Бриф живет в браузере, исходные DOCX/XLSX/PDF и таблицы сохраняются сервером. */
(function () {
'use strict';

/* ---------- утилиты ---------- */
var API = window.localStorage.getItem('df_api_base') || '';
var paths = {
  arrow: 'M5 12h14m-6-6 6 6-6 6', back: 'M19 12H5m6-6-6 6 6 6',
  paperclip: 'm20 11-8.8 8.8a5 5 0 0 1-7.1-7.1l10-10a3.4 3.4 0 0 1 4.8 4.8l-10 10a1.8 1.8 0 0 1-2.5-2.5L15 7',
  file: 'M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Zm0 0v6h6M8 13h8M8 17h6',
  link: 'M10 13a5 5 0 0 0 7 .1l3-3a5 5 0 0 0-7-7l-2 2m3 6a5 5 0 0 0-7-.1l-3 3a5 5 0 0 0 7 7l2-2',
  edit: 'm16 3 5 5M4 20l4-1 13-13-4-4L4 15Z', check: 'm5 12 4 4L19 6',
  spark: 'm12 3 2.3 6.7L21 12l-6.7 2.3L12 21l-2.3-6.7L3 12l6.7-2.3ZM20 2v4m-2-2h4',
  refresh: 'M20 7v5h-5M4 17v-5h5M6.1 7a7 7 0 0 1 11.6-2L20 8M4 16l2.3 3A7 7 0 0 0 18 17',
  target: 'M12 8a4 4 0 1 0 4 4M21 12a9 9 0 1 1-9-9m0 9 9-9m-5 0h5v5',
  grid: 'M3 3h7v7H3Zm11 0h7v7h-7ZM3 14h7v7H3Zm11 0h7v7h-7Z', download: 'M12 3v12m-5-5 5 5 5-5M4 16v5h16v-5',
  shield: 'm12 3 8 3v6c0 5-8 9-8 9s-8-4-8-9V6Zm-4 9 3 3 5-6', info: 'M12 11v6m0-10h.01M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0',
  chevron: 'm9 5 7 7-7 7', warning: 'm12 3 10 18H2Zm0 6v5m0 3h.01'
};
function icon(n) { return '<svg class="icon" aria-hidden="true" viewBox="0 0 24 24"><path d="' + (paths[n] || paths.file) + '"/></svg>'; }
function raw(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
/* неразрывные пробелы после коротких слов и перед тире */
function t(s) {
  return raw(s).replace(/(^|[\s>])(а|в|во|и|из|к|на|не|но|о|об|по|с|со|у|за|от|до|для|А|В|И|К|На|Не|О|С) /g, '$1$2\u00a0').replace(/ (—)/g, '\u00a0$1');
}
function plural(n, f) { var a = Math.abs(n) % 100, b = a % 10; if (a > 10 && a < 20) return f[2]; if (b > 1 && b < 5) return f[1]; if (b === 1) return f[0]; return f[2]; }
function button(label, action, kind, ico, extra) {
  kind = kind || 'primary'; ico = ico === undefined ? 'arrow' : ico;
  return '<button class="' + kind + '" data-action="' + action + '" ' + (extra || '') + '>' + (ico === 'back' ? icon('back') : '') + t(label) + (ico && ico !== 'back' ? icon(ico) : '') + '</button>';
}
function $(id) { return document.getElementById(id); }
function notify(text, bad) {
  var el = $('notice'); el.textContent = text; el.classList.toggle('toast-err', !!bad); el.classList.add('visible');
  clearTimeout(window.noticeTimer); window.noticeTimer = setTimeout(function () { el.classList.remove('visible'); }, 5200);
}
async function api(path, opts) {
  opts = opts || {};
  var resp;
  try { resp = await fetch(API + path, opts); }
  catch (e) { throw new Error('Нет связи с сервером. Проверьте соединение и повторите попытку.'); }
  if (!resp.ok) {
    var detail = '';
    try { var j = await resp.json(); detail = typeof j.detail === 'string' ? j.detail : JSON.stringify(j.detail || j); } catch (e) { detail = resp.statusText; }
    throw new Error(detail || ('Ошибка сервера ' + resp.status));
  }
  var ct = resp.headers.get('content-type') || '';
  return ct.indexOf('application/json') !== -1 ? resp.json() : resp;
}
function postJson(path, body) { return api(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) }); }
function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

/* ---------- состояние сессии ---------- */
var S = {
  screen: 0, maxScreen: 0,
  tpl: { name: '', id: '', status: 'none', error: '', ds: null, promise: null },
  topic: '', materials: [], matBusy: 0,
  audience: '', goal: '', idea: '', ideaBusy: false, ideaError: '', clarify: null, clarifyAnswer: '', ideaHistory: '',
  min: 8, max: 10, tone: 'Деловой и понятный', must: '',
  job: null, jobId: '', runId: 0, runError: '', runKind: '',
  plan: 'A', busySlides: {}, planBusy: false,
  stale: false, usedBrief: '',
  design: '', designMode: 'columns', cmp: 0, previews: {}, previewErr: {},
  selectedSlide: 0, stamp: Date.now(), refineBusy: '',
  audit: 'choice', auditMap: null, auditBusy: false, format: 'pptx'
};
var lastScreen = null;
var TONES = ['Деловой и понятный', 'Спокойный и доказательный', 'Дружелюбный', 'Энергичный, убеждающий'];
var VL = { variant_a: 'A', variant_b: 'B', variant_c: 'C' };
var VARIANT_FALLBACK = {
  variant_a: ['Лаконичный', 'Заголовки тезисов и короткие подписи, много воздуха'],
  variant_b: ['Сбалансированный', 'Заголовок и короткое пояснение к каждому тезису'],
  variant_c: ['Насыщенный', 'Развернутые пояснения, заметки-выводы, фото, слайды-разделители']
};
var KIND_RU = {
  cover: 'обложка', agenda: 'оглавление', section: 'разделитель', cards: 'карточки', team: 'команда', diagram: 'схема',
  gallery: 'галерея', timeline: 'таймлайн', metrics: 'метрики', table: 'таблица', chart: 'график', statement: 'тезис',
  text: 'текст', image_text: 'фото и текст', closing: 'финал', qa: 'вопросы', cta: 'призыв', problems: 'проблемы',
  columns: 'колонки', title_only: 'заголовок', speaker: 'спикер', list: 'список', quote: 'цитата', comparison: 'сравнение'
};
function kindRu(k) { return KIND_RU[k] || (k || '').replace(/_/g, ' '); }
function deyo(s) { return String(s == null ? '' : s).replace(/ё/g, 'е').replace(/Ё/g, 'Е'); }

/* ---------- материалы ---------- */
function okMaterials() { return S.materials.filter(function (m) { return m.status === 'ok'; }); }
function materialsText() {
  return okMaterials().map(function (m) { return '### ' + m.name + '\n' + m.text; }).join('\n\n');
}
function materialsMeta() {
  return S.materials.map(function (m) { return { name: m.name, kind: m.kind, chars: m.chars || 0, status: m.status, note: m.note || '' }; });
}
function materialIds() { return okMaterials().filter(function (m) { return !!m.materialId; }).map(function (m) { return m.materialId; }); }
function tablePreview(data, label) {
  if (!data) return '';
  return '<div class="source-table-wrap"><table class="source-table"><caption>' + raw(label || 'Таблица') + '</caption><thead><tr>' +
    (data.headers || []).map(function (h) { return '<th scope="col">' + raw(h) + '</th>'; }).join('') + '</tr></thead><tbody>' +
    (data.rows || []).map(function (row) { return '<tr>' + row.map(function (cell) { return '<td>' + raw(cell) + '</td>'; }).join('') + '</tr>'; }).join('') +
    '</tbody></table></div>';
}
async function addFiles(files) {
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    var rec = { id: 'm' + Date.now() + i, name: f.name, kind: 'file', status: 'busy', note: 'Читаем файл…', chars: 0, text: '' };
    S.materials.push(rec); S.matBusy++; refreshMaterials();
    var fd = new FormData(); fd.append('file', f);
    try { applyExtract(rec, await api('/materials/extract', { method: 'POST', body: fd })); }
    catch (e) { rec.status = 'error'; rec.note = e.message; }
    S.matBusy--; if (rec.status === 'ok') markChanged(); refreshMaterials();
  }
}
async function addLink(url) {
  var rec = { id: 'm' + Date.now(), name: url, kind: 'link', status: 'busy', note: 'Читаем страницу…', chars: 0, text: '' };
  S.materials.push(rec); S.matBusy++; refreshMaterials();
  var fd = new FormData(); fd.append('url', url);
  try { applyExtract(rec, await api('/materials/extract', { method: 'POST', body: fd })); }
  catch (e) { rec.status = 'error'; rec.note = e.message; }
  S.matBusy--; if (rec.status === 'ok') markChanged(); refreshMaterials();
}
function applyExtract(rec, r) {
  rec.status = r.status === 'ok' ? 'ok' : 'error'; rec.note = r.note || ''; rec.chars = r.chars || 0; rec.text = r.text || '';
  rec.materialId = r.material_id || ''; rec.tables = r.tables || [];
  if (rec.status === 'ok') rec.note = (r.note ? r.note + ' · ' : '') + rec.chars.toLocaleString('ru-RU') + ' знаков';
}
function fileRows() {
  return S.materials.map(function (m, i) {
    var lead = m.status === 'busy' ? '<span class="spinner"></span>' : icon(m.kind === 'link' ? 'link' : 'file');
    var tables = (m.tables || []).map(function (table) {
      return '<details class="source-table-details"><summary>Таблица ' + table.ordinal + (table.available ? '' : table.analysis_available ? ' · доступна для анализа' : ' · недоступна для слайда') + '</summary><p>' +
        raw(table.summary) + '</p>' + (table.available ? tablePreview(table.data, m.name) : '<p class="table-limitation">' + raw(table.limitation) + '</p>') + '</details>';
    }).join('');
    return '<div class="material-entry"><div class="file-row ' + (m.status === 'error' ? 'bad' : '') + '">' + lead + '<div class="file-details"><b>' + t(m.name) + '</b><small>' + t(m.status === 'error' ? 'Не прочитано: ' + m.note : m.note) + '</small></div><button class="file-remove" data-remove="' + i + '" aria-label="Убрать ' + raw(m.name) + '">×</button></div>' + tables + '</div>';
  }).join('');
}
function refreshMaterials() {
  var box = $('files-box'); if (box) box.innerHTML = fileRows();
  var box2 = $('files-box-panel'); if (box2) box2.innerHTML = fileRows();
  updateSide();
}

/* ---------- изменения брифа после построения плана ---------- */
function briefSignature() {
  return JSON.stringify([S.topic, materialsText(), materialIds(), S.audience, S.goal, S.idea, S.min, S.max, S.tone, S.must, S.tpl.id]);
}
function markChanged() {
  if (S.job && S.usedBrief && briefSignature() !== S.usedBrief) S.stale = true;
  updateSide();
}

/* ---------- каркас ---------- */
var heroHtml = '<header class="hero"><h1 class="hero-brand brand">AI-дизайнер <span class="grad">презентаций</span></h1><p class="hero-sub">' +
  t('Загрузите презентацию-ориентир по стилю и расскажите о задаче — мы подготовим план выступления и соберем три варианта оформления, в которых ваши мысли прозвучат убедительно') + '</p></header>';
function buildShell() {
  // Верхней полосы с логотипом нет: в рабочем режиме логотип стоит в начале блока с контентом, слева.
  var workBrand = grp(S.screen) === 'work' ? '<div class="work-brand"><button class="brand" data-action="home" title="На главную">AI-дизайнер <span class="grad">презентаций</span></button></div>' : '';
  $('app').innerHTML = '<div id="shell" class="shell"><div class="layout"><main id="workspace" class="main">' + workBrand + heroHtml +
    '<div id="form-area" class="form-area"></div></main><aside class="sidebar" id="sidebar" aria-label="Ваша презентация"><div class="side-sticky" id="side-inner"></div></aside></div></div>';
}
function grp(s) { return s === null ? null : s < 4 ? 'brief' : 'work'; }
function render() {
  var prev = lastScreen;
  var shell = $('shell');
  var rebuilt = !shell || grp(prev) !== grp(S.screen);
  if (rebuilt) buildShell();
  shell = $('shell');
  var mode = (S.screen === 0 ? 'is-start' : S.screen < 4 ? 'is-brief' : 'is-work') + (S.screen >= 5 ? ' compact' : '');
  document.body.className = S.screen === 0 ? 'start' : S.screen < 4 ? 'brief' : 'work';
  var form = $('form-area');
  var changed = prev !== null && prev !== S.screen;
  var reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
  var inBrief = prev !== null && prev < 4 && S.screen < 4;
  var apply = function () {
    form.innerHTML = mainContent(); form.classList.remove('swap-out');
    if (changed && !reduced) { form.classList.remove('swap-in'); void form.offsetWidth; form.classList.add('swap-in'); }
    if (window.auraCtl) window.auraCtl(S.screen > 0);
    afterRender();
  };
  if (S.screen > 0) $('side-inner').innerHTML = sidebar();
  if (rebuilt && prev !== null && !reduced && grp(prev) !== grp(S.screen)) { shell.className = 'shell ' + mode; apply(); }
  else if (inBrief && !reduced) {
    clearTimeout(window.swapTimer); form.classList.add('swap-out'); shell.className = 'shell ' + mode;
    window.swapTimer = setTimeout(apply, prev === 0 || S.screen === 0 ? 260 : 150);
  } else { shell.className = 'shell ' + mode; apply(); }
  lastScreen = S.screen;
}
function updateSide() { var el = $('side-inner'); if (el && S.screen > 0) el.innerHTML = sidebar(); }
function remember() {
  ['topic', 'audience', 'goal', 'must'].forEach(function (k) { var e = $(k); if (e && e.value !== undefined) S[k] = e.value; });
  var tn = $('tone'); if (tn) S.tone = tn.value;
  var a = $('volume-min'), b = $('volume-max');
  if (a) S.min = +a.value; if (b) S.max = +b.value;
}
function go(n) {
  remember();
  S.screen = Math.min(9, Math.max(0, +n));
  S.maxScreen = Math.max(S.maxScreen, S.screen);
  render(); window.scrollTo(0, 0);
}

/* ---------- правая панель ---------- */
function sideBlock(title, inner, status, current) {
  return '<section class="side-block ' + (current ? 'current' : '') + '"><h2 class="side-heading">' + title + (status ? '<span class="side-state">' + status + '</span>' : '') + '</h2>' + inner + '</section>';
}
function paletteMini() {
  var cols = (S.tpl.ds && S.tpl.ds.colors) || [];
  var pick = cols.filter(function (c) { return c.group === 'accent'; }).slice(0, 3);
  var bg = cols.filter(function (c) { return c.group !== 'accent'; });
  pick = pick.concat(bg.slice(0, 4 - pick.length));
  return pick.length ? '<div class="palette-mini">' + pick.map(function (c) { return '<span class="swatch" style="background:' + raw(c.hex) + '" title="' + raw(c.label) + '"></span>'; }).join('') + '</div>' : '';
}
function planSlides() {
  var v = S.job && S.job.plan_variants && S.job.plan_variants[S.plan];
  return v ? v.plan.slides : [];
}
function sidebar() {
  var out = '';
  var tp = S.tpl;
  if (tp.status === 'analyzing' || tp.status === 'uploading') {
    out += sideBlock('Дизайн-система', '<p class="side-text" style="margin-bottom:12px">' + t(tp.name) + '</p><div class="statusline"><span class="spinner"></span>Анализируем шаблон</div><p class="side-text" style="margin-top:12px">Можно продолжать — анализ идет в фоне.</p>');
  } else if (tp.status === 'error') {
    out += sideBlock('Дизайн-система', '<p class="side-text">' + t(tp.name) + '</p><p class="err">' + t(tp.error) + '</p><button class="side-link" data-action="choose-template">Загрузить другой файл ' + icon('chevron') + '</button>');
  } else {
    out += sideBlock('Дизайн-система', '<p class="side-text">' + t(tp.name) + '</p>' + paletteMini() + '<button class="side-link" data-panel="design-system">Открыть дизайн-систему ' + icon('chevron') + '</button>', 'Готово');
  }
  if (S.screen >= 2) {
    var okN = okMaterials().length;
    var mats = S.screen >= 5 ? '<p class="side-text">Описание' + (S.materials.length ? ', ' + S.materials.length + ' ' + plural(S.materials.length, ['материал', 'материала', 'материалов']) : '') + '</p>' :
      '<p class="side-text">' + t(S.topic.slice(0, 110)) + (S.topic.length > 110 ? '…' : '') + '</p>' +
      S.materials.slice(0, 3).map(function (m) { return '<div class="side-material">' + icon(m.kind === 'link' ? 'link' : 'file') + '<span>' + t(m.name) + (m.status === 'error' ? ' · не прочитано' : '') + '</span></div>'; }).join('');
    out += sideBlock('Материалы', mats + '<button class="side-link" data-panel="materials">Посмотреть и дополнить ' + icon('chevron') + '</button>', (okN + 1) + ' ' + plural(okN + 1, ['источник', 'источника', 'источников']));
  }
  if (S.screen >= 3 && (S.ideaBusy || S.idea)) {
    var m = S.ideaBusy ? '<div class="statusline"><span class="spinner"></span>Определяем главную идею</div>' :
      '<p class="small-label" style="margin-bottom:8px">Главная идея</p><button class="meaning-side" data-panel="meaning">' + t(S.idea) + '<span class="edit-hint">' + t('Нажмите, чтобы скорректировать') + ' ↗</span></button>' +
      (S.screen >= 5 ? '' : '<div class="side-brief-meta"><b>Для кого</b><br>' + t(S.audience) + '<br><br><b>Решение аудитории</b><br>' + t(S.goal) + '</div>') +
      '<button class="side-link" data-panel="meaning">' + (S.screen >= 5 ? 'Изменить идею, аудиторию и цель' : 'Изменить аудиторию и цель') + ' ' + icon('chevron') + '</button>';
    out += sideBlock('Смыслы', m, '');
  }
  if (S.screen >= 4) {
    out += sideBlock('Параметры', '<p class="side-text">' + S.min + '–' + S.max + ' слайдов<br>' + t(S.tone) + '</p><button class="side-link" data-panel="params">Изменить инструкции ' + icon('chevron') + '</button>');
  }
  if (S.screen >= 5 && S.job && S.job.plan_variants && S.job.plan_variants[S.plan]) {
    var pv = S.job.plan_variants[S.plan];
    out += sideBlock('План выступления', '<p class="side-text">Содержание ' + S.plan + ' · ' + t(pv.name) + '<br>' + pv.plan.slides.length + ' ' + plural(pv.plan.slides.length, ['слайд', 'слайда', 'слайдов']) + '</p>' +
      (S.stale ? '<div class="stale-label">Нужно обновить после правок</div>' : '') +
      '<button class="side-link" data-go="5">' + (S.screen === 5 ? 'Редактировать план' : 'Открыть план') + ' ' + icon('chevron') + '</button>', S.stale ? '' : (S.job.selected_plan ? 'Выбран' : ''), S.screen === 5);
  }
  if (S.screen >= 6 && S.design) {
    var jd = S.job.variant_descriptions || {};
    var nm = (jd[S.design] && jd[S.design].label) || (VARIANT_FALLBACK[S.design] || [''])[0];
    out += sideBlock('Дизайн', '<p class="side-text">Дизайн ' + VL[S.design] + ' · ' + t(nm) + '<br>Содержание ' + (S.job.selected_plan || S.plan) + ' сохранено</p><button class="side-link" data-go="' + (S.screen >= 7 ? 7 : 6) + '">' + (S.screen >= 7 ? 'Открыть слайды' : 'Изменить выбор') + ' ' + icon('chevron') + '</button>', S.screen >= 7 ? 'Выбран' : '', S.screen === 6 || S.screen === 7);
  }
  if (S.screen >= 8) {
    var st = S.audit === 'skipped' ? 'Пропущен по вашему решению' : S.audit === 'done' ? 'Проверки завершены' : S.audit === 'running' ? 'Идет проверка' : 'Еще не запускался';
    out += sideBlock('Аудит', '<p class="side-text">' + st + '</p><button class="side-link" data-go="8">' + (S.audit === 'done' ? 'Открыть результаты' : 'Провести аудит') + ' ' + icon('chevron') + '</button>', S.audit === 'done' ? 'Готово' : '', S.screen === 8);
  }
  return '<p class="sidebar-title">Ваша презентация</p>' + out;
}

/* ---------- экраны брифования ---------- */
function startView() {
  var tp = S.tpl; var busy = tp.status === 'uploading' || tp.status === 'analyzing';
  var label = tp.name ? t(tp.name) + ' <span>' + (busy ? 'загружаем и анализируем…' : tp.status === 'error' ? 'не удалось разобрать · выбрать другой' : '.pptx загружен · заменить') + '</span>' : 'Загрузите презентацию-ориентир по стилю <span>.pptx</span>';
  return '<div class="upload-shell ' + (tp.name ? 'loaded' : '') + ' ' + (tp.status === 'error' ? 'error' : '') + '"><button class="text-button" data-action="choose-template" aria-label="Выбрать презентацию-ориентир">' + icon('paperclip') + '</button>' +
    '<button class="upload-copy" data-action="choose-template">' + label + '</button>' +
    '<button class="round-next" data-action="start-next" aria-label="Перейти к описанию презентации" ' + (!tp.name || tp.status === 'error' ? 'disabled' : '') + '>' + icon('arrow') + '</button></div>' +
    (tp.status === 'error' ? '<p class="err" style="text-align:center">' + t(tp.error) + '</p>' : '');
}
function about() {
  return '<h2 class="form-title">О чем презентация?</h2><p class="subtext">Опишите, что хотите рассказать. Добавьте заметки, документы, таблицы или ссылки — все, из чего можно собрать содержание.</p>' +
    '<div class="composer"><label class="sr-only" for="topic">О чем хотите рассказать</label><textarea id="topic" placeholder="Например: хочу предложить новый подход к работе с клиентами…">' + raw(S.topic) + '</textarea>' +
    '<div class="composer-footer"><div class="attach-buttons"><button class="text-button" data-action="add-file">' + icon('paperclip') + 'Файл</button><button class="text-button" data-panel="add-link">' + icon('link') + 'Ссылка</button></div><span class="small">Текст, PDF, DOCX, XLSX</span></div></div>' +
    '<div class="files" id="files-box">' + fileRows() + '</div>' +
    '<div class="action-row"><button class="text-button" data-go="0">' + icon('back') + 'Назад</button>' + button('Далее: для кого и зачем', 'about-next') + '</div>';
}
function audienceView() {
  var c = S.clarify;
  var clarify = c ? '<div class="clarify"><div class="clarify-head">' + icon('spark') + 'Уточним смысл</div><p>' + t(c.question) + '</p><div class="chips">' +
    (c.options || []).map(function (o, i) { return '<button class="chip" data-clarify="' + i + '">' + t(o) + '</button>'; }).join('') +
    '<button class="chip" data-panel="clarification">Уточню своими словами</button></div></div>' : '';
  return '<h2 class="form-title">Для кого и зачем?</h2><p class="subtext">Поможем выстроить рассказ под конкретного человека и решение, которое он должен принять.</p>' +
    '<div class="field"><label for="audience">Кто будет смотреть презентацию?</label><input id="audience" value="' + raw(S.audience) + '" placeholder="Должность, роль, уровень знакомства с темой"><div class="chips">' +
    [['Руководство компании', 'Руководство'], ['Внутренняя команда', 'Команда'], ['Потенциальный клиент', 'Клиент'], ['Инвесторы', 'Инвесторы']].map(function (x) { return '<button class="chip" data-audience="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div></div>' +
    '<div class="field"><label for="goal">Что аудитория должна решить или сделать после?</label><textarea id="goal" rows="2">' + raw(S.goal) + '</textarea><div class="chips">' +
    [['Согласовать бюджет на пилот', 'Согласовать бюджет'], ['Принять предложенный план действий', 'Принять план'], ['Согласовать предложение и следующий шаг', 'Согласовать предложение']].map(function (x) { return '<button class="chip" data-goal="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div></div>' +
    clarify + (S.ideaError ? '<div class="err-box"><b>Не удалось определить главную идею</b>' + t(S.ideaError) + '</div>' : '') +
    '<div class="action-row"><button class="text-button" data-go="1">' + icon('back') + 'Назад</button>' +
    button(S.ideaBusy ? 'Определяем главную идею…' : (c ? 'Продолжить с уточнением' : 'Далее: объем и инструкции'), 'audience-next', 'primary', 'arrow', S.ideaBusy ? 'disabled' : '') + '</div>';
}
function paramsView() {
  return '<h2 class="form-title">Объем и дополнительные инструкции</h2><p class="subtext">Выберите длину рассказа. Тон и обязательные детали можно не заполнять.</p>' +
    '<div class="field"><label for="volume-min">Сколько слайдов?</label><div class="volume-row"><span>От</span><input id="volume-min" type="number" min="3" max="40" value="' + S.min + '"><span>до</span><input id="volume-max" aria-label="Максимальное число слайдов" type="number" min="3" max="40" value="' + S.max + '"><span>слайдов</span></div>' +
    '<div class="chips">' + [['5,7', 'Коротко · 5–7'], ['8,10', 'По делу · 8–10'], ['12,15', 'Подробно · 12–15']].map(function (v) { return '<button class="chip ' + (S.min + ',' + S.max === v[0] ? 'selected' : '') + '" data-volume="' + v[0] + '">' + v[1] + '</button>'; }).join('') + '</div></div>' +
    '<div class="field"><label for="tone">Тон изложения <span class="optional">необязательно</span></label><select id="tone">' + TONES.map(function (x) { return '<option ' + (S.tone === x ? 'selected' : '') + '>' + x + '</option>'; }).join('') + '</select></div>' +
    '<div class="field"><label for="must">Что обязательно должно быть на слайдах? <span class="optional">необязательно</span></label><textarea id="must" rows="2">' + raw(S.must) + '</textarea></div>' +
    '<div class="action-row"><button class="text-button" data-go="2">' + icon('back') + 'Назад</button>' + button(S.stale ? 'Обновить план' : 'Сформировать план', 'generate-plan', 'primary', 'spark') + '</div>' +
    '<p class="meaning-note" style="margin-top:16px;margin-left:auto">' + icon('info') + 'Проверьте главную идею справа. По ней построим три варианта содержания.</p>';
}
function loadingView() {
  var designing = S.runKind === 'design';
  if (S.runError) {
    return '<div class="load-center"><div class="load-indicator">' + icon('warning') + '</div><h1 class="form-title">Не получилось</h1><div class="err-box" style="text-align:left"><b>' + t(designing ? 'Не удалось собрать варианты дизайна' : 'Не удалось построить план выступления') + '</b>' + t(S.runError) + '</div>' +
      '<div class="action-row" style="justify-content:center"><button class="text-button" data-go="3">' + icon('back') + 'Вернуться к параметрам</button>' + button('Повторить', 'retry-run', 'primary', 'refresh') + '</div></div>';
  }
  var note = (S.job && S.job.busy_note) || '';
  return '<div class="load-center"><div class="load-indicator">' + icon('spark') + '</div><h1 class="form-title">' + (designing ? 'Собираем варианты дизайна' : 'Генерация плана выступления') + '</h1>' +
    '<p>' + (designing ? 'Раскладываем выбранное содержание по макетам шаблона. Получится три варианта оформления.' : 'Строим три способа рассказать вашу историю. У каждого слайда будет понятная роль в достижении цели.') + '</p>' +
    (note ? '<p class="progress-note">' + t(note) + '</p>' : '') +
    '<div class="load-skeleton"><div class="skeleton-card"></div><div class="skeleton-card"></div><div class="skeleton-card"></div></div></div>';
}

/* ---------- план ---------- */
function slideBody(sl) {
  var out = '';
  (sl.content_blocks || []).forEach(function (b) {
    if (b.bullets && b.bullets.length) out += '<ul class="plan-body">' + b.bullets.map(function (x) { return '<li>' + t(x) + '</li>'; }).join('') + '</ul>';
    else if (b.text) out += '<div class="plan-body"><p>' + t(b.text) + '</p></div>';
    else if (b.chart) out += '<div class="plan-body"><p>График: ' + t((b.chart.categories || []).join(', ')) + '</p></div>';
    else if (b.table) out += tablePreview(b.table, sl.data_note || (sl.source_table_id ? 'Таблица из материалов' : 'Таблица'));
  });
  return out;
}
function slideText(sl) {
  var lines = [];
  (sl.content_blocks || []).forEach(function (b) { if (b.bullets) lines = lines.concat(b.bullets); else if (b.text) lines.push(b.text); });
  return lines.join('\n');
}
function staleBanner() {
  return '<div class="stale-banner"><span>Бриф изменен. План и дизайн пока остались прежними.</span><button class="text-button" data-action="refresh-dependent">Обновить результаты ' + icon('refresh') + '</button></div>';
}
function planView() {
  var pvs = S.job.plan_variants || {};
  var keys = Object.keys(pvs).sort();
  if (!keys.length) return '<div class="err-box"><b>План не получен</b>Попробуйте сформировать его заново.</div>';
  if (!pvs[S.plan]) S.plan = keys[0];
  var P = pvs[S.plan];
  var sel = S.job.selected_plan === S.plan;
  var tabs = keys.map(function (k) {
    var p = pvs[k];
    return '<button class="variant-tab ' + (k === S.plan ? 'active' : '') + '" role="tab" aria-selected="' + (k === S.plan) + '" data-plan="' + k + '"><span class="letter">' + k + '</span><span><span class="tab-name">' + t(p.name) + '</span><small>' + p.plan.slides.length + ' ' + plural(p.plan.slides.length, ['слайд', 'слайда', 'слайдов']) + ' · содержание</small></span></button>';
  }).join('');
  var missing = ['A', 'B', 'C'].filter(function (k) { return !pvs[k]; });
  var slides = P.plan.slides.map(function (sl, i) {
    var busy = S.busySlides[S.plan + i];
    return '<article class="plan-slide ' + (busy ? 'busy' : '') + '"><div class="slide-num">' + String(i + 1).padStart(2, '0') + '</div><div><h2>' + t(sl.title) + '</h2>' + slideBody(sl) +
      (sl.why ? '<div class="purpose-line">' + icon('target') + '<span><b>Зачем этот слайд.</b> ' + t(sl.why) + '</span></div>' : '') +
      (vizReason(sl) ? '<div class="proof"><b>Как показаны данные.</b> ' + t(vizReason(sl)) + '</div>' : '') +
      (sl.data_gap ? '<div class="proof missing">' + t(sl.data_gap) + '</div>' : '') +
      '<div class="slide-controls"><button class="text-button" data-regen-slide="' + i + '">' + icon('refresh') + (busy ? 'Перегенерируем…' : 'Перегенерировать') + '</button><button class="text-button" data-edit-slide="' + i + '">' + icon('edit') + 'Уточнить контент</button></div></div></article>';
  }).join('');
  var gaps = P.plan.slides.filter(function (s) { return s.data_gap; }).length;
  var analysis = S.job.table_analysis;
  // Обоснование вида отображения (таблица / полосы / линия…) и рассмотренные альтернативы — из анализа данных.
  function vizReason(sl) {
    if (!sl.source_table_id) return '';
    var ta = S.job.table_analysis; var pool = ((ta && ta.artifacts) || []).concat(S.job.source_tables || []);
    for (var k = 0; k < pool.length; k++) {
      if (pool[k].table_id === sl.source_table_id && pool[k].presentation_reason) {
        var r = deyo(pool[k].presentation_reason).replace(/\s+/g, ' ').trim();
        return r.length > 320 ? r.slice(0, 317).replace(/[\s,;:—-]+\S*$/, '') + '…' : r;
      }
    }
    return '';
  }
  var analysisNote = analysis && ['skipped', 'disabled'].indexOf(analysis.status) < 0 ? '<div class="proof"><b>Анализ таблиц.</b> ' +
    ((analysis.artifacts || []).length ? 'Подготовлено результатов: ' + analysis.artifacts.length + '. Доступны всем вариантам плана. ' : 'Дополнительные расчеты не подготовлены. ') +
    (analysis.notes || []).map(function (note) { return raw(note); }).join(' ') + '</div>' : '';
  return '<div class="work-head"><div><h1>План выступления</h1><p>Сначала согласуем содержание. Дизайн выберем на следующем этапе.</p></div></div>' + (S.stale ? staleBanner() : '') +
    analysisNote +
    (missing.length && missing.length < 3 ? '<div class="stale-banner"><span>' + (missing.length === 1 ? 'Вариант ' + missing[0] + ' не удалось построить.' : 'Варианты ' + missing.join(' и ') + ' не удалось построить.') + '</span></div>' : '') +
    '<div class="variants-tabs" role="tablist" aria-label="Варианты содержания" style="grid-template-columns:repeat(' + keys.length + ',1fr)">' + tabs + '</div><p class="story-line">' + t(P.story || P.description) + '</p>' +
    '<div class="plan-toolbar"><span class="small">Заголовок → содержание → роль слайда</span><button class="text-button" data-action="regenerate-plan" ' + (S.planBusy ? 'disabled' : '') + '>' + icon('refresh') + (S.planBusy ? 'Строим другой вариант…' : 'Другой вариант ' + S.plan) + '</button></div>' +
    '<div class="slide-list">' + slides + '</div>' +
    '<div class="sticky-action"><span class="small">Содержание ' + S.plan + ' · ' + P.plan.slides.length + ' ' + plural(P.plan.slides.length, ['слайд', 'слайда', 'слайдов']) + '<br>' + (gaps ? 'Недостающие данные отмечены, цифры не выдумываем' : 'Все данные для слайдов есть в материалах') + '</span>' +
    button(sel && !S.job.designs_stale ? 'Продолжить с содержанием ' + S.plan : 'Выбрать содержание ' + S.plan + ' и перейти к дизайну', 'plan-next') + '</div>';
}

/* ---------- дизайн ---------- */
function variantKeys() { return Object.keys((S.job && S.job.variant_paths) || {}).sort(); }
function vInfo(key) {
  var d = ((S.job.variant_descriptions || {})[key]) || {}; var fb = VARIANT_FALLBACK[key] || ['', ''];
  return { label: deyo(d.label || fb[0]), desc: fb[1] };
}
function hasSections(key) { return ((S.job.variant_kinds || {})[key] || []).indexOf('section') >= 0; }
function srcMap(key) {
  var src = (S.job.variant_sources || {})[key]; var map = {};
  var n = ((S.job.variant_kinds || {})[key] || []).length;
  if (src && src.length) src.forEach(function (p, i) { if (map[p] === undefined) map[p] = i; }); else for (var i = 0; i < n; i++) map[i] = i;
  return map;
}
function diffCount(key) {
  var plans = S.job.variant_plans || {}; var mine = plans[key] || []; var my = srcMap(key);
  var others = variantKeys().filter(function (k) { return k !== key; });
  var om = {}; others.forEach(function (o) { om[o] = srcMap(o); });
  var n = 0;
  Object.keys(my).forEach(function (p) {
    var id = mine[my[p]];
    if (others.length && others.every(function (o) { var j = om[o][p]; return j === undefined || (plans[o] || [])[j] !== id; })) n++;
  });
  return n;
}
function planTotal() { var p = S.job.content_plan; return p && p.slides ? p.slides.length : 0; }
function variantNotes(key) { return ((S.job && S.job.variant_notes) || {})[key] || []; }
function notesHtml(key) {
  var n = variantNotes(key); if (!n.length) return '';
  return '<p class="variant-note" title="' + n.map(deyo).join('\n').replace(/"/g, '&quot;') + '">' + icon('refresh') + t(n.length === 1 ? 'Один слайд пересобран автоматически: композиция не поместилась в шаблон' : n.length + ' слайда пересобраны автоматически: композиции не поместились в шаблон') + '</p>';
}
function variantErrorsHtml() {
  var e = (S.job && S.job.variant_errors) || {}; var keys = Object.keys(e).sort(); if (!keys.length) return '';
  return '<div class="variant-errors">' + icon('warning') + '<span>' + t(keys.map(function (k) { return 'Дизайн ' + VL[k] + ' не собрался в этом шаблоне (' + deyo(e[k]) + ')'; }).join('. ') + '. Остальные варианты готовы, их можно выбирать.') + '</span></div>';
}
function variantInfoCard(key) {
  var inf = vInfo(key); var total = ((S.job.variant_plans || {})[key] || []).length;
  var anySec = variantKeys().some(hasSections);
  return '<div class="variant-info"><div><span class="name">Дизайн ' + VL[key] + ' · ' + t(inf.label) + '</span><p>' + t(inf.desc) + '</p><div class="chips-static"><span>' + total + ' ' + plural(total, ['слайд', 'слайда', 'слайдов']) + (anySec ? (hasSections(key) ? ' · с разделителями' : ' · без разделителей') : '') + '</span><span>отличий в композиции: ' + diffCount(key) + ' из ' + planTotal() + '</span></div>' + notesHtml(key) + '</div><span class="variant-check" aria-hidden="true">✓</span></div>';
}
function shotHtml(key, i) {
  var u = (S.previews[key] || [])[i];
  return u ? '<div class="deck-shot"><img src="' + API + u + '?v=' + S.stamp + '" alt="Дизайн ' + VL[key] + ', слайд ' + (i + 1) + '" loading="lazy"><span class="ds-num">' + (i + 1) + '</span></div>' : '<div class="deck-shot placeholder">Загружаем…</div>';
}
function deckColumn(key) {
  if (S.previewErr[key]) return '<div>' + t('Не удалось загрузить предпросмотр: ' + S.previewErr[key]) + '</div>';
  var n = (S.previews[key] || []).length;
  if (!n) return '<div class="deck-shot placeholder">Загружаем предпросмотр…</div>';
  var out = ''; for (var i = 0; i < n; i++) out += shotHtml(key, i); return out;
}
function designsView() {
  var keys = variantKeys(); var m = S.designMode;
  var inf = S.design ? vInfo(S.design) : null;
  var body = m === 'columns' ?
    '<div class="deck-columns" id="deck-columns">' + keys.map(function (k) {
      return '<article class="deck-card ' + (S.design === k ? 'selected' : '') + ' ' + (S.previewErr[k] ? 'err-card' : '') + '" data-design="' + k + '" tabindex="0" role="button" aria-pressed="' + (S.design === k) + '">' + variantInfoCard(k) + '<div class="deck-scroll" data-scroll>' + deckColumn(k) + '</div>' +
        '<a class="text-button deck-download" href="' + API + '/jobs/' + S.jobId + '/export?format=pptx&variant=' + k + '" target="_blank" rel="noopener">' + icon('download') + 'Скачать .pptx этого варианта</a></article>';
    }).join('') + '</div>' : '<div id="compare-area">' + compareView() + '</div>';
  var cnt = keys.length === 3 ? 'Три презентации' : (keys.length === 2 ? 'Две презентации' : 'Презентация');
  return '<div class="work-head"><div><h1>Выберите вариант оформления</h1><p>' + cnt + ' по одному плану содержания ' + (S.job.selected_plan || S.plan) + '. Отличаются плотностью подачи и характером композиции.</p></div><div class="seg" role="group" aria-label="Режим просмотра"><button class="' + (m === 'columns' ? 'on' : '') + '" data-mode="columns" aria-pressed="' + (m === 'columns') + '">Колонки</button><button class="' + (m === 'compare' ? 'on' : '') + '" data-mode="compare" aria-pressed="' + (m === 'compare') + '">Слайд за слайдом</button></div></div>' +
    (S.stale ? staleBanner() : '') + variantErrorsHtml() + body +
    '<div class="sticky-action"><span class="small">' + (S.design ? 'Выбран дизайн ' + VL[S.design] + ' · ' + t(inf.label) : 'Нажмите на колоду или карточку, чтобы выбрать вариант') + '</span>' +
    button('Выбрать и продолжить', 'design-next', 'primary', 'arrow', S.design ? '' : 'disabled') + '</div>';
}
function compareView() {
  var keys = variantKeys(); var total = planTotal() || Math.max.apply(null, keys.map(function (k) { return (S.previews[k] || []).length; }).concat([0]));
  if (!total) return '<p class="small">Предпросмотр готовится…</p>';
  var i = S.cmp = Math.max(0, Math.min(S.cmp, total - 1));
  var planSlide = ((S.job.content_plan || {}).slides || [])[i] || {};
  var maps = {}; keys.forEach(function (k) { maps[k] = srcMap(k); });
  var cards = keys.map(function (k) {
    var vi = maps[k][i]; var inf = vInfo(k);
    var kind = vi === undefined ? '' : kindRu(((S.job.variant_kinds || {})[k] || [])[vi]);
    return '<article class="compare-card ' + (S.design === k ? 'selected' : '') + '" data-design="' + k + '" tabindex="0" role="button" aria-pressed="' + (S.design === k) + '"><div class="cap"><b>Дизайн ' + VL[k] + ' · ' + t(inf.label) + '</b><span>' + (vi === undefined ? '—' : kind + ' · слайд ' + (vi + 1)) + '</span></div>' +
      (vi === undefined ? '<div class="cmp-missing">В этом варианте слайда нет: колода короче.</div>' : shotHtml(k, vi)) + '</article>';
  }).join('');
  var dots = ''; for (var d = 0; d < total; d++) dots += '<i class="' + (d === i ? 'on' : '') + '" data-cmp-to="' + d + '" title="Слайд ' + (d + 1) + '"></i>';
  return '<div class="compare-nav"><button data-cmp="-1" aria-label="Предыдущий слайд" ' + (i === 0 ? 'disabled' : '') + '>‹</button><span>Слайд плана <b>' + (i + 1) + '</b> из ' + total + (planSlide.title ? ' · ' + t(planSlide.title) : '') + '</span><div class="strip">' + dots + '</div><button data-cmp="1" aria-label="Следующий слайд" ' + (i >= total - 1 ? 'disabled' : '') + '>›</button></div><div class="compare-grid" style="grid-template-columns:repeat(' + keys.length + ',minmax(0,1fr))">' + cards + '</div><p class="small" style="margin-top:12px">Стрелки ← → на клавиатуре листают слайды. Клик по карточке выбирает вариант.</p>';
}
async function loadPreviews() {
  var keys = variantKeys(); var jid = S.jobId;
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (S.previews[k] && S.previews[k].length) continue;
    try { var r = await api('/jobs/' + jid + '/variants/' + k + '/preview'); S.previews[k] = r.slides || []; delete S.previewErr[k]; }
    catch (e) { S.previewErr[k] = e.message; }
    if (jid !== S.jobId) return;
    refreshDesigns();
  }
}
function refreshDesigns() {
  if (S.screen !== 6 && S.screen !== 7) return;
  if (S.screen === 7) { var tb = $('thumbs-box'); if (tb) tb.innerHTML = thumbsHtml(); var cv = $('canvas-box'); if (cv && !S.refineBusy) cv.innerHTML = canvasHtml(); return; }
  if (S.designMode === 'compare') { var ca = $('compare-area'); if (ca) ca.innerHTML = compareView(); return; }
  var cols = document.querySelectorAll('.deck-card');
  cols.forEach(function (c) {
    var k = c.getAttribute('data-design'); var sc = c.querySelector('.deck-scroll'); var top = sc.scrollTop;
    sc.innerHTML = deckColumn(k); sc.scrollTop = top; c.classList.toggle('err-card', !!S.previewErr[k]);
  });
}

/* ---------- доработка слайдов ---------- */
function refineBase() {
  var v = S.design; var kinds = (S.job.variant_kinds || {})[v] || []; var n = Math.max(kinds.length, (S.previews[v] || []).length);
  return { v: v, kinds: kinds, n: n };
}
function thumbsHtml() {
  var b = refineBase(); var out = '';
  for (var i = 0; i < b.n; i++) {
    var u = (S.previews[b.v] || [])[i];
    out += '<button class="thumb-button ' + (i === S.selectedSlide ? 'selected' : '') + '" data-slide="' + i + '" aria-label="Слайд ' + (i + 1) + '" aria-pressed="' + (i === S.selectedSlide) + '">' +
      (u ? '<img class="thumb-img" src="' + API + u + '?v=' + S.stamp + '" alt="" loading="lazy">' : '<div class="thumb-img" style="height:60px"></div>') + '<div class="thumb-number">' + (i + 1) + ' / ' + b.n + '</div></button>';
  }
  return out;
}
function planIndexOf(i) { var src = ((S.job.variant_sources || {})[S.design] || [])[i]; return src === undefined ? i : src; }
function canvasHtml() {
  var i = S.selectedSlide; var u = (S.previews[S.design] || [])[i];
  return '<div class="slide-canvas shot ' + (S.refineBusy ? 'busy' : '') + '" data-busy="' + raw(S.refineBusy) + '">' + (u ? '<img src="' + API + u + '?v=' + S.stamp + '" alt="Слайд ' + (i + 1) + '">' : '') + '</div>';
}
function refineView() {
  var b = refineBase(); var i = S.selectedSlide = Math.min(S.selectedSlide, Math.max(0, b.n - 1));
  var kind = b.kinds[i]; var isSection = kind === 'section';
  var ps = ((S.job.content_plan || {}).slides || [])[planIndexOf(i)];
  var inf = vInfo(S.design);
  return '<div class="work-head"><div><h1>Доработайте отдельные слайды</h1><p>Меняйте текст или компоновку отдельного слайда в рамках выбранного дизайна — остальная презентация не пересобирается.</p></div><button class="text-button" data-go="6">' + icon('back') + 'К выбору дизайна</button></div>' + (S.stale ? staleBanner() : '') +
    '<div class="editor-shell"><nav class="slide-thumbs" id="thumbs-box" aria-label="Слайды">' + thumbsHtml() + '</nav><section><div id="canvas-box">' + canvasHtml() + '</div>' +
    '<p class="compose-note">Слайд ' + (i + 1) + ' из ' + b.n + ' · ' + t(kindRu(kind)) + ' · дизайн ' + VL[S.design] + ' · ' + t(inf.label) + '</p>' +
    (isSection ? '<p class="small" style="margin-top:12px">Разделитель собирается автоматически из плана. Его заголовок можно изменить в плане выступления.</p>' :
      '<div class="edit-actions">' + button('Перегенерировать текст', 'rewrite', 'secondary', 'refresh', S.refineBusy ? 'disabled' : '') + button('Перекомпоновать макет', 'recompose', 'secondary', 'grid', S.refineBusy ? 'disabled' : '') + button('Уточнить вручную', 'edit-current', 'secondary', 'edit', S.refineBusy ? 'disabled' : '') + '</div>') +
    (ps && ps.why ? '<div class="slide-rationale"><div class="purpose-line">' + icon('target') + '<span><b>Работает на цель.</b> ' + t(ps.why) + '</span></div></div>' : '') +
    '<div class="inline-feedback" id="slide-feedback"></div><div class="action-row"><span class="small">Содержание ' + (S.job.selected_plan || S.plan) + ' · дизайн ' + VL[S.design] + '</span>' + button('К проверке презентации', 'refine-next') + '</div></section></div>';
}
async function runSlideEdit(action, extra) {
  if (S.refineBusy) return;
  var i = S.selectedSlide;
  S.refineBusy = action === 'recompose' ? 'Перекомпоновываем…' : action === 'rewrite' ? 'Переписываем текст…' : 'Сохраняем…';
  var cv = $('canvas-box'); if (cv) cv.innerHTML = canvasHtml();
  document.querySelectorAll('.edit-actions button').forEach(function (b) { b.disabled = true; });
  var fb = $('slide-feedback'); if (fb) fb.textContent = '';
  try {
    var body = Object.assign({ action: action }, extra || {});
    var res = await postJson('/jobs/' + S.jobId + '/variants/' + S.design + '/slides/' + i + '/edit', body);
    S.job = await api('/jobs/' + S.jobId); S.stamp = Date.now();
    S.audit = S.audit === 'done' ? 'stale' : S.audit;
    try { var pr = await api('/jobs/' + S.jobId + '/variants/' + S.design + '/preview'); S.previews[S.design] = pr.slides || []; } catch (e) { /* превью подгрузятся сами */ }
    S.refineBusy = '';
    render();
    var f2 = $('slide-feedback');
    if (f2) f2.textContent = action === 'recompose' ? (res.after === res.before ? 'Других компоновок под этот слайд в шаблоне нет.' : 'Макет перекомпонован. Дизайн-система, текст и остальные слайды не изменились.') : action === 'rewrite' ? 'Текст слайда переписан. Остальные слайды не изменились.' : 'Текст слайда обновлен.';
  } catch (e) {
    S.refineBusy = ''; render(); notify('Не удалось изменить слайд: ' + e.message, true);
  }
}

/* ---------- аудит ---------- */
function auditView() {
  if (S.audit === 'running') {
    return '<div class="work-head"><div><h1>Проверяем презентацию…</h1><p>Сверяем рассказ с главной идеей и целью, затем проверяем верстку каждого слайда.</p></div></div><div class="load-center" style="padding-top:var(--space-6)"><div class="load-indicator"><span class="spinner" style="width:26px;height:26px"></span></div><div class="load-skeleton"><div class="skeleton-card"></div><div class="skeleton-card"></div><div class="skeleton-card"></div></div></div>';
  }
  if (S.audit === 'done' || S.audit === 'stale') return auditDone();
  var jobErr = S.runError && S.runKind === 'audit' ? '<div class="err-box"><b>Аудит не выполнен</b>' + t(S.runError) + '</div>' : '';
  return '<div class="work-head"><div><h1>Проверим презентацию перед экспортом?</h1><p>Сопоставим содержание с главной идеей и целью, затем проверим верстку.</p></div></div>' + jobErr +
    '<div class="audit-choices"><button class="audit-choice recommended" data-action="run-audit">' + icon('shield') + '<h2>Провести аудит</h2><p>Главная идея, логика аргументов, действие аудитории и качество слайдов.</p></button>' +
    '<button class="audit-choice" data-action="skip-audit">' + icon('arrow') + '<h2>Пропустить аудит</h2><p>Перейти к экспорту сейчас. К проверке можно вернуться из правой панели.</p></button></div>';
}
function auditDone() {
  var al = S.job.alignment; var rows = '';
  function row(ok, title, text) { return '<div class="align-row ' + (ok ? 'ok' : 'warn') + '">' + icon(ok ? 'check' : 'warning') + '<div><b>' + t(title) + '</b><span>' + t(text) + '</span></div></div>'; }
  if (al) {
    if (al.idea_through_story) rows += row(al.idea_through_story.ok, 'Главная идея проходит через весь рассказ', al.idea_through_story.comment);
    if (al.final_action) rows += row(al.final_action.ok, 'В финале есть конкретное действие аудитории', al.final_action.comment);
    (al.slide_remarks || []).forEach(function (r) { rows += row(false, 'Слайд ' + r.slide + ' плана', r.remark); });
  }
  return '<div class="work-head"><div><h1>Проверка завершена</h1><p>Проверяем не только верстку, но и то, помогает ли рассказ достичь цели.</p></div></div>' +
    (S.audit === 'stale' ? '<div class="stale-banner"><span>Слайды менялись после аудита — замечания ниже могут не соответствовать.</span><button class="text-button" data-action="run-audit">Повторить аудит ' + icon('refresh') + '</button></div>' : '') +
    (rows ? '<div class="align-list">' + rows + '</div>' : '') +
    '<div class="audit-block legacy"><h2>Верстка слайдов</h2><div id="auditBody"><p class="small">Загружаем карту аудита…</p></div></div>' +
    '<div class="action-row"><button class="text-button" data-go="7">' + icon('back') + 'Доработать слайды</button>' + button('Перейти к экспорту', 'export-next') + '</div>';
}
function exportTitle() {
  var v = S.job.plan_variants && S.job.plan_variants[S.job.selected_plan || S.plan];
  var first = v && v.plan.slides[0];
  return (first && first.title) || 'Презентация';
}
function exportView() {
  var fm = [['pptx', 'PPTX', 'Редактировать в PowerPoint'], ['pdf', 'PDF', 'Отправить или показать'], ['html', 'HTML', 'Открыть в браузере']];
  var n = ((S.job.variant_plans || {})[S.design] || []).length;
  var st = S.audit === 'skipped' ? 'Аудит пропущен. Вы можете запустить его из правой панели.' : S.audit === 'done' ? 'Аудит завершен.' : S.audit === 'stale' ? 'Слайды менялись после аудита.' : 'Аудит не запускался. Его можно провести до экспорта.';
  return '<div class="work-head"><div><h1>Презентация готова к экспорту</h1><p>' + t(exportTitle()) + ' · содержание ' + (S.job.selected_plan || S.plan) + ' · дизайн ' + VL[S.design] + ' · ' + n + ' ' + plural(n, ['слайд', 'слайда', 'слайдов']) + '</p></div></div>' +
    '<div class="export-formats">' + fm.map(function (f) { return '<button class="format-card ' + (S.format === f[0] ? 'selected' : '') + '" data-format="' + f[0] + '" aria-pressed="' + (S.format === f[0]) + '"><b>' + f[1] + '</b><p>' + f[2] + '</p></button>'; }).join('') + '</div>' +
    '<div class="purpose-line">' + icon('target') + '<span><b>Цель презентации.</b> ' + t(S.goal) + '</span></div><p class="small" style="margin-top:20px">' + st + '</p>' +
    '<div class="action-row"><button class="text-button" data-go="7">' + icon('back') + 'Вернуться к слайдам</button>' + button('Скачать .' + S.format, 'download', 'primary', 'download') + '</div><p class="inline-feedback" id="export-feedback"></p>';
}

function mainContent() {
  switch (S.screen) {
    case 0: return startView(); case 1: return about(); case 2: return audienceView(); case 3: return paramsView();
    case 4: return loadingView(); case 5: return S.job ? planView() : ''; case 6: return S.job ? designsView() : '';
    case 7: return S.job ? refineView() : ''; case 8: return S.job ? auditView() : ''; default: return S.job ? exportView() : '';
  }
}
function afterRender() {
  if (S.screen === 6) { bindScrollSync(); loadPreviews(); }
  if (S.screen === 7) loadPreviews();
  if (S.screen === 8 && (S.audit === 'done' || S.audit === 'stale')) loadAuditMap();
}
function bindScrollSync() {
  var cols = Array.prototype.slice.call(document.querySelectorAll('[data-scroll]')); var busy = false;
  cols.forEach(function (c) {
    c.addEventListener('scroll', function () {
      if (busy) return; busy = true;
      var ratio = c.scrollTop;
      cols.forEach(function (o) { if (o !== c) o.scrollTop = ratio; });
      busy = false;
    });
  });
}

/* ---------- панели ---------- */
function openPanel(name, index) {
  remember();
  var root = $('overlay'); var title = '', sub = '', inner = '', cls = '';
  if (name === 'design-system') {
    title = 'Дизайн-система'; sub = 'Стиль презентации-ориентира: ' + S.tpl.name; cls = 'xwide';
    inner = '<div class="legacy" id="ds-body"><p class="small">Загружаем…</p></div><div class="drawer-action" style="margin-top:24px">' + button('Заменить ориентир', 'replace-template', 'secondary', 'paperclip') + '</div>';
  } else if (name === 'materials') {
    title = 'Материалы презентации'; sub = 'Отдельно от шаблона, который задает только стиль';
    inner = '<div class="field"><label for="panel-topic">Что рассказываем</label><textarea id="panel-topic">' + raw(S.topic) + '</textarea></div><div class="files" id="files-box-panel">' + fileRows() + '</div>' +
      '<div class="attach-buttons"><button class="text-button" data-action="add-file">' + icon('paperclip') + 'Добавить файл</button><button class="text-button" data-panel="add-link">' + icon('link') + 'Добавить ссылку</button></div>' +
      '<div class="drawer-action">' + button('Сохранить материалы', 'save-materials', 'primary', 'check') + '</div>';
  } else if (name === 'meaning') {
    title = 'Смыслы презентации'; sub = 'Главная идея — утверждение, которое аудитория должна унести с собой';
    inner = '<div class="field"><label for="panel-idea">Главная идея</label><textarea id="panel-idea">' + raw(S.idea) + '</textarea></div>' +
      '<div class="field"><label for="panel-audience">Для кого</label><textarea id="panel-audience" rows="2">' + raw(S.audience) + '</textarea></div>' +
      '<div class="field"><label for="panel-goal">Что аудитория должна решить или сделать</label><textarea id="panel-goal">' + raw(S.goal) + '</textarea></div>' +
      '<button class="text-button" data-action="regenerate-idea">' + icon('refresh') + 'Предложить другую формулировку</button>' +
      (S.job ? '<p class="example-note">После сохранения план и дизайн останутся на месте. Мы пометим, что их нужно обновить, и спросим перед пересборкой.</p>' : '') +
      '<div class="drawer-action">' + button('Сохранить смыслы', 'save-meaning', 'primary', 'check') + '</div>';
  } else if (name === 'params') {
    title = 'Объем и инструкции'; sub = 'Параметры будущей презентации';
    inner = paramsView().replace(/<h2 class="form-title">[\s\S]*?<\/p>/, '').replace(/<div class="action-row">[\s\S]*$/, '<div class="drawer-action">' + button('Сохранить параметры', 'save-params', 'primary', 'check') + '</div>');
  } else if (name === 'slide-edit') {
    var sl = planSlides()[index]; S.editIndex = index; S.editScope = 'plan';
    title = 'Уточнить контент · слайд ' + (index + 1); sub = 'Правки относятся к содержанию ' + S.plan;
    inner = '<div class="field"><label for="slide-title-edit">Заголовок — главный тезис слайда</label><textarea id="slide-title-edit" rows="2">' + raw(sl.title) + '</textarea></div>' +
      '<div class="field"><label for="slide-body-edit">Что говорим на слайде (каждая строка — отдельный пункт)</label><textarea id="slide-body-edit" rows="6">' + raw(slideText(sl)) + '</textarea></div>' +
      '<div class="field"><label for="slide-purpose-edit">Как слайд работает на цель</label><textarea id="slide-purpose-edit">' + raw(sl.why || '') + '</textarea></div>' +
      '<div class="field"><label for="slide-instruction">Или дайте инструкцию для перегенерации</label><textarea id="slide-instruction" placeholder="Например: сделать акцент на рисках и добавить аргумент для IT"></textarea></div>' +
      '<div class="drawer-action">' + button('Сохранить вручную', 'save-slide', 'primary', 'check') + '<button class="text-button" style="margin-top:12px" data-action="apply-slide-instruction">' + icon('spark') + 'Перегенерировать по инструкции</button></div>';
  } else if (name === 'deck-slide-edit') {
    var psl = ((S.job.content_plan || {}).slides || [])[planIndexOf(index)] || { title: '', content_blocks: [] };
    title = 'Уточнить вручную · слайд ' + (index + 1); sub = 'Текст слайда выбранного дизайна';
    inner = '<div class="field"><label for="deck-title-edit">Заголовок</label><textarea id="deck-title-edit" rows="2">' + raw(psl.title) + '</textarea></div>' +
      '<div class="field"><label for="deck-body-edit">Текст (каждая строка — отдельный пункт)</label><textarea id="deck-body-edit" rows="6">' + raw(slideText(psl)) + '</textarea></div>' +
      '<div class="field"><label for="deck-instruction">Или дайте инструкцию для перегенерации</label><textarea id="deck-instruction" placeholder="Например: короче и без терминов"></textarea></div>' +
      '<div class="drawer-action">' + button('Сохранить вручную', 'save-deck-slide', 'primary', 'check') + '<button class="text-button" style="margin-top:12px" data-action="deck-instruction">' + icon('spark') + 'Перегенерировать по инструкции</button></div>';
  } else if (name === 'add-link') {
    title = 'Добавить ссылку'; sub = 'Материал для содержания, не ориентир по стилю';
    inner = '<div class="field"><label for="new-link">Адрес страницы или документа</label><input id="new-link" type="url" placeholder="https://…"></div><p class="example-note">Страницу прочитает сервер. Если она закрыта или недоступна, мы скажем об этом, а не будем считать ее прочитанной.</p><div class="drawer-action">' + button('Добавить ссылку', 'save-link', 'primary', 'check') + '</div>';
  } else if (name === 'clarification') {
    title = 'Уточните своими словами'; sub = (S.clarify && S.clarify.question) || 'Одно уточнение, которое меняет смысл рассказа';
    inner = '<div class="field"><label for="clarification-text">Ваш ответ</label><textarea id="clarification-text" placeholder="Например: сначала нужен бюджет только на обследование"></textarea></div><div class="drawer-action">' + button('Применить', 'save-clarification', 'primary', 'check') + '</div>';
  } else if (name === 'rebuild') {
    title = 'Обновить зависимые результаты?'; sub = 'Вы изменили исходный бриф';
    inner = '<p>Новые данные повлияют на аргументы и содержание слайдов.</p><h3>Что обновим</h3><ul><li>Три варианта плана выступления.</li><li>После выбора плана - варианты дизайна и аудит.</li></ul><p class="example-note">Ручные правки текущего плана при пересборке не переносятся. Скачайте нужные файлы, если они еще нужны.</p><div class="drawer-action">' + button('Обновить результаты', 'confirm-rebuild', 'primary', 'refresh') + '</div>';
  } else if (name === 'ds-confirm') { return; }
  root.innerHTML = '<div class="backdrop"><section class="drawer ' + cls + '" role="dialog" aria-modal="true" aria-label="' + raw(title) + '"><div class="drawer-head"><div><h2>' + t(title) + '</h2><p>' + t(sub) + '</p></div><button class="close" data-action="close" aria-label="Закрыть">×</button></div>' + inner + '</section></div>';
  setTimeout(function () { var f = root.querySelector('textarea,input,button.primary'); if (f) f.focus(); }, 20);
  if (name === 'design-system') fillDesignSystem();
}
function closePanel() { $('overlay').innerHTML = ''; }
async function fillDesignSystem() {
  var box = $('ds-body'); if (!box) return;
  try {
    if (!S.tpl.ds) S.tpl.ds = await api('/templates/' + S.tpl.id + '/design-system');
    window.renderDesignSystem(S.tpl.ds, box);
  } catch (e) { box.innerHTML = '<p class="err">Не удалось загрузить дизайн-систему: ' + t(e.message) + '</p>'; }
}

/* ---------- шаг 1: шаблон ---------- */
function chooseTemplate() {
  var input = document.createElement('input'); input.type = 'file'; input.accept = '.pptx';
  input.onchange = function () {
    var f = input.files[0]; if (!f) return;
    if (!/\.pptx$/i.test(f.name)) { notify('Нужен файл в формате .pptx. Другие форматы ориентира не поддерживаются.', true); return; }
    uploadTemplate(f);
  };
  input.click();
}
function uploadTemplate(f) {
  var replacing = S.job && S.tpl.id;
  S.tpl = { name: f.name, id: '', status: 'uploading', error: '', ds: null, promise: null };
  var fd = new FormData(); fd.append('template', f);
  S.tpl.promise = (async function () {
    try {
      var a = await api('/templates/analyze', { method: 'POST', body: fd });
      if (a.status === 'failed') throw new Error(a.error || 'Не удалось разобрать презентацию-ориентир.');
      S.tpl.id = a.template_analysis_id; S.tpl.status = 'analyzing';
      try { S.tpl.ds = await api('/templates/' + a.template_analysis_id + '/design-system'); } catch (e) { /* дизайн-система подгрузится при открытии */ }
      S.tpl.status = 'ready'; if (replacing) markChanged();
    } catch (e) {
      S.tpl.status = 'error';
      var m = e.message || String(e);
      if (/408|Failed to fetch|Нет связи/i.test(m)) m = 'Файл не дошел до сервера. Проверьте соединение и загрузите его еще раз.';
      S.tpl.error = m;
    }
    updateSide();
    if (S.screen === 0) { var fa = $('form-area'); if (fa) fa.innerHTML = startView(); }
  })();
  if (S.screen === 0) go(1); else updateSide();
}

/* ---------- главная идея ---------- */
function ideaPayload(extra) {
  var mt = materialsText();
  return Object.assign({ topic: S.topic, materials_text: mt.slice(0, 40000), audience: S.audience, goal: S.goal }, extra || {});
}
async function requestIdea(extra) {
  S.ideaBusy = true; S.ideaError = ''; updateSide();
  try {
    var r = await postJson('/briefing/idea', ideaPayload(extra));
    S.idea = deyo(r.idea || '');
    S.clarify = r.clarification && r.clarification.question ? r.clarification : null;
    S.ideaBusy = false;
    return true;
  } catch (e) {
    S.ideaBusy = false; S.ideaError = e.message; return false;
  }
}

/* ---------- запуск генерации ---------- */
async function startPlanning() {
  S.runError = ''; S.runKind = 'plan'; var run = ++S.runId;
  S.stale = false; S.job = null; S.previews = {}; S.previewErr = {}; S.design = ''; S.audit = 'choice'; S.auditMap = null; S.maxScreen = Math.min(S.maxScreen, 4);
  go(4);
  try {
    if (S.tpl.promise) await S.tpl.promise;
    if (S.tpl.status !== 'ready') throw new Error(S.tpl.error || 'Презентация-ориентир не готова. Загрузите ее еще раз.');
    var fd = new FormData();
    fd.append('template_analysis_id', S.tpl.id); fd.append('brief', S.topic); fd.append('idea', S.idea);
    fd.append('audience', S.audience); fd.append('goal', S.goal); fd.append('tone', S.tone); fd.append('must_have', S.must);
    fd.append('materials_text', materialsText()); fd.append('materials_json', JSON.stringify(materialsMeta()));
    fd.append('material_ids_json', JSON.stringify(materialIds()));
    fd.append('slide_count', String(S.min)); fd.append('slide_count_max', String(S.max));
    var r = await api('/jobs', { method: 'POST', body: fd });
    S.jobId = r.job_id;
    S.usedBrief = briefSignature();
    var job = await pollJob(r.job_id, run, ['awaiting_plan_choice']);
    if (!job) return;
    S.job = job; S.plan = Object.keys(job.plan_variants || {}).sort()[0] || 'A';
    go(5);
  } catch (e) { if (run === S.runId) { S.runError = e.message; var fa = $('form-area'); if (fa && S.screen === 4) fa.innerHTML = loadingView(); } }
}
async function pollJob(jobId, run, targets) {
  for (;;) {
    await sleep(2000);
    if (run !== S.runId) return null;
    var job;
    try { job = await api('/jobs/' + jobId); } catch (e) { continue; }
    if (job.status === 'failed') throw new Error(job.error || 'Генерация завершилась ошибкой.');
    if (targets.indexOf(job.status) >= 0) return job;
    var fa = $('form-area');
    if (job.busy_note !== undefined && S.screen === 4 && fa) { S.job = job; var p = fa.querySelector('.progress-note'); if (p) p.textContent = job.busy_note; }
  }
}
async function selectPlan() {
  var key = S.plan; var run = ++S.runId; S.runKind = 'design'; S.runError = '';
  try {
    var r = await postJson('/jobs/' + S.jobId + '/select-plan', { plan: key });
    if (r.rebuilt) {
      S.previews = {}; S.previewErr = {}; S.design = ''; S.audit = 'choice'; S.auditMap = null;
      go(4);
      var job = await pollJob(S.jobId, run, ['awaiting_variant_choice']);
      if (!job) return; S.job = job;
    } else { S.job = await api('/jobs/' + S.jobId); }
    S.job.selected_plan = key;
    if (!S.design && S.job.selected_variant && (S.job.variant_paths || {})[S.job.selected_variant]) S.design = S.job.selected_variant;
    go(6);
  } catch (e) {
    if (S.screen === 4) { S.runError = e.message; $('form-area').innerHTML = loadingView(); } else notify('Не удалось выбрать содержание: ' + e.message, true);
  }
}

/* ---------- аудит: карта замечаний (перенесена из прежнего интерфейса) ---------- */
var CATEGORY_LABELS = { layout: 'Верстка', template: 'Шаблон', density: 'Плотность', integrity: 'Целостность', content: 'Контент' };
async function loadAuditMap() {
  var body = $('auditBody'); if (!body) return;
  try { S.auditMap = await api('/jobs/' + S.jobId + '/variants/' + S.design + '/audit-map'); }
  catch (e) { body.innerHTML = '<p class="err">Не удалось загрузить карту аудита: ' + t(e.message) + '</p>'; return; }
  renderAuditMap();
}
function shapeIssues(slide, shape) {
  return slide.issues.filter(function (i) { return i.shape_id !== null && i.shape_id !== undefined && i.shape_id === shape.shape_id; });
}
function shapeLabel(sh) {
  var n = sh.name || '';
  if (/^title/i.test(n)) return 'Заголовок'; if (/^subtitle/i.test(n)) return 'Подзаголовок'; if (/^footer|^date|^slide number/i.test(n)) return 'Колонтитул';
  if (sh.kind === 'picture') return 'Изображение'; if (sh.kind === 'chart') return 'Диаграмма'; if (sh.kind === 'table') return 'Таблица'; return 'Текстовый блок';
}
function renderAuditMap() {
  var body = $('auditBody'); var map = S.auditMap; if (!body || !map) return;
  var all = []; map.slides.forEach(function (s) { all = all.concat(s.issues); }); all = all.concat(map.deck_issues || []);
  var fixable = all.filter(function (i) { return i.fixable; }); var fixed = (map.fixed_issue_ids || []).length;
  var checked = map.slides.length; var manual = all.length - fixable.length;
  var summary = '<div class="audit-summary"><span class="chip">Проверено ' + checked + ' ' + plural(checked, ['слайд', 'слайда', 'слайдов']) + '</span>';
  if (!all.length) summary += '<span class="chip ok">Замечаний нет: верстка и визуальная проверка каждого слайда пройдены</span>';
  else {
    summary += '<span class="chip bad">' + all.length + ' ' + plural(all.length, ['замечание', 'замечания', 'замечаний']) + '</span>';
    if (fixable.length) summary += '<span class="chip">' + fixable.length + ' можно исправить автоматически</span><button class="btn small danger" id="fixAllBtn"' + (S.auditBusy ? ' disabled' : '') + '>Исправить все</button>';
    if (manual > 0) summary += '<span class="chip">' + manual + ' ' + plural(manual, ['требует', 'требуют', 'требуют']) + ' ручной правки или пересборки слайда</span>';
  }
  if (fixed) summary += '<span class="chip fixed">Исправлено: ' + fixed + '</span>';
  summary += '</div><div class="audit-legend"><span class="l-bad"><i></i>есть замечание — кликните, чтобы посмотреть и исправить</span><span class="l-ok"><i></i>проверки пройдены</span></div>';
  var cards = map.slides.map(function (sl) {
    var slideLevel = sl.issues.filter(function (i) { return i.shape_id === null || i.shape_id === undefined; });
    var bad = sl.issues.length > 0;
    var frames = sl.shapes.map(function (sh) {
      var iss = shapeIssues(sl, sh);
      var st = 'left:' + (sh.bbox.x * 100).toFixed(2) + '%;top:' + (sh.bbox.y * 100).toFixed(2) + '%;width:' + (sh.bbox.width * 100).toFixed(2) + '%;height:' + (sh.bbox.height * 100).toFixed(2) + '%;';
      return '<div class="audit-frame ' + (iss.length ? 'bad' : 'ok') + '" style="' + st + '" data-shape="' + sh.shape_id + '" title="' + raw(shapeLabel(sh)) + '">' + (iss.length ? '<span class="count">' + iss.length + '</span>' : '') + '</div>';
    }).join('');
    var slideIssues = slideLevel.length ? '<div class="audit-slide-issues">' + slideLevel.map(function (i) {
      return '<div class="audit-issue"><span class="badge ' + i.category + '">' + (CATEGORY_LABELS[i.category] || i.category) + '</span><div><div class="desc">' + raw(i.description) + '</div><div class="hint">Замечание ко всему слайду — исправляется вручную или пересборкой</div></div></div>';
    }).join('') + '</div>' : '';
    return '<div class="audit-card' + (bad ? ' has-issues' : '') + '" data-slide="' + sl.slide_id + '"><div class="head"><span><b>Слайд ' + sl.number + '</b></span><span class="state ' + (bad ? 'bad' : 'ok') + '">' + (bad ? sl.issues.length + ' ' + plural(sl.issues.length, ['замечание', 'замечания', 'замечаний']) : 'Без замечаний') + '</span></div>' +
      '<div class="audit-slide"><img src="' + API + '/jobs/' + S.jobId + '/variants/' + S.design + '/preview/' + sl.number + '?v=' + S.stamp + '" alt="Слайд ' + sl.number + '" loading="lazy">' + frames + '</div>' + slideIssues + '</div>';
  }).join('');
  var deck = (map.deck_issues || []).length ? '<div style="margin-top:16px">' + map.deck_issues.map(function (i) {
    return '<div class="audit-issue"><span class="badge ' + i.category + '">' + (CATEGORY_LABELS[i.category] || i.category) + '</span><div><div class="desc">' + raw(i.description) + '</div></div></div>';
  }).join('') + '</div>' : '';
  body.innerHTML = summary + '<div class="audit-grid">' + cards + '</div>' + deck;
  var fa = $('fixAllBtn'); if (fa) fa.addEventListener('click', function () { runFix(fixable.map(function (i) { return i.issue_id; }), null); });
  body.querySelectorAll('.audit-frame.bad').forEach(function (fr) { fr.addEventListener('click', function (ev) { ev.stopPropagation(); openIssuePopover(fr); }); });
}
function closeIssuePopover() {
  document.querySelectorAll('.audit-pop').forEach(function (p) { p.remove(); });
  document.querySelectorAll('.audit-frame.active').forEach(function (f) { f.classList.remove('active'); });
}
function openIssuePopover(frame) {
  closeIssuePopover();
  var card = frame.closest('.audit-card'); var slideId = card.getAttribute('data-slide'); var shapeId = parseInt(frame.getAttribute('data-shape'), 10);
  var slide = S.auditMap.slides.filter(function (s) { return s.slide_id === slideId; })[0];
  var shape = slide.shapes.filter(function (s) { return s.shape_id === shapeId; })[0];
  var issues = shapeIssues(slide, shape); frame.classList.add('active');
  var ids = issues.filter(function (i) { return i.fixable; }).map(function (i) { return i.issue_id; });
  var pop = document.createElement('div'); pop.className = 'audit-pop';
  pop.innerHTML = '<div class="title"><span>' + shapeLabel(shape) + '</span><button type="button" aria-label="Закрыть">×</button></div>' + issues.map(function (i) {
    return '<div class="audit-issue"><span class="badge ' + i.category + '">' + (CATEGORY_LABELS[i.category] || i.category) + '</span><div><div class="desc">' + raw(i.description) + '</div>' + (i.fixable ? '' : '<div class="slide">Автоисправления нет — правится вручную в PowerPoint</div>') + '</div></div>';
  }).join('') + '<div class="actions">' + (ids.length ? '<button class="btn small primary fix-btn">Исправить' + (ids.length > 1 ? ' (' + ids.length + ')' : '') + '</button>' : '<span class="manual">Для этих замечаний автоисправления нет</span>') + '</div>';
  pop.addEventListener('click', function (ev) { ev.stopPropagation(); });
  pop.querySelector('.title button').addEventListener('click', closeIssuePopover);
  var fb = pop.querySelector('.fix-btn');
  if (fb) fb.addEventListener('click', function () { fb.disabled = true; fb.textContent = 'Исправляем…'; runFix(ids, slideId); });
  var el = card.querySelector('.audit-slide'); el.appendChild(pop);
  var fTop = parseFloat(frame.style.top), fH = parseFloat(frame.style.height), fLeft = parseFloat(frame.style.left);
  var pw = pop.offsetWidth, sw = el.clientWidth, sh = el.clientHeight;
  var left = Math.min(Math.max(0, fLeft / 100 * sw), Math.max(0, sw - pw - 6)); var top = (fTop + fH) / 100 * sh + 6;
  if (top + pop.offsetHeight > sh) top = Math.max(4, fTop / 100 * sh - pop.offsetHeight - 6);
  pop.style.left = left + 'px'; pop.style.top = top + 'px';
}
async function runFix(ids, slideId) {
  if (S.auditBusy || !ids.length) return; S.auditBusy = true; closeIssuePopover();
  var targets = slideId ? document.querySelectorAll('.audit-card[data-slide="' + slideId + '"] .audit-slide') : document.querySelectorAll('.audit-slide');
  targets.forEach(function (e) { e.classList.add('busy'); });
  var fa = $('fixAllBtn'); if (fa) fa.disabled = true;
  try {
    var res = await postJson('/jobs/' + S.jobId + '/fix-issues', { issue_ids: ids });
    S.stamp = Date.now(); notify(res.applied && res.applied.length ? 'Исправлено: ' + res.applied.length + '. Аудит выполнен повторно.' : 'Исправления не потребовались.');
  } catch (e) { notify('Не удалось исправить: ' + e.message, true); }
  S.auditBusy = false; try { var pr2 = await api('/jobs/' + S.jobId + '/variants/' + S.design + '/preview'); S.previews[S.design] = pr2.slides || []; } catch (e) { /* ignore */ } await loadAuditMap();
}
async function runAudit() {
  S.audit = 'running'; S.runError = ''; S.runKind = 'audit'; var run = ++S.runId; render(); updateSide();
  try {
    await postJson('/jobs/' + S.jobId + '/audit');
    var job = await pollJob(S.jobId, run, ['done']);
    if (!job) return; S.job = job; S.audit = 'done'; S.stamp = Date.now();
  } catch (e) { S.audit = 'choice'; S.runError = e.message; }
  render();
}

/* ---------- события ---------- */
document.addEventListener('click', function (e) {
  var el = e.target.closest('button,[data-go],[data-design],[data-cmp-to],a.deck-download');
  if (!el) return;
  var d = el.dataset;
  if (el.matches('a.deck-download')) { e.stopPropagation(); return; }
  if (d.go !== undefined) { if (+d.go === 5 && S.job) S.plan = S.job.selected_plan || S.plan; go(d.go); return; }
  if (d.panel) { if (d.panel === 'design-system' && !S.tpl.id) return; openPanel(d.panel); return; }
  if (d.plan) { remember(); S.plan = d.plan; render(); return; }
  if (d.design) {
    S.design = d.design;
    document.querySelectorAll('[data-design]').forEach(function (c) { var on = c.dataset.design === S.design; c.classList.toggle('selected', on); c.setAttribute('aria-pressed', on); });
    var inf = vInfo(S.design); var sa = document.querySelector('.sticky-action'); if (sa) { sa.querySelector('.small').textContent = 'Выбран дизайн ' + VL[S.design] + ' · ' + inf.label; var nb = sa.querySelector('.primary'); if (nb) nb.disabled = false; }
    return;
  }
  if (d.mode) { S.designMode = d.mode; render(); return; }
  if (d.cmp !== undefined) { S.cmp += +d.cmp; $('compare-area').innerHTML = compareView(); return; }
  if (d.cmpTo !== undefined) { S.cmp = +d.cmpTo; $('compare-area').innerHTML = compareView(); return; }
  if (d.slide !== undefined) { if (S.refineBusy) return; S.selectedSlide = +d.slide; render(); return; }
  if (d.audience) { $('audience').value = d.audience; return; }
  if (d.goal) { $('goal').value = d.goal + '.'; return; }
  if (d.clarify !== undefined) { answerClarification(S.clarify.options[+d.clarify]); return; }
  if (d.volume) { var mm = d.volume.split(','); $('volume-min').value = mm[0]; $('volume-max').value = mm[1]; document.querySelectorAll('[data-volume]').forEach(function (c) { c.classList.toggle('selected', c === el); }); return; }
  if (d.remove !== undefined) { S.materials.splice(+d.remove, 1); markChanged(); refreshMaterials(); return; }
  if (d.editSlide !== undefined) { openPanel('slide-edit', +d.editSlide); return; }
  if (d.regenSlide !== undefined) { regenPlanSlide(+d.regenSlide, ''); return; }
  if (d.format) { S.format = d.format; render(); return; }
  var a = d.action; if (!a) return;
  actions[a] && actions[a](el);
});
document.addEventListener('click', function (e) {
  if (e.target.classList.contains('backdrop')) closePanel();
  if (!e.target.closest('.audit-pop')) closeIssuePopover();
});
document.addEventListener('keydown', function (e) {
  if (e.key === 'Escape') closePanel();
  if (!document.querySelector('.drawer') && S.screen === 6 && S.designMode === 'compare' && (e.key === 'ArrowLeft' || e.key === 'ArrowRight') && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) {
    S.cmp += e.key === 'ArrowRight' ? 1 : -1; $('compare-area').innerHTML = compareView();
  }
  if ((e.key === 'Enter' || e.key === ' ') && e.target.matches && e.target.matches('[data-design][role=button]')) { e.preventDefault(); e.target.click(); }
});

async function answerClarification(text) {
  if (!text || !text.trim()) return;
  S.clarifyAnswer = text.trim(); closePanel(); remember();
  var ok = await requestIdea({ clarification: S.clarifyAnswer });
  S.ideaHistory = JSON.stringify([S.topic, S.audience, S.goal, materialsText().length]);
  if (ok && !S.clarify) { markChanged(); go(3); } else { render(); }
}
async function regenPlanSlide(i, instruction) {
  var key = S.plan; S.busySlides[key + i] = true; closePanel(); render();
  try {
    var r = await postJson('/jobs/' + S.jobId + '/plans/' + key + '/slides/' + i + '/regenerate', { instruction: instruction });
    S.job.plan_variants[key].plan.slides[i] = r.slide; S.job.designs_stale = r.designs_stale;
  } catch (e) { notify('Не удалось перегенерировать слайд: ' + e.message, true); }
  delete S.busySlides[key + i]; render();
}

var actions = {
  home: function () { go(0); },
  'choose-template': chooseTemplate,
  'start-next': function () { if (S.tpl.name && S.tpl.status !== 'error') go(1); },
  'about-next': function () { remember(); if (!S.topic.trim() && !okMaterials().length) { notify('Добавьте тему или описание презентации.', true); return; } if (S.matBusy) { notify('Дождитесь, пока материалы будут прочитаны.', true); return; } go(2); },
  'audience-next': async function () {
    remember();
    if (!S.audience.trim() || !S.goal.trim()) { notify('Укажите аудиторию и ожидаемое решение.', true); return; }
    var changedInput = !S.idea || S.ideaHistory !== JSON.stringify([S.topic, S.audience, S.goal, materialsText().length]);
    if (S.clarify && !changedInput) { notify('Выберите один из ответов в карточке уточнения.', true); return; }
    if (!changedInput && S.idea) { go(3); return; }
    S.clarify = null; S.clarifyAnswer = '';
    var btn = document.querySelector('[data-action=audience-next]'); if (btn) { btn.disabled = true; btn.firstChild.textContent = 'Определяем главную идею…'; }
    var ok = await requestIdea();
    S.ideaHistory = JSON.stringify([S.topic, S.audience, S.goal, materialsText().length]);
    updateSide();
    if (ok && !S.clarify) { markChanged(); go(3); } else { if (S.screen === 2) { var fa = $('form-area'); fa.innerHTML = audienceView(); } }
  },
  'generate-plan': function () {
    remember();
    if (S.min < 3 || S.max < S.min || S.max > 40) { notify('Проверьте диапазон: минимум - от 3, максимум не меньше минимума и не больше 40.', true); return; }
    startPlanning();
  },
  'retry-run': function () { if (S.runKind === 'design') selectPlan(); else startPlanning(); },
  'regenerate-plan': async function () {
    if (S.planBusy) return; var key = S.plan; S.planBusy = true; render();
    try { var r = await postJson('/jobs/' + S.jobId + '/plans/' + key + '/regenerate'); S.job.plan_variants[key] = r.variant; S.job.designs_stale = r.designs_stale; }
    catch (e) { notify('Не удалось построить другой вариант: ' + e.message, true); }
    S.planBusy = false; render();
  },
  'plan-next': function () { if (S.stale) { openPanel('rebuild'); return; } selectPlan(); },
  'design-next': async function (el) {
    if (!S.design) return; el.disabled = true; el.firstChild.textContent = 'Сохраняем выбор…';
    try { await postJson('/jobs/' + S.jobId + '/select-variant', { variant: S.design }); S.job = await api('/jobs/' + S.jobId); S.selectedSlide = 0; S.audit = 'choice'; go(7); }
    catch (e) { notify('Не удалось сохранить выбор: ' + e.message, true); render(); }
  },
  rewrite: function () { runSlideEdit('rewrite'); },
  recompose: function () { runSlideEdit('recompose'); },
  'edit-current': function () { openPanel('deck-slide-edit', S.selectedSlide); },
  'save-deck-slide': function () {
    var title = $('deck-title-edit').value, text = $('deck-body-edit').value;
    if (!title.trim() || !text.trim()) { notify('Заполните заголовок и текст.', true); return; }
    closePanel(); runSlideEdit('manual', { title: title, text: text });
  },
  'deck-instruction': function () { var ins = $('deck-instruction').value.trim(); if (!ins) { notify('Напишите, что нужно изменить.', true); return; } closePanel(); runSlideEdit('rewrite', { instruction: ins }); },
  'refine-next': function () { go(8); },
  'run-audit': runAudit,
  'skip-audit': async function () {
    try { await postJson('/jobs/' + S.jobId + '/skip-audit'); S.job = await api('/jobs/' + S.jobId); S.audit = 'skipped'; go(9); }
    catch (e) { notify('Не удалось пропустить аудит: ' + e.message, true); }
  },
  'export-next': function () { go(9); },
  download: function () {
    var a = document.createElement('a'); a.href = API + '/jobs/' + S.jobId + '/export?format=' + S.format + '&variant=' + S.design; a.download = '';
    document.body.appendChild(a); a.click(); a.remove();
    var fb = $('export-feedback'); if (fb) fb.textContent = 'Скачивание .' + S.format + ' началось. PDF и HTML могут готовиться несколько секунд.';
  },
  close: closePanel,
  'replace-template': function () { closePanel(); chooseTemplate(); },
  'add-file': function () {
    var input = document.createElement('input'); input.type = 'file'; input.multiple = true; input.accept = '.txt,.md,.csv,.json,.pdf,.docx,.xlsx';
    input.onchange = function () { addFiles(Array.prototype.slice.call(input.files)); }; input.click();
  },
  'save-link': function () {
    var v = $('new-link').value.trim();
    if (!/^https?:\/\/\S+/.test(v)) { notify('Укажите адрес, начинающийся с https:// или http://', true); return; }
    closePanel(); addLink(v);
  },
  'save-materials': function () { S.topic = $('panel-topic').value; closePanel(); markChanged(); render(); },
  'save-meaning': function () {
    var v = $('panel-idea').value; if (!v.trim()) { notify('Главная идея не должна быть пустой.', true); return; }
    S.idea = deyo(v); S.audience = $('panel-audience').value; S.goal = $('panel-goal').value; closePanel(); markChanged(); render();
  },
  'regenerate-idea': async function (el) {
    var prev = $('panel-idea').value; el.disabled = true;
    S.audience = $('panel-audience').value; S.goal = $('panel-goal').value;
    try {
      var r = await postJson('/briefing/idea', ideaPayload({ previous_idea: prev, clarification: S.clarifyAnswer }));
      if ($('panel-idea')) $('panel-idea').value = deyo(r.idea || prev);
    } catch (e) { notify('Не удалось предложить формулировку: ' + e.message, true); }
    el.disabled = false;
  },
  'save-params': function () {
    var a = +$('volume-min').value, b = +$('volume-max').value;
    if (a < 3 || b < a || b > 40) { notify('Проверьте диапазон: минимум - от 3, максимум не меньше минимума и не больше 40.', true); return; }
    remember(); closePanel(); markChanged(); render();
  },
  'save-slide': async function () {
    var title = $('slide-title-edit').value, text = $('slide-body-edit').value, why = $('slide-purpose-edit').value, i = S.editIndex, key = S.plan;
    if (!title.trim() || !text.trim()) { notify('Заполните заголовок и содержание.', true); return; }
    try {
      var r = await api('/jobs/' + S.jobId + '/plans/' + key + '/slides/' + i, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ title: title, text: text, why: why }) });
      S.job.plan_variants[key].plan.slides[i] = r.slide; S.job.designs_stale = r.designs_stale;
      closePanel(); render();
    } catch (e) { notify('Не удалось сохранить: ' + e.message, true); }
  },
  'apply-slide-instruction': function () {
    var ins = $('slide-instruction').value.trim(); if (!ins) { notify('Напишите, что нужно изменить в слайде.', true); return; }
    regenPlanSlide(S.editIndex, ins);
  },
  'save-clarification': function () { answerClarification($('clarification-text').value); },
  'refresh-dependent': function () { openPanel('rebuild'); },
  'confirm-rebuild': function () { closePanel(); startPlanning(); }
};

/* ---------- старт ---------- */
render();
})();
