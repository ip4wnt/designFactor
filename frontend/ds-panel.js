/* Панель дизайн-системы: перенесена из прежнего интерфейса, стили изолированы классом .legacy */
function escapeHtml(s){
  return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; });
}
function dsSectionTitleHtml(label, extra){
  return '<div class="ds-section-title"><span>' + label + (extra ? ' ' + extra : '') + '</span></div>';
}
function dsAddButtonHtml(title){ return ''; }
function pluralRu(n, one, few, many){
  var m10 = n % 10, m100 = n % 100;
  if(m10 === 1 && m100 !== 11) return one;
  if(m10 >= 2 && m10 <= 4 && (m100 < 10 || m100 >= 20)) return few;
  return many;
}
function formatPt(v){ return (Math.round(v * 10) / 10).toString().replace(/\.0$/, '') + 'pt'; }
var COMPOSITION_RU = {
  single_focus: 'один фокус', title_body: 'заголовок + текст', two_columns: 'две колонки', vertical_sequence: 'вертикальный ряд',
  horizontal_sequence: 'горизонтальный ряд', grid: 'сетка', image_text: 'фото + текст', full_bleed_image: 'фото на весь слайд',
  metric_focus: 'фактоид', quote: 'цитата', comparison: 'сравнение', list: 'список', timeline: 'таймлайн'
};
function compositionRu(c){ return COMPOSITION_RU[c] || (c || '').replace(/_/g, ' '); }
var SOURCE_GROUP_RU = { sample: 'Слайды-образцы', layout: 'Макеты', library: 'Библиотека компоновок' };
function renderDesignSystem(ds, dsPanelBody){
  dsPanelBody.innerHTML = '';

  // Источник дизайна: откуда генератор берет композиции и что он распознал
  if(ds.design_source){
    var srcSection = document.createElement('div');
    srcSection.className = 'ds-section';
    var info = ds.design_source_info || {};
    srcSection.innerHTML = dsSectionTitleHtml('Откуда берется дизайн');
    var stats = document.createElement('div');
    stats.className = 'ds-source';
    stats.innerHTML =
      '<div class="stat main"><div class="n">' + escapeHtml(info.label || ds.design_source) + '</div><div class="l">источник композиций</div></div>' +
      (info.slides_total != null ? '<div class="stat"><div class="n">' + info.sample_slides + ' / ' + info.slides_total + '</div><div class="l">слайдов-образцов распознано</div></div>' : '') +
      (info.layouts_total != null ? '<div class="stat"><div class="n">' + info.layouts_with_body + ' / ' + info.layouts_total + '</div><div class="l">макетов с телом слайда</div></div>' : '');
    srcSection.appendChild(stats);
    var comps = ds.compositions || [];
    ['sample','layout','library'].forEach(function(src){
      var rows = comps.filter(function(c){ return c.source === src; });
      if(!rows.length) return;
      var g = document.createElement('div');
      g.className = 'ds-comp-group';
      var total = rows.reduce(function(a,c){ return a + c.count; }, 0);
      g.innerHTML = '<div class="g">' + SOURCE_GROUP_RU[src] + ' · ' + total + '</div><div class="chips">' +
        rows.map(function(c){ return '<span title="' + escapeHtml((c.examples||[]).join(', ')) + '">' + escapeHtml(c.label) + (c.count > 1 ? '<b>×' + c.count + '</b>' : '') + '</span>'; }).join('') + '</div>';
      srcSection.appendChild(g);
    });
    var note = document.createElement('p');
    note.style.cssText = 'font-size:12px;color:var(--text-2);margin:12px 0 0;line-height:1.5;';
    note.textContent = ds.design_source === 'library'
      ? 'В шаблоне нет ни слайдов-образцов, ни макетов с телом — компоновки берутся из собственной библиотеки и раскрашиваются шрифтами и цветами шаблона.'
      : 'Слайды-образцы клонируются и подстраиваются под число тезисов; недостающие виды (таблица, график, крупный тезис) достраиваются из библиотеки на чистом макете шаблона.';
    srcSection.appendChild(note);
    dsPanelBody.appendChild(srcSection);
  }

  // Colors
  var colorSection = document.createElement('div');
  colorSection.className = 'ds-section';
  colorSection.innerHTML = dsSectionTitleHtml('Цвета');
  var swatchWrap = document.createElement('div');
  swatchWrap.className = 'ds-swatches';
  (ds.colors || []).forEach(function(c){
    var sw = document.createElement('div');
    sw.className = 'ds-swatch';
    sw.style.background = c.hex;
    sw.setAttribute('data-label', c.label + ' — ' + c.hex);
    swatchWrap.appendChild(sw);
  });
  var addColorBtn = document.createElement('span');
  addColorBtn.className = 'ds-add-btn';
  addColorBtn.title = 'Добавить цвет (скоро)';
  addColorBtn.textContent = '+';
  swatchWrap.appendChild(addColorBtn);
  colorSection.appendChild(swatchWrap);
  dsPanelBody.appendChild(colorSection);

  // Typography — полная шкала шаблона по семантическим ролям (из Template JSON движка).
  var fontSection = document.createElement('div');
  fontSection.className = 'ds-section';
  var scale = ds.typography_scale || [];
  var theme = ds.theme_fonts || {};
  var families = (theme.families && theme.families.length) ? theme.families : [];
  fontSection.innerHTML = dsSectionTitleHtml('Типографика', scale.length ? '(' + scale.length + ' стил' + pluralRu(scale.length, 'ь', 'я', 'ей') + ')' : '');
  if(families.length){
    var famRow = document.createElement('div');
    famRow.className = 'ds-font-row';
    famRow.innerHTML = '<span class="role">Гарнитуры</span><span class="val">' + escapeHtml(families.join(' · ')) + '</span>';
    fontSection.appendChild(famRow);
  }
  if(theme.major || theme.minor){
    var themeRow = document.createElement('div');
    themeRow.className = 'ds-font-row';
    themeRow.innerHTML = '<span class="role">Тема (заголовки / текст)</span><span class="val">' + escapeHtml((theme.major || '—') + ' / ' + (theme.minor || '—')) + '</span>';
    fontSection.appendChild(themeRow);
  }
  if(scale.length){
    // Живой образец шкалы: несколько ключевых кеглей в пропорции.
    var sample = document.createElement('div');
    sample.className = 'ds-type-sample';
    var sampleSizes = [];
    scale.forEach(function(r){ (r.sizes_pt || []).forEach(function(sz){ if(sampleSizes.indexOf(sz) < 0) sampleSizes.push(sz); }); });
    sampleSizes.sort(function(a,b){ return b - a; });
    sampleSizes.slice(0, 5).forEach(function(sz){
      var sp = document.createElement('span');
      sp.style.fontSize = Math.max(9, Math.min(34, sz * 0.5)) + 'px';
      sp.style.fontFamily = "'" + (families[0] || 'inherit') + "', inherit";
      sp.textContent = 'Аа ' + sz + 'pt';
      sample.appendChild(sp);
    });
    fontSection.appendChild(sample);

    var table = document.createElement('table');
    table.className = 'ds-type-table';
    table.innerHTML = '<thead><tr><th>Роль текста</th><th>Кегль</th><th class="n">Макетов</th></tr></thead>';
    var tbody = document.createElement('tbody');
    scale.forEach(function(r){
      var tr = document.createElement('tr');
      var dots = (r.colors || []).slice(0, 4).map(function(c){ return '<i style="background:' + escapeHtml(c) + '"></i>'; }).join('');
      tr.innerHTML =
        '<td><div class="role">' + escapeHtml(r.label) + '</div><div class="fam">' + escapeHtml(r.family) + (r.bold ? ', жирный' : '') + (dots ? '<span class="ds-color-dots">' + dots + '</span>' : '') + '</div></td>' +
        '<td class="sz">' + (r.sizes_pt || []).map(function(v){ return formatPt(v); }).join(' / ') + (r.estimated ? '<div class="est">оценка по слоту</div>' : '') + '</td>' +
        '<td class="n" title="' + escapeHtml((r.layouts || []).join(', ')) + '">' + r.layouts_count + '</td>';
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    fontSection.appendChild(table);
  } else {
    (ds.fonts || []).forEach(function(f){
      var row = document.createElement('div');
      row.className = 'ds-font-row';
      var fontName = f.resolved_font_label || f.raw_font;
      row.innerHTML = '<span class="role">' + escapeHtml(f.label) + '</span><span class="val">' + escapeHtml(fontName) + ', ' + f.size_pt + 'pt' + (f.bold ? ', bold' : '') + '</span>';
      fontSection.appendChild(row);
    });
  }
  dsPanelBody.appendChild(fontSection);

  // Layouts — все макеты шаблона (slideLayout) с геометрией слотов, сгруппированные по роли в нарративе.
  var layoutSection = document.createElement('div');
  layoutSection.className = 'ds-section';
  var tLayouts = ds.template_layouts || [];
  if(tLayouts.length){
    layoutSection.innerHTML = dsSectionTitleHtml('Макеты шаблона', '(' + tLayouts.length + ')');
    var groupsOrder = ['Обложка', 'Раздел', 'Контент', 'Финал'];
    var groups = {};
    tLayouts.forEach(function(l){
      var g = l.narrative_label || 'Контент';
      (groups[g] = groups[g] || []).push(l);
    });
    Object.keys(groups).forEach(function(g){ if(groupsOrder.indexOf(g) < 0) groupsOrder.push(g); });
    groupsOrder.forEach(function(g){
      if(!groups[g]) return;
      var title = document.createElement('div');
      title.className = 'ds-layout-group';
      title.textContent = g + ' · ' + groups[g].length;
      layoutSection.appendChild(title);
      var grid = document.createElement('div');
      grid.className = 'ds-layouts-grid';
      groups[g].forEach(function(l){
        var card = document.createElement('div');
        card.className = 'ds-layout-card';
        var mock = document.createElement('div');
        mock.className = 'ds-layout-mock';
        if(l.background_hex) mock.style.background = l.background_hex;
        (l.elements || []).forEach(function(el){
          var box = document.createElement('div');
          var cls = 'mock-el';
          if(el.kind === 'image') cls += ' is-image';
          else if(el.kind === 'decor_image' || el.kind === 'decor' || el.role === 'decorative') cls += ' is-decor';
          else if(/title/.test(el.role || '')) cls += ' is-title';
          box.className = cls;
          box.style.left = (el.left_pct || 0) + '%';
          box.style.top = (el.top_pct || 0) + '%';
          box.style.width = Math.max(el.width_pct || 0, 2) + '%';
          box.style.height = Math.max(el.height_pct || 0, 3) + '%';
          if(el.kind === 'text' && el.color && el.role !== 'decorative') box.style.background = el.color;
          box.title = (el.role || el.kind) + (el.font_size_pt ? ' · ' + formatPt(el.font_size_pt) : '');
          mock.appendChild(box);
        });
        card.appendChild(mock);
        var nameEl = document.createElement('div');
        nameEl.className = 'ds-layout-name';
        nameEl.textContent = l.name;
        nameEl.title = l.name + (l.best_for && l.best_for.length ? ' — ' + l.best_for.join(', ') : '');
        card.appendChild(nameEl);
        var meta = document.createElement('div');
        meta.className = 'ds-layout-meta';
        var bits = [];
        if(l.composition) bits.push(compositionRu(l.composition));
        bits.push(l.text_slots + ' текст.');
        if(l.image_slots) bits.push(l.image_slots + ' фото');
        meta.textContent = bits.join(' · ');
        meta.title = meta.textContent;
        card.appendChild(meta);
        grid.appendChild(card);
      });
      layoutSection.appendChild(grid);
    });
  } else {
    layoutSection.innerHTML = dsSectionTitleHtml('Макеты', '(' + (ds.layouts||[]).length + ')');
    var layoutsGrid = document.createElement('div');
    layoutsGrid.className = 'ds-layouts-grid';
    (ds.layouts || []).forEach(function(l){
      var card = document.createElement('div');
      card.className = 'ds-layout-card';
      var mock = document.createElement('div');
      mock.className = 'ds-layout-mock';
      var example = (l.examples || [])[0];
      if(example && example.elements && example.elements.length){
        example.elements.forEach(function(el){
          var box = document.createElement('div');
          box.className = 'mock-el' + (el.is_title_role ? ' is-title' : '');
          box.style.left = (el.left_pct || 0) + '%';
          box.style.top = (el.top_pct || 0) + '%';
          box.style.width = Math.max(el.width_pct || 0, 3) + '%';
          box.style.height = Math.max(el.height_pct || 0, 4) + '%';
          if(el.fill_hex) box.style.background = '#' + el.fill_hex.replace('#','');
          mock.appendChild(box);
        });
      } else {
        mock.innerHTML = '<div style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:var(--text-2);font-size:10px;">нет превью</div>';
      }
      card.appendChild(mock);
      var nameEl = document.createElement('div');
      nameEl.className = 'ds-layout-name';
      nameEl.textContent = l.name || l.category;
      nameEl.title = l.name || l.category;
      card.appendChild(nameEl);
      layoutsGrid.appendChild(card);
    });
    layoutSection.appendChild(layoutsGrid);
  }
  dsPanelBody.appendChild(layoutSection);

  // Table / chart style
  var styleSection = document.createElement('div');
  styleSection.className = 'ds-section';
  styleSection.innerHTML =
    dsSectionTitleHtml('Стили таблиц и графиков') +
    '<div class="ds-font-row"><span class="role">Таблицы</span><span class="val">' + (ds.table_style && ds.table_style.found ? 'найден стиль в шаблоне' : 'по умолчанию из палитры') + '</span></div>' +
    '<div class="ds-font-row"><span class="role">Графики</span><span class="val">' + (ds.chart_style && ds.chart_style.found ? 'найден стиль в шаблоне' : 'по умолчанию из палитры') + '</span></div>';
  dsPanelBody.appendChild(styleSection);

  // Raw JSON toggle
  var jsonToggle = document.createElement('div');
  jsonToggle.className = 'ds-json-toggle';
  jsonToggle.textContent = 'Показать как JSON (манифест сборки)';
  var jsonBlock = document.createElement('div');
  jsonBlock.className = 'ds-json-block';
  jsonBlock.textContent = JSON.stringify(ds.manifest, null, 2);
  jsonToggle.addEventListener('click', function(){
    jsonBlock.classList.toggle('open');
    jsonToggle.textContent = jsonBlock.classList.contains('open') ? 'Скрыть JSON' : 'Показать как JSON (манифест сборки)';
  });
  dsPanelBody.appendChild(jsonToggle);
  dsPanelBody.appendChild(jsonBlock);
}
