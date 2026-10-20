// Run from any directory: node backend/tests/test_docx_frontend.cjs
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../../frontend/app.js'), 'utf8');
function section(start, end) { return source.slice(source.indexOf(start), source.indexOf(end, source.indexOf(start))); }
const context = { S: { materials: [] } };
vm.createContext(context);
vm.runInContext(section('function raw(', '/* неразрывные'), context);
vm.runInContext(section('function okMaterials(', 'async function addFiles('), context);
vm.runInContext(section('function applyExtract(', 'function fileRows('), context);
vm.runInContext(section('function slideBody(', 'function slideText('), context);
vm.runInContext('function t(value) { return raw(value); }', context);
const data = { headers: ['Регион', 'Выручка'], rows: [['Орёл', '1 234,50'], ['<script>alert(1)</script>', '0']] };
const html = context.slideBody({ source_table_id: 'source.table_1', content_blocks: [{ table: data }] });
assert.ok(html.includes('Орёл'));
assert.ok(html.includes('1 234,50'));
assert.ok(html.includes('<td>0</td>'));
assert.ok(!html.includes('<script>'));
assert.ok(html.includes('&lt;script&gt;'));
const record = {};
context.applyExtract(record, { status: 'ok', material_id: 'server_id', text: 'summary', tables: [{ data }], chars: 7 });
context.S.materials.push(record, { status: 'error', materialId: 'failed_id' });
assert.equal(JSON.stringify(context.materialIds()), '["server_id"]');
assert.equal(record.tables[0].data.rows[0][0], 'Орёл');
console.log('DOCX UI: exact cells, escaping, server IDs and error filtering passed');
