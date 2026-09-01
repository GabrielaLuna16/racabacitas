import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import XLSX from 'xlsx';

const source = process.argv[2];
const target = process.argv[3] || 'Dashboard_Citas_Racaba_2026.html';
if (!source) throw new Error('Uso: node update_dashboard_from_excel.mjs <archivo.xlsx> [dashboard.html]');

const text = (value, fallback = '') => String(value ?? '').trim() || fallback;
const cleanGroup = (value, fallback = '-') => text(value, fallback).replace(/\s*\(\s*\d+\s*\)\s*$/, '') || fallback;
const excelValue = value => value instanceof Date ? value.toISOString() : (value ?? '');

const workbook = XLSX.readFile(source, { cellDates: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
const headerIndex = rows.findIndex(row => row.includes('Full Name') && row.includes('Created Time'));
if (headerIndex < 0) throw new Error('No se encontraron las columnas Full Name y Created Time.');

const headers = rows[headerIndex].map(value => text(value));
let currentProject = '-';
let currentHost = '-';
const records = [];

for (const values of rows.slice(headerIndex + 1)) {
  const raw = Object.fromEntries(headers.map((header, index) => [header, values[index]]));
  const contacto = text(raw['Full Name']);
  if (!contacto) continue;

  const meetingDate = raw['From'];
  let proyecto = '-';
  let asesor = '-';
  if (meetingDate !== '' && meetingDate != null) {
    if (raw['Proyecto (Meetings)']) currentProject = cleanGroup(raw['Proyecto (Meetings)']);
    if (raw['Host (Meetings)']) currentHost = cleanGroup(raw['Host (Meetings)']);
    proyecto = currentProject;
    asesor = currentHost;
  }

  records.push({
    id: records.length + 1,
    proyecto,
    asesor,
    contacto,
    campana: text(raw['Campaign'], 'Sin campaña'),
    fecha: excelValue(meetingDate),
    nota: text(raw['Note Content']),
    contactStatus: text(raw['Contact Status'], 'Sin estatus'),
    citaStatus: text(raw['Estatus de cita (Meetings)'], 'Sin estatus'),
    anuncio: text(raw['Anuncio MKT'], 'Sin anuncio'),
    fechaApartado: excelValue(raw['Fecha Apartado']),
    fechaContrato: excelValue(raw['Fecha Contrato']),
    recordId: text(raw['Record Id']).replace(/^zcrm_/i, ''),
    fechaCreacion: excelValue(raw['Created Time'])
  });
}

if (!records.length) throw new Error('El archivo no contiene registros utilizables.');
const html = fs.readFileSync(target, 'utf8');
const payload = JSON.stringify(records).replaceAll('</', '<\\/');
const matches = html.match(/const DATA=\[.*?\];/gs) || [];
if (matches.length !== 1) throw new Error('No se encontró un único bloque const DATA para actualizar.');
fs.writeFileSync(target, html.replace(matches[0], `const DATA=${payload};`), 'utf8');

console.log(JSON.stringify({ source: path.resolve(source), target: path.resolve(target), records: records.length }));
