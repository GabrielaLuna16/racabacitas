(() => {
  const SHARED_DATA_URL = '/api/dashboard-data';
  const text = (value, fallback = '') => String(value ?? '').trim() || fallback;
  const cleanGroup = (value, fallback = '-') => text(value).replace(/\s*\(\s*\d+\s*\)\s*$/, '') || fallback;
  const excelValue = value => value instanceof Date ? value.toISOString() : (value == null ? '' : value);

  function parseWorkbook(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    const headerIndex = rows.findIndex(row => row.includes('Full Name') && row.includes('Created Time'));
    if (headerIndex < 0) throw new Error('No se encontraron las columnas Full Name y Created Time.');
    const headers = rows[headerIndex].map(value => text(value));
    let currentProject = '-';
    let currentHost = '-';
    const records = [];
    rows.slice(headerIndex + 1).forEach(values => {
      const raw = Object.fromEntries(headers.map((header, index) => [header, values[index]]));
      const contacto = text(raw['Full Name']);
      if (!contacto) return;
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
        id: records.length + 1, proyecto, asesor, contacto,
        campana: text(raw['Campaign'], 'Sin campaña'), fecha: excelValue(meetingDate),
        nota: text(raw['Note Content']), contactStatus: text(raw['Contact Status'], 'Sin estatus'),
        citaStatus: text(raw['Estatus de cita (Meetings)'], 'Sin estatus'), anuncio: text(raw['Anuncio MKT'], 'Sin anuncio'),
        fechaApartado: excelValue(raw['Fecha Apartado']), fechaContrato: excelValue(raw['Fecha Contrato']),
        recordId: text(raw['Record Id']).replace(/^zcrm_/i, ''), fechaCreacion: excelValue(raw['Created Time'])
      });
    });
    if (!records.length) throw new Error('El archivo no contiene registros utilizables.');
    return records;
  }

  function applyRecords(records, filename) {
    DATA.splice(0, DATA.length, ...records);
    activePeriod = null;
    projectFilter.value = '';
    hostFilter.value = '';
    buildTabs();
    render();
    uploadNote.classList.remove('error');
    uploadNote.textContent = `${records.length} registros compartidos${filename ? ` · ${filename}` : ''}`;
  }

  async function requestJson(url, options) {
    const response = await fetch(url, options);
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || 'No fue posible actualizar los datos.');
    return result;
  }

  async function loadSharedRecords() {
    try {
      const result = await requestJson(`${SHARED_DATA_URL}?t=${Date.now()}`, { cache: 'no-store' });
      if (result?.records?.length) applyRecords(result.records, result.filename);
    } catch (error) {
      if (!/Aún no hay|todavía no está configurado/.test(error.message)) {
        uploadNote.classList.add('error');
        uploadNote.textContent = `Datos incluidos en la página · ${error.message}`;
      }
    }
  }

  uploadButton.addEventListener('click', () => excelInput.click());
  excelInput.addEventListener('change', async event => {
    const file = event.target.files[0];
    if (!file) return;
    uploadButton.disabled = true;
    uploadButton.textContent = 'Procesando…';
    uploadNote.classList.remove('error');
    uploadNote.textContent = 'Leyendo archivo…';
    try {
      const records = parseWorkbook(await file.arrayBuffer());
      uploadNote.textContent = 'Publicando para todos los usuarios…';
      await requestJson(SHARED_DATA_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ records, filename: file.name }) });
      applyRecords(records, file.name);
      uploadNote.textContent = `${records.length} registros publicados para todos · ${file.name}`;
    } catch (error) {
      uploadNote.classList.add('error');
      uploadNote.textContent = `No se pudo publicar: ${error.message}`;
    } finally {
      uploadButton.disabled = false;
      uploadButton.textContent = 'Cargar Excel';
      excelInput.value = '';
    }
  });
  loadSharedRecords();
})();
