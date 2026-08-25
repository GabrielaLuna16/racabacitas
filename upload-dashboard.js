(() => {
  const STORAGE_KEY = 'racaba-dashboard-excel-v1';

  function cleanGroup(value, fallback = '-') {
    const result = String(value ?? '').trim().replace(/\s*\(\s*\d+\s*\)\s*$/, '');
    return result || fallback;
  }

  function text(value, fallback = '') {
    const result = String(value ?? '').trim();
    return result || fallback;
  }

  function excelValue(value) {
    if (value instanceof Date) return value.toISOString();
    return value == null ? '' : value;
  }

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
    });

    if (!records.length) throw new Error('El archivo no contiene registros utilizables.');
    return records;
  }

  function applyRecords(records, filename, persist = true) {
    DATA.splice(0, DATA.length, ...records);
    activePeriod = null;
    projectFilter.value = '';
    hostFilter.value = '';
    buildTabs();
    render();
    uploadNote.classList.remove('error');
    uploadNote.textContent = `${records.length} registros cargados${filename ? ` · ${filename}` : ''}`;
    if (persist) {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ records, filename, savedAt: Date.now() }));
      } catch (_) {
        uploadNote.textContent += ' · No fue posible guardarlos para la próxima visita';
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
      applyRecords(parseWorkbook(await file.arrayBuffer()), file.name);
    } catch (error) {
      uploadNote.classList.add('error');
      uploadNote.textContent = `No se pudo cargar: ${error.message}`;
    } finally {
      uploadButton.disabled = false;
      uploadButton.textContent = 'Cargar Excel';
      excelInput.value = '';
    }
  });

  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (saved?.records?.length) applyRecords(saved.records, saved.filename, false);
  } catch (_) {
    localStorage.removeItem(STORAGE_KEY);
  }
})();
