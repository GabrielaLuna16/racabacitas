import { list, put } from '@vercel/blob';

const DATA_PATH = 'racaba/dashboard-data.json';
const MAX_RECORDS = 20000;

function send(response, status, body) {
  response.setHeader('Cache-Control', 'no-store, max-age=0');
  return response.status(status).json(body);
}

export default async function handler(request, response) {
  if (!process.env.BLOB_READ_WRITE_TOKEN) return send(response, 503, { error: 'El almacenamiento compartido todavía no está configurado.' });
  try {
    if (request.method === 'GET') {
      const result = await list({ prefix: DATA_PATH, limit: 1 });
      const blob = result.blobs.find(item => item.pathname === DATA_PATH);
      if (!blob) return send(response, 404, { error: 'Aún no hay una actualización compartida.' });
      const updated = new Date(blob.uploadedAt).getTime();
      const blobResponse = await fetch(`${blob.url}?v=${updated}`, { cache: 'no-store' });
      if (!blobResponse.ok) throw new Error('No fue posible leer la actualización compartida.');
      return send(response, 200, await blobResponse.json());
    }
    if (request.method === 'POST') {
      const { records, filename } = request.body || {};
      if (!Array.isArray(records) || !records.length || records.length > MAX_RECORDS) return send(response, 400, { error: 'Los datos del Excel no son válidos.' });
      const payload = { records, filename: String(filename || 'Excel actualizado').slice(0, 180), savedAt: new Date().toISOString() };
      await put(DATA_PATH, JSON.stringify(payload), { access: 'public', allowOverwrite: true, contentType: 'application/json', cacheControlMaxAge: 60 });
      return send(response, 200, { ok: true, count: records.length, savedAt: payload.savedAt });
    }
    response.setHeader('Allow', 'GET, POST');
    return send(response, 405, { error: 'Método no permitido.' });
  } catch (error) {
    console.error(error);
    return send(response, 500, { error: error.message || 'Error al actualizar los datos.' });
  }
}
