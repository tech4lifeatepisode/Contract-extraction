/**
 * HTTP server for Render Web Service (binds to PORT).
 * Optional: AUTO_EXTRACT_FROM_STORAGE=true runs extraction from Supabase Storage after listen.
 * Optional: GET/POST /extract with header X-Extract-Secret: EXTRACT_TRIGGER_SECRET
 */
import http from 'http';
import { runExtractFromSupabaseStorage } from './extract-from-storage.mjs';

const port = Number(process.env.PORT) || 3000;

let extractionRunning = false;

async function safeRunExtract(label) {
  if (extractionRunning) {
    console.warn(`[${label}] Extraction already running, skip.`);
    return { ok: false, skipped: true, message: 'already_running' };
  }
  extractionRunning = true;
  try {
    const r = await runExtractFromSupabaseStorage();
    console.log(`[${label}] Extraction finished:`, r);
    return { ok: true, ...r };
  } catch (e) {
    console.error(`[${label}] Extraction failed:`, e);
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  } finally {
    extractionRunning = false;
  }
}

const server = http.createServer(async (req, res) => {
  const url = req.url?.split('?')[0] || '/';

  if (url === '/' || url === '/health') {
    res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(
      'Contract extraction service OK.\n' +
        'Set AUTO_EXTRACT_FROM_STORAGE=true to pull PDFs from Supabase Storage on boot.\n' +
        'Trigger manually: POST /extract with header X-Extract-Secret: <EXTRACT_TRIGGER_SECRET> (if set).\n',
    );
    return;
  }

  if (url === '/extract' && (req.method === 'POST' || req.method === 'GET')) {
    const secret = process.env.EXTRACT_TRIGGER_SECRET;
    const provided = req.headers['x-extract-secret'];
    if (secret && provided !== secret) {
      res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ ok: false, error: 'unauthorized' }));
      return;
    }
    if (!process.env.OPENAI_API_KEY) {
      res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ ok: false, error: 'OPENAI_API_KEY not set' }));
      return;
    }

    res.writeHead(202, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: true, accepted: true, message: 'extraction started in background' }));

    setImmediate(() => {
      safeRunExtract('http').catch((e) => console.error(e));
    });
    return;
  }

  res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('Not found\n');
});

server.listen(port, '0.0.0.0', () => {
  console.log(`Listening on 0.0.0.0:${port}`);

  if (process.env.AUTO_EXTRACT_FROM_STORAGE === 'true') {
    console.log('AUTO_EXTRACT_FROM_STORAGE: scheduling extraction from Supabase Storage...');
    setImmediate(() => {
      safeRunExtract('startup').catch((e) => console.error(e));
    });
  }
});
