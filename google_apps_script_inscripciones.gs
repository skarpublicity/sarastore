const SHEET_NAME = 'Inscripciones';
const WHATSAPP_PHONE = '573104993263';

function doPost(e) {
  const payload = parsePayload_(e);
  const sheet = getOrCreateSheet_();

  sheet.appendRow([
    payload.timestamp || new Date().toISOString(),
    payload.nombre || '',
    payload.cedula || '',
    payload.matricula_profesional || '',
    payload.telefono || '',
    payload.email || '',
    payload.tipo_inscripcion || '',
    payload.metodo_pago || '',
    payload.user_agent || '',
    payload.origen || ''
  ]);

  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet_() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
    sheet.appendRow([
      'timestamp',
      'nombre',
      'cedula',
      'matricula_profesional',
      'telefono',
      'email',
      'tipo_inscripcion',
      'metodo_pago',
      'user_agent',
      'origen'
    ]);
  }

  return sheet;
}

function parsePayload_(e) {
  let payload = {};

  try {
    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '';
    payload = JSON.parse(rawBody || '{}');
  } catch (error) {
    payload = {};
  }

  if (!payload || Object.keys(payload).length === 0) {
    const params = e && e.parameter ? e.parameter : {};
    payload = {
      timestamp: params.timestamp || new Date().toISOString(),
      nombre: params.nombre || '',
      cedula: params.cedula || '',
      matricula_profesional: params.matricula_profesional || '',
      telefono: params.telefono || '',
      email: params.email || '',
      tipo_inscripcion: params.tipo_inscripcion || '',
      metodo_pago: params.metodo_pago || '',
      user_agent: params.user_agent || '',
      origen: params.origen || ''
    };
  }

  return payload;
}

function buildSuccessPage_(payload) {
  const message = [
    'Hola, ya realice mi inscripcion al XII Seminario Regional VEPA Cauca.',
    '',
    'Nombre: ' + (payload.nombre || ''),
    'Cedula: ' + (payload.cedula || ''),
    'Telefono: ' + (payload.telefono || ''),
    'Email: ' + (payload.email || ''),
    'Matricula profesional: ' + (payload.matricula_profesional || 'No aplica'),
    'Tipo de inscripcion: ' + (payload.tipo_inscripcion || ''),
    'Metodo de pago: ' + (payload.metodo_pago || ''),
    '',
    'Adjunto mi comprobante para validacion.'
  ].join('\n');

  const whatsappUrl = 'https://wa.me/' + WHATSAPP_PHONE + '?text=' + encodeURIComponent(message);
  const encodedMessage = encodeURIComponent(message);

  return `
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Inscripcion registrada</title>
  <style>
    :root{
      --bg:#08172d;
      --card:rgba(255,255,255,0.08);
      --stroke:rgba(255,255,255,0.14);
      --text:#eef5ff;
      --muted:#bfd3ee;
      --primary:#39d0c7;
      --primary-strong:#85fff6;
      --success:#58d89b;
      --shadow:0 24px 70px rgba(0,0,0,0.28);
    }
    *{box-sizing:border-box}
    body{
      margin:0;
      font-family:Inter,system-ui,sans-serif;
      color:var(--text);
      background:
        radial-gradient(circle at top left, rgba(57,208,199,0.14), transparent 30%),
        radial-gradient(circle at top right, rgba(94,124,255,0.24), transparent 32%),
        linear-gradient(180deg,#061222 0%,#091b34 38%,#07111f 100%);
      min-height:100vh;
      display:grid;
      place-items:center;
      padding:24px;
    }
    .card{
      width:min(760px,100%);
      background:linear-gradient(180deg,rgba(255,255,255,.09),rgba(255,255,255,.05));
      border:1px solid rgba(255,255,255,.12);
      border-radius:28px;
      padding:32px;
      box-shadow:var(--shadow);
      backdrop-filter:blur(12px);
    }
    .eyebrow{
      display:inline-flex;
      align-items:center;
      gap:10px;
      padding:8px 14px;
      border-radius:999px;
      font-size:.86rem;
      font-weight:700;
      letter-spacing:.04em;
      text-transform:uppercase;
      color:#dffcf7;
      background:rgba(57,208,199,.14);
      border:1px solid rgba(57,208,199,.18);
    }
    h1{
      margin:18px 0 12px;
      font-size:clamp(2rem,4vw,3rem);
      line-height:1.02;
      letter-spacing:-.04em;
    }
    p{margin:0 0 12px;color:var(--muted);font-size:1.03rem;line-height:1.65}
    .ok{
      margin-top:18px;
      padding:18px 20px;
      border-radius:18px;
      border:1px solid rgba(88,216,155,.28);
      background:rgba(88,216,155,.12);
      color:#eafff1;
      font-weight:700;
    }
    .summary{
      margin-top:20px;
      padding:20px;
      border-radius:20px;
      background:rgba(255,255,255,.06);
      border:1px solid rgba(255,255,255,.14);
      display:grid;
      gap:8px;
    }
    .actions{
      display:flex;
      flex-wrap:wrap;
      gap:14px;
      margin-top:24px;
    }
    .btn{
      display:inline-flex;
      align-items:center;
      justify-content:center;
      border-radius:999px;
      font-weight:800;
      border:none;
      text-decoration:none;
      cursor:pointer;
      font:inherit;
      padding:16px 24px;
    }
    .btn-primary{
      color:#06243a;
      background:linear-gradient(135deg,var(--primary),var(--primary-strong));
      box-shadow:0 18px 34px rgba(57,208,199,.26);
    }
    .btn-secondary{
      color:var(--text);
      border:1px solid rgba(255,255,255,.15);
      background:rgba(255,255,255,.07);
    }
    @media (max-width:820px){
      .card{padding:24px}
      .actions{grid-template-columns:1fr}
      .btn{width:100%}
    }
  </style>
</head>
<body>
  <div class="card">
    <span class="eyebrow">Registro exitoso</span>
    <h1>Tu inscripcion ya quedo registrada</h1>
    <p>La informacion se guardo correctamente en Google Sheets. El siguiente paso es enviar el comprobante directamente por WhatsApp.</p>
    <div class="ok">Ya puedes continuar con el envio del comprobante.</div>
    <div class="summary">
      <strong>Resumen del registro</strong>
      <span>Nombre: ${escapeHtml_(payload.nombre || '')}</span>
      <span>Cedula: ${escapeHtml_(payload.cedula || '')}</span>
      <span>Matricula profesional: ${escapeHtml_(payload.matricula_profesional || 'No aplica')}</span>
      <span>Telefono: ${escapeHtml_(payload.telefono || '')}</span>
      <span>Email: ${escapeHtml_(payload.email || '')}</span>
      <span>Tipo de inscripcion: ${escapeHtml_(payload.tipo_inscripcion || '')}</span>
      <span>Metodo de pago: ${escapeHtml_(payload.metodo_pago || '')}</span>
    </div>
    <div class="actions">
      <a class="btn btn-primary" href="${whatsappUrl}">Abrir WhatsApp</a>
      <a class="btn btn-secondary" href="https://wa.me/?text=${encodedMessage}">Copiar mensaje</a>
    </div>
  </div>
</body>
</html>`;
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
