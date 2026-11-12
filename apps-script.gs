/**
 * Studio REVI — Briefing de Identidade Visual
 * Recebe as respostas do formulário (GitHub Pages), grava uma linha na planilha
 * e envia um e-mail com as respostas organizadas + PDF anexo.
 *
 * COMO IMPLANTAR (uma vez só):
 *   1. Crie uma planilha nova no Google Sheets (ex: "Briefings — Identidade Visual").
 *   2. Menu Extensões > Apps Script. Apague o código de exemplo e cole este arquivo inteiro.
 *   3. Salve (ícone de disquete).
 *   4. Implantar > Nova implantação > engrenagem > "App da Web".
 *        Executar como: Eu
 *        Quem pode acessar: Qualquer pessoa
 *      Clique em Implantar e autorize o acesso com a sua conta Google.
 *   5. Copie a URL que termina em /exec e mande para o Claude.
 *
 * PARA ATUALIZAR DEPOIS: Implantar > Gerenciar implantações > lápis >
 *   Versão: "Nova versão" > Implantar (a URL /exec continua a mesma).
 */

// ===================== CONFIG =====================
// Deixe vazio para enviar ao dono da planilha (a conta que implantou).
var EMAIL_DESTINO = '';
var SHEET_NAME = 'Briefings';
// ==================================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    var answers = body.answers || {};
    var sections = body.sections || [];

    var questions = [];
    sections.forEach(function (s) { s.questions.forEach(function (q) { questions.push(q); }); });

    salvarNaPlanilha(questions, answers);
    enviarEmail(sections, answers);

    return json({ success: true });
  } catch (err) {
    console.error(err);
    return json({ success: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, servico: 'Briefing de Identidade Visual — Studio REVI' });
}

/* ─── PLANILHA ─────────────────────────────────────── */
function salvarNaPlanilha(questions, answers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  var headers = ['Recebido em'].concat(questions.map(function (q) { return q.label; }));
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#030917').setFontColor('#4AA6F7');
    sheet.setFrozenRows(1);
  }

  var row = [new Date()].concat(questions.map(function (q) { return valor(answers[q.id], ''); }));
  sheet.appendRow(row);
}

/* ─── E-MAIL ───────────────────────────────────────── */
function enviarEmail(sections, answers) {
  var destino = EMAIL_DESTINO || Session.getEffectiveUser().getEmail();
  var marca = answers.marca || answers.nome || 'Cliente';
  var html = montarHTML(sections, answers);

  var pdf = Utilities.newBlob(html, 'text/html', 'briefing.html')
    .getAs('application/pdf')
    .setName('briefing-iv-' + slug(marca) + '.pdf');

  var opts = { htmlBody: html, attachments: [pdf], name: 'Briefing Studio REVI' };
  if (answers.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(answers.email)) opts.replyTo = answers.email;

  MailApp.sendEmail(destino, '[Briefing IV] ' + marca, 'Novo briefing de Identidade Visual recebido.', opts);
}

function montarHTML(sections, answers) {
  var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  var rows = '';
  sections.forEach(function (sec) {
    rows += '<tr><td colspan="2" style="background:#030917;color:#4AA6F7;font-weight:700;font-size:12px;letter-spacing:2px;padding:12px 16px;text-transform:uppercase">' +
      esc(sec.num + '. ' + sec.title) + '</td></tr>';
    sec.questions.forEach(function (q) {
      var raw = answers[q.id];
      var vazio = !raw || (Array.isArray(raw) && !raw.length);
      var val = vazio ? '<span style="color:#999">—</span>'
        : Array.isArray(raw)
          ? raw.map(function (v) { return '<span style="background:#eef2ff;color:#1141BD;padding:2px 8px;border-radius:4px;font-size:12px;margin:2px;display:inline-block">' + esc(v) + '</span>'; }).join(' ')
          : esc(raw).replace(/\n/g, '<br>');
      rows += '<tr><td style="padding:10px 16px;font-size:12px;font-weight:600;color:#444;border-bottom:1px solid #f0f0f0;width:35%;vertical-align:top">' + esc(q.label) +
        '</td><td style="padding:10px 16px;font-size:13px;color:#1a1a2e;border-bottom:1px solid #f0f0f0;vertical-align:top">' + val + '</td></tr>';
    });
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif">' +
    '<div style="max-width:700px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden">' +
    '<div style="background:#030917;padding:28px 32px"><div style="color:#fff;font-size:22px;font-weight:800">STUDIO REVI</div>' +
    '<div style="color:#4AA6F7;font-size:13px;margin-top:4px">Novo Briefing de Identidade Visual recebido</div></div>' +
    '<div style="padding:20px 32px;background:#f8f9ff;border-bottom:1px solid #e8eaf0;font-size:13px;color:#444">' +
    '<strong>Marca:</strong> ' + esc(answers.marca || '—') + ' &nbsp;|&nbsp; <strong>Cliente:</strong> ' + esc(answers.nome || '—') +
    ' (' + esc(answers.email || '—') + ') &nbsp;|&nbsp; <span style="color:#999">' + agora + '</span></div>' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">' + rows + '</table>' +
    '<div style="padding:20px 32px;background:#f8f9ff;text-align:center;font-size:11px;color:#999">Studio REVI</div>' +
    '</div></body></html>';
}

/* ─── UTILS ────────────────────────────────────────── */
function valor(raw, vazio) {
  if (!raw || (Array.isArray(raw) && !raw.length)) return vazio;
  return Array.isArray(raw) ? raw.join(', ') : String(raw);
}
function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function slug(s) {
  return String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
