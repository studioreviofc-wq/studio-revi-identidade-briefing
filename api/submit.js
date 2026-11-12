const { Resend } = require('resend');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, BorderStyle } = require('docx');

const STUDIO_EMAIL = process.env.STUDIO_EMAIL || 'studiorevi.ofc@gmail.com';

/* ─── PDF ──────────────────────────────────────────── */
function generatePDF(sections, answers) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 60, size: 'A4', info: { Title: 'Briefing de Identidade Visual — Studio REVI', Author: 'Studio REVI' } });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end',  () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const W    = doc.page.width;
    const DARK = '#030917';
    const BLUE = '#1141BD';
    const LBLU = '#4AA6F7';

    doc.rect(0, 0, W, 72).fill(DARK);
    doc.fillColor('#fff').fontSize(20).font('Helvetica-Bold').text('STUDIO REVI', 60, 20);
    doc.fillColor(LBLU).fontSize(11).font('Helvetica').text('Briefing de Identidade Visual', 60, 44);
    const now = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    doc.fillColor('#888').fontSize(8).text(`Recebido em: ${now}`, W - 220, 56, { width: 160, align: 'right' });
    doc.moveDown(2);

    const proj  = answers.marca || '—';
    const email = answers.email || '—';
    doc.fillColor('#444').fontSize(10).font('Helvetica')
       .text(`Marca: ${proj}   |   Cliente: ${email}`, 60, doc.y);
    doc.moveDown(1.5);

    sections.forEach(sec => {
      if (doc.y > doc.page.height - 160) doc.addPage();
      const sy = doc.y;
      doc.rect(60, sy, W - 120, 32).fill(DARK);
      doc.fillColor(LBLU).fontSize(8).font('Helvetica-Bold').text(sec.num, 72, sy + 10);
      doc.fillColor('#fff').fontSize(13).font('Helvetica-Bold').text(sec.title, 92, sy + 10, { width: W - 170 });
      doc.moveDown(1.8);

      sec.questions.forEach(q => {
        if (doc.y > doc.page.height - 120) doc.addPage();
        const raw = answers[q.id];
        const val = (!raw || (Array.isArray(raw) && !raw.length))
          ? 'Não respondido'
          : Array.isArray(raw) ? raw.join(', ') : String(raw);

        doc.fillColor(DARK).fontSize(10).font('Helvetica-Bold').text(q.label, 60, doc.y, { width: W - 120 });
        doc.moveDown(0.25);
        const ay = doc.y;
        doc.rect(60, ay, 2.5, 14).fill(BLUE);
        doc.fillColor(val === 'Não respondido' ? '#bbb' : '#222')
           .fontSize(10).font('Helvetica').text(val, 70, ay, { width: W - 132 });
        doc.moveDown(1.1);
      });
      doc.moveDown(0.5);
    });

    const fy = doc.page.height - 32;
    doc.fillColor('#bbb').fontSize(8).font('Helvetica')
       .text('Studio REVI  ·  studiorevi.ofc@gmail.com', 60, fy, { width: W - 120, align: 'center' });
    doc.end();
  });
}

/* ─── DOCX ─────────────────────────────────────────── */
async function generateDOCX(sections, answers) {
  const now  = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  const proj = answers.marca || '—';
  const email= answers.email || '—';

  const children = [
    new Paragraph({ children: [new TextRun({ text: 'STUDIO REVI — BRIEFING DE IDENTIDADE VISUAL', bold: true, size: 44, color: '1141BD' })], spacing: { after: 160 } }),
    new Paragraph({ children: [new TextRun({ text: `Marca: ${proj}`, bold: true, size: 24 }), new TextRun({ text: `   |   Cliente: ${email}`, size: 24, color: '666666' })], spacing: { after: 80 } }),
    new Paragraph({ children: [new TextRun({ text: `Recebido em: ${now}`, size: 20, color: '999999' })], spacing: { after: 480 } }),
  ];

  sections.forEach(sec => {
    children.push(new Paragraph({
      children: [new TextRun({ text: `${sec.num}. ${sec.title.toUpperCase()}`, bold: true, size: 28, color: '1141BD' })],
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 480, after: 160 },
      border: { bottom: { color: '4AA6F7', style: BorderStyle.SINGLE, size: 6 } },
    }));

    sec.questions.forEach(q => {
      const raw = answers[q.id];
      const val = (!raw || (Array.isArray(raw) && !raw.length))
        ? '—'
        : Array.isArray(raw) ? raw.join(', ') : String(raw);

      children.push(new Paragraph({ children: [new TextRun({ text: q.label, bold: true, size: 22 })], spacing: { before: 240, after: 80 } }));
      children.push(new Paragraph({ children: [new TextRun({ text: val, size: 22, color: val === '—' ? '999999' : '1A1A2E', italics: val === '—' })], indent: { left: 360 }, spacing: { after: 120 } }));
    });
  });

  const docx = new Document({ creator: 'Studio REVI', title: `Briefing IV — ${proj}`, sections: [{ properties: {}, children }] });
  return await Packer.toBuffer(docx);
}

/* ─── EMAIL HTML ────────────────────────────────────── */
const escH = v => String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
function buildEmailHTML(sections, answers) {
  const proj  = answers.marca || '—';
  const email = answers.email || '—';
  const now   = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });

  let rows = '';
  sections.forEach(sec => {
    rows += `<tr><td colspan="2" style="background:#030917;color:#4AA6F7;font-weight:700;font-size:12px;letter-spacing:2px;padding:12px 16px;text-transform:uppercase;">${sec.num}. ${sec.title}</td></tr>`;
    sec.questions.forEach(q => {
      const raw = answers[q.id];
      const val = (!raw || (Array.isArray(raw) && !raw.length))
        ? '<span style="color:#999">—</span>'
        : Array.isArray(raw)
          ? raw.map(v => `<span style="background:#eef2ff;color:#1141BD;padding:2px 8px;border-radius:4px;font-size:12px;margin:2px;display:inline-block">${escH(v)}</span>`).join(' ')
          : String(raw).replace(/\n/g, '<br>');
      rows += `<tr>
        <td style="padding:10px 16px;font-size:12px;font-weight:600;color:#444;border-bottom:1px solid #f0f0f0;width:35%;vertical-align:top">${q.label}</td>
        <td style="padding:10px 16px;font-size:13px;color:#1a1a2e;border-bottom:1px solid #f0f0f0;vertical-align:top">${val}</td>
      </tr>`;
    });
  });

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f5f5f5;font-family:Inter,Arial,sans-serif">
    <div style="max-width:700px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08)">
      <div style="background:#030917;padding:28px 32px">
        <div style="color:#fff;font-size:22px;font-weight:800;letter-spacing:-0.5px">STUDIO REVI</div>
        <div style="color:#4AA6F7;font-size:13px;margin-top:4px">Novo Briefing de Identidade Visual recebido</div>
      </div>
      <div style="padding:20px 32px;background:#f8f9ff;border-bottom:1px solid #e8eaf0">
        <span style="font-size:13px;color:#444"><strong>Marca:</strong> ${escH(proj)}</span>
        &nbsp;&nbsp;|&nbsp;&nbsp;
        <span style="font-size:13px;color:#444"><strong>Cliente:</strong> ${escH(email)}</span>
        &nbsp;&nbsp;|&nbsp;&nbsp;
        <span style="font-size:13px;color:#999">${now}</span>
      </div>
      <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">${rows}</table>
      <div style="padding:20px 32px;background:#f8f9ff;text-align:center;font-size:11px;color:#999">Studio REVI · studiorevi.ofc@gmail.com</div>
    </div>
  </body></html>`;
}

/* ─── HANDLER ───────────────────────────────────────── */
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')   return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { answers, sections } = req.body;
    if (!answers || !sections) return res.status(400).json({ error: 'Missing answers or sections' });

    const [pdfBuffer, docxBuffer] = await Promise.all([
      generatePDF(sections, answers),
      generateDOCX(sections, answers),
    ]);

    const projName = answers.marca || answers.email || 'Cliente';
    const resend = new Resend(process.env.RESEND_API_KEY);

    await resend.emails.send({
      from:        'Briefing Studio REVI <onboarding@resend.dev>',
      reply_to:    answers.email || undefined,
      to:          [STUDIO_EMAIL],
      subject:     `[Briefing IV] ${projName}`,
      html:        buildEmailHTML(sections, answers),
      attachments: [
        { filename: `briefing-iv-${projName.replace(/\s+/g,'-').toLowerCase()}.pdf`,  content: pdfBuffer.toString('base64') },
        { filename: `briefing-iv-${projName.replace(/\s+/g,'-').toLowerCase()}.docx`, content: docxBuffer.toString('base64') },
      ],
    });

    return res.status(200).json({ success: true });
  } catch (err) {
    console.error('submit error:', err);
    return res.status(500).json({ error: 'Internal server error' });
  }
};
