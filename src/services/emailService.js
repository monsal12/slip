const nodemailer = require("nodemailer");

function smtpConfigured() {
  return Boolean(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS);
}

function createTransporter() {
  const secure = String(process.env.SMTP_SECURE || "false") === "true";

  return nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });
}

async function verifySmtpConnection() {
  if (!smtpConfigured()) {
    throw new Error("SMTP belum dikonfigurasi.");
  }

  const transporter = createTransporter();
  await transporter.verify();
}

async function sendSlipEmail({ to, employeeName, periodLabel, pdfPath }) {
  if (!smtpConfigured()) {
    throw new Error("SMTP belum dikonfigurasi. Isi SMTP_HOST, SMTP_USER, dan SMTP_PASS di .env");
  }

  const transporter = createTransporter();
  const fromName = process.env.SMTP_FROM_NAME || "HRD";
  const smtpUser = process.env.SMTP_USER;
  const fromEmail = process.env.SMTP_FROM_EMAIL || smtpUser;

  await transporter.sendMail({
    from: `${fromName} <${fromEmail}>`,
    to,
    subject: `Slip Gaji ${periodLabel} - ${employeeName}`,
    text: `Halo ${employeeName},\n\nTerlampir slip gaji periode ${periodLabel}.\n\nSalam,\nHRD`,
    attachments: [
      {
        filename: `slip-gaji-${periodLabel.replace(/\s+/g, "-").toLowerCase()}.pdf`,
        path: pdfPath
      }
    ]
  });
}

module.exports = { sendSlipEmail, smtpConfigured, verifySmtpConnection };
