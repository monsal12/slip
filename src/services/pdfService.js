const fs = require("fs");
const path = require("path");
const PDFDocument = require("pdfkit");
const QRCode = require("qrcode");
const { formatRupiah } = require("../utils/currency");

function drawRow(doc, label, value, y, options = {}) {
  const xLabel = options.xLabel || 40;
  const xValue = options.xValue || 340;
  const showDashForZero = options.showDashForZero || false;
  const displayValue = showDashForZero && Number(value) === 0 ? "-" : formatRupiah(value);

  doc.font("Helvetica").fontSize(11).text(label, xLabel, y);
  doc.font("Helvetica").fontSize(11).text("Rp", xValue - 35, y, { width: 30, align: "left" });
  doc
    .font("Helvetica-Bold")
    .fontSize(11)
    .text(displayValue, xValue, y, { width: 180, align: "right" });
}

async function generateSlipPdf(payload) {
  const outputDir = path.resolve("generated_slips");
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const fileName = `${payload.slipNumber}.pdf`;
  const filePath = path.join(outputDir, fileName);
  const isDokterUmum =
    payload.slipVariant === "dokter_umum" ||
    (payload.slipVariant !== "karyawan" && payload.employee.position === "Dokter Umum");
  const displayOptions = {
    showGajiJasa: payload.displayOptions?.showGajiJasa !== false,
    showGajiJaga: payload.displayOptions?.showGajiJaga === true,
    showBpjsPendapatan: payload.displayOptions?.showBpjsPendapatan !== false,
    showBonus: payload.displayOptions?.showBonus !== false,
    showPotonganLain: payload.displayOptions?.showPotonganLain !== false
  };

  const doc = new PDFDocument({ size: "A4", margin: 30 });
  const stream = fs.createWriteStream(filePath);
  doc.pipe(stream);
  const kopPath = path.resolve("public", "images", "header.png");

  const signatureBlocks = [
    {
      role: "Penerima",
      name: payload.employee.name,
      title: ""
    },
    {
      role: "Mengetahui",
      name: "Ns. Rehmaita Malem, S.Kep., M.Kep",
      title: "Kabid. SDM dan Umum"
    },
    {
      role: "Menyetujui",
      name: "dr. Arief Tirtana Putra, M.Si",
      title: "Direktur"
    }
  ];

  const qrBuffers = await Promise.all(
    signatureBlocks.map(async (item) => {
      const qrPayload = [
        `No Slip: ${payload.slipNumber}`,
        `Peran: ${item.role}`,
        `Nama: ${item.name}`,
        `Periode: ${payload.periodLabel}`,
        `Total Diterima: Rp ${formatRupiah(payload.totalDiterima)}`
      ].join("\n");

      try {
        return await QRCode.toBuffer(qrPayload, {
          errorCorrectionLevel: "M",
          margin: 1,
          width: 180
        });
      } catch (error) {
        return null;
      }
    })
  );

  try {

      doc.rect(20, 20, 555, 780).stroke();
      if (fs.existsSync(kopPath)) {
        doc.image(kopPath, 30, 30, { fit: [535, 86], align: "center" });
      } else {
        doc.font("Helvetica-Bold").fontSize(38).fillColor("#31c481").text("RUMAH SAKIT MULIA RAYA", 40, 45, {
          align: "center"
        });
        doc
          .moveTo(120, 100)
          .lineTo(540, 100)
          .lineWidth(3)
          .strokeColor("#31c481")
          .stroke();
      }

      doc.fillColor("black").font("Helvetica-Bold").fontSize(15).text("SLIP GAJI KARYAWAN", 0, 120, {
        align: "center"
      });

      doc.font("Helvetica").fontSize(11);
      doc.text("Institusi", 35, 160);
      doc.text(":", 110, 160);
      doc.font("Helvetica-Bold").text(payload.institution, 120, 160);

      doc.font("Helvetica").text("Periode", 35, 180);
      doc.text(":", 110, 180);
      doc.font("Helvetica-Bold").text(payload.periodLabel, 120, 180);

      doc.rect(30, 210, 535, 24).fillAndStroke("#d9d9d9", "#d9d9d9");
      doc.fillColor("black").font("Helvetica-Bold").fontSize(12).text("Data Karyawan", 40, 216);

      doc.font("Helvetica").fontSize(11);
      doc.text("Nama", 35, 245);
      doc.text(":", 110, 245);
      doc.font("Helvetica-Bold").text(payload.employee.name, 120, 245);

      doc.font("Helvetica").text("Posisi", 35, 265);
      doc.text(":", 110, 265);
      doc.font("Helvetica-Bold").text(payload.employee.position, 120, 265);

      doc.font("Helvetica").text("No Rekening", 35, 285);
      doc.text(":", 110, 285);
      doc.font("Helvetica-Bold").text(payload.employee.accountNumber || "-", 120, 285);

      doc.rect(30, 318, 535, 24).fillAndStroke("#d9d9d9", "#d9d9d9");
      doc.fillColor("black").font("Helvetica-Bold").fontSize(12).text("Pendapatan", 40, 324);

      let y = 350;
      drawRow(doc, "Gaji Pokok", payload.salary.gajiPokok, y);
      if (displayOptions.showGajiJasa || payload.salary.gajiJasa > 0) {
        y += 24;
        drawRow(doc, "Gaji Jasa", payload.salary.gajiJasa, y, { showDashForZero: isDokterUmum });
      }
      if (displayOptions.showGajiJaga || payload.salary.gajiJaga > 0) {
        y += 24;
        drawRow(doc, "Gaji Jaga", payload.salary.gajiJaga || 0, y, {
          showDashForZero: isDokterUmum
        });
      }
      if (displayOptions.showBpjsPendapatan) {
        y += 24;
        drawRow(doc, "BPJS Ketenagakerjaan", payload.salary.bpjsKetenagakerjaanPendapatan, y, {
          showDashForZero: isDokterUmum
        });
        y += 24;
        drawRow(doc, "BPJS Kesehatan", payload.salary.bpjsKesehatanPendapatan, y, {
          showDashForZero: isDokterUmum
        });
      }
      if (displayOptions.showBonus || payload.salary.bonus > 0) {
        y += 24;
        drawRow(doc, "Bonus", payload.salary.bonus, y, { showDashForZero: isDokterUmum });
      }
      y += 28;
      drawRow(doc, "Total Pendapatan", payload.totalPendapatan, y, { xLabel: 40, xValue: 340 });

      y += 38;
      doc.rect(30, y, 535, 24).fillAndStroke("#d9d9d9", "#d9d9d9");
      doc.fillColor("black").font("Helvetica-Bold").fontSize(12).text("Pengurangan", 40, y + 6);

      y += 32;
      drawRow(doc, "BPJS Ketenagakerjaan", payload.salary.bpjsKetenagakerjaanPotongan, y);
      y += 24;
      drawRow(doc, "BPJS Kesehatan", payload.salary.bpjsKesehatanPotongan, y, { showDashForZero: isDokterUmum });
      if (displayOptions.showPotonganLain || payload.salary.potonganLain > 0) {
        y += 24;
        drawRow(doc, "Potongan Lain", payload.salary.potonganLain, y, { showDashForZero: isDokterUmum });
      }
      y += 28;
      drawRow(doc, "Total Pengurangan", payload.totalPengurangan, y);

      y += 40;
      doc.rect(30, y, 535, 30).fillAndStroke("#cfcfcf", "#cfcfcf");
      doc
        .fillColor("black")
        .font("Helvetica-Bold")
        .fontSize(13)
        .text(isDokterUmum ? "Total Pendapatan" : "Total Pendapatan yang Diterima", 40, y + 8);
      doc.font("Helvetica-Bold").text(`Rp ${formatRupiah(payload.totalDiterima)}`, 340, y + 8, { width: 180, align: "right" });

      doc.rect(30, 690, 535, 1).fillAndStroke("#000", "#000");
      const signatureStartX = 30;
      const signatureWidth = 535;
      const colWidth = signatureWidth / 3;
      const signatureColumns = [
        { x: signatureStartX, width: colWidth },
        { x: signatureStartX + colWidth, width: colWidth },
        { x: signatureStartX + colWidth * 2, width: colWidth }
      ];
      const qrY = 720;
      const qrSize = 50;

      signatureBlocks.forEach((item, index) => {
        const column = signatureColumns[index];
        const centerX = column.x + column.width / 2;

        doc.font("Helvetica").fontSize(11).text(item.role, column.x, 708, {
          width: column.width,
          align: "center"
        });

        const qrBuffer = qrBuffers[index];
        if (qrBuffer) {
          doc.image(qrBuffer, centerX - qrSize / 2, qrY, { fit: [qrSize, qrSize] });
        }

        doc.font("Helvetica-Bold").fontSize(9).text(item.name, column.x, 774, {
          width: column.width,
          align: "center"
        });

        if (item.title) {
          doc.font("Helvetica").fontSize(8).text(item.title, column.x, 786, {
            width: column.width,
            align: "center"
          });
        }
      });

    doc.end();

    await new Promise((resolve, reject) => {
      stream.on("finish", resolve);
      stream.on("error", reject);
    });

    return filePath;
  } catch (error) {
    throw error;
  }
}

module.exports = { generateSlipPdf };
