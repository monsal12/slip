const express = require("express");
const path = require("path");
const dayjs = require("dayjs");
const multer = require("multer");
const XLSX = require("xlsx");
const Employee = require("../models/Employee");
const Slip = require("../models/Slip");
const { toNumber, formatRupiah } = require("../utils/currency");
const { generateSlipPdf } = require("../services/pdfService");
const { sendSlipEmail, smtpConfigured, verifySmtpConnection } = require("../services/emailService");

const router = express.Router();
const LOGIN_USER = process.env.APP_LOGIN_USER || "putri";
const LOGIN_PASS = process.env.APP_LOGIN_PASS || "putri";
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 }
});
const BULAN_ID = [
  "Januari",
  "Februari",
  "Maret",
  "April",
  "Mei",
  "Juni",
  "Juli",
  "Agustus",
  "September",
  "Oktober",
  "November",
  "Desember"
];
const ALLOWED_SLIP_VARIANTS = ["dokter_umum", "dokter_spesialis", "karyawan"];

const TEMPLATE_SHEETS = {
  KARYAWAN: "template_karyawan",
  DOKTER_UMUM: "template_dokter_umum",
  DOKTER_SPESIALIS: "template_dokter_spesialis",
  RADIOLOGI: "template_spesialis_radiologi"
};

const EMAIL_RETRY_LIMIT = Math.max(0, Number(process.env.EMAIL_RETRY_LIMIT || 3));
const EMAIL_RETRY_BASE_MS = Math.max(0, Number(process.env.EMAIL_RETRY_BASE_MS || 5000));
const BATCH_EMAIL_DELAY_MS = Math.max(0, Number(process.env.BATCH_EMAIL_DELAY_MS || 2000));

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function sendEmailWithRetry(payload, maxRetries) {
  let lastError = null;

  for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
    try {
      await sendSlipEmail(payload);
      return { ok: true };
    } catch (error) {
      lastError = error;
      if (attempt < maxRetries) {
        const backoff = EMAIL_RETRY_BASE_MS * Math.pow(2, attempt);
        await sleep(backoff);
      }
    }
  }

  return { ok: false, error: lastError };
}

function requireAuth(req, res, next) {
  if (req.session && req.session.isAuthenticated) {
    return next();
  }

  return res.redirect("/login");
}

function buildPeriodLabel(month, year) {
  const validMonth = Math.min(12, Math.max(1, Number(month) || 1));
  return `${BULAN_ID[validMonth - 1]} ${year}`;
}

function generateSlipNumber(indexHint = 0) {
  const timestamp = Date.now();
  const random = Math.floor(Math.random() * 9000) + 1000;
  return `SLIP-${timestamp}-${indexHint}-${random}`;
}

function parseBoolean(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ["1", "true", "yes", "y", "on", "kirim"].includes(normalized);
}

function parseOptionalBoolean(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) {
    return null;
  }
  if (["1", "true", "yes", "y", "on", "tampil"].includes(normalized)) {
    return true;
  }
  if (["0", "false", "no", "n", "off", "hide", "sembunyi"].includes(normalized)) {
    return false;
  }
  return null;
}

function normalizeSlipVariant(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) {
    return "dokter_spesialis";
  }

  const map = {
    otomatis: "dokter_spesialis",
    auto: "dokter_spesialis",
    "dokter umum": "dokter_umum",
    dokter_umum: "dokter_umum",
    dokum: "dokter_umum",
    "karyawan": "karyawan",
    "dokter spesialis": "dokter_spesialis",
    dokter_spesialis: "dokter_spesialis",
    "dokter umu": "dokter_umum",
    radiologi: "dokter_spesialis",
    "dokter spesialis radiologi": "dokter_spesialis"
  };

  return map[raw] || "dokter_spesialis";
}

function normalizePosition(value) {
  return String(value || "").trim();
}

function pickValue(source, keys, fallback = "") {
  for (const key of keys) {
    if (source[key] !== undefined && source[key] !== null && String(source[key]).trim() !== "") {
      return source[key];
    }
  }

  return fallback;
}

function resolveDisplayOptions(input, position, slipVariant) {
  const isDokterUmumProfile = slipVariant === "dokter_umum";
  const defaults = {
    showGajiJasa: isDokterUmumProfile,
    showJasaKjs: true,
    showGajiJaga: isDokterUmumProfile,
    showTunjangan: true,
    showTunjanganJabaran: true,
    showTunjanganHariRaya: true,
    showPengurangan: true,
    showBpjsPendapatan: true,
    showBonus: true,
    showPotonganLain: true
  };

  const overrides = {
    showGajiJasa: parseOptionalBoolean(pickValue(input, ["Tampilkan Gaji Jasa", "showGajiJasa"], "")),
    showJasaKjs: parseOptionalBoolean(pickValue(input, ["Tampilkan Jasa KJS", "showJasaKjs"], "")),
    showGajiJaga: parseOptionalBoolean(pickValue(input, ["Tampilkan Gaji Jaga", "showGajiJaga"], "")),
    showTunjangan: parseOptionalBoolean(pickValue(input, ["Tampilkan Tunjangan", "showTunjangan"], "")),
    showTunjanganJabaran: parseOptionalBoolean(
      pickValue(input, ["Tampilkan Tunjangan Jabaran", "showTunjanganJabaran"], "")
    ),
    showTunjanganHariRaya: parseOptionalBoolean(
      pickValue(input, ["Tampilkan Tunjangan Hari Raya", "showTunjanganHariRaya"], "")
    ),
    showPengurangan: parseOptionalBoolean(pickValue(input, ["Tampilkan Pemotongan", "showPengurangan"], "")),
    showBpjsPendapatan: parseOptionalBoolean(
      pickValue(input, ["Tampilkan BPJS Pendapatan", "showBpjsPendapatan"], "")
    ),
    showBonus: parseOptionalBoolean(pickValue(input, ["Tampilkan Bonus", "showBonus"], "")),
    showPotonganLain: parseOptionalBoolean(pickValue(input, ["Tampilkan Potongan Lain", "showPotonganLain"], ""))
  };

  return {
    showGajiJasa: overrides.showGajiJasa === null ? defaults.showGajiJasa : overrides.showGajiJasa,
    showJasaKjs: overrides.showJasaKjs === null ? defaults.showJasaKjs : overrides.showJasaKjs,
    showGajiJaga: overrides.showGajiJaga === null ? defaults.showGajiJaga : overrides.showGajiJaga,
    showTunjangan: overrides.showTunjangan === null ? defaults.showTunjangan : overrides.showTunjangan,
    showTunjanganJabaran:
      overrides.showTunjanganJabaran === null ? defaults.showTunjanganJabaran : overrides.showTunjanganJabaran,
    showTunjanganHariRaya:
      overrides.showTunjanganHariRaya === null ? defaults.showTunjanganHariRaya : overrides.showTunjanganHariRaya,
    showPengurangan: overrides.showPengurangan === null ? defaults.showPengurangan : overrides.showPengurangan,
    showBpjsPendapatan:
      overrides.showBpjsPendapatan === null ? defaults.showBpjsPendapatan : overrides.showBpjsPendapatan,
    showBonus: overrides.showBonus === null ? defaults.showBonus : overrides.showBonus,
    showPotonganLain: overrides.showPotonganLain === null ? defaults.showPotonganLain : overrides.showPotonganLain
  };
}

function inferTemplateMetaFromSheet(sheetName) {
  const normalized = String(sheetName || "").trim().toLowerCase();
  if (normalized.includes("karyawan")) {
    return { defaultPosition: "Karyawan", defaultVariant: "karyawan" };
  }
  if (normalized.includes("dokter_umum") || normalized.includes("dokter umum")) {
    return { defaultPosition: "Dokter Umum", defaultVariant: "dokter_umum" };
  }
  if (normalized.includes("radiologi")) {
    return { defaultPosition: "Dokter Spesialis", defaultVariant: "dokter_spesialis" };
  }
  if (normalized.includes("dokter_spesialis") || normalized.includes("dokter spesialis")) {
    return { defaultPosition: "Dokter Spesialis", defaultVariant: "dokter_spesialis" };
  }
  return { defaultPosition: "", defaultVariant: "dokter_spesialis" };
}

function normalizePayload(input) {
  const templateMeta = inferTemplateMetaFromSheet(input.__sheetName || "");
  const rawPosition = pickValue(input, ["position", "posisi", "Posisi"], templateMeta.defaultPosition);
  const slipVariant = normalizeSlipVariant(
    pickValue(input, ["Tipe Slip", "slipVariant", "Versi Slip"], templateMeta.defaultVariant)
  );

  const jasaPoli = toNumber(pickValue(input, ["Jasa Medis Pasien Poli", "jasaMedisPasienPoli"], 0));
  const jasaMadco = toNumber(pickValue(input, ["Jasa Medis Pasien Madco", "jasaMedisPasienMadco"], 0));
  const jasaRawatInap = toNumber(pickValue(input, ["Jasa Medis Rawat Inap", "jasaMedisRawatInap"], 0));
  const jasaOperasi = toNumber(pickValue(input, ["Jasa Medis Tindakan Operasi", "jasaMedisTindakanOperasi"], 0));
  const jasaRadiologi = toNumber(pickValue(input, ["Jasa Pasien Radiologi", "jasaPasienRadiologi"], 0));
  const pphPasal21 = toNumber(pickValue(input, ["PPH Pasal 21", "pphPasal21"], 0));
  const gajiJasaInput = toNumber(pickValue(input, ["gajiJasa", "Gaji Jasa"], 0));
  const jasaKjs = toNumber(pickValue(input, ["jasaKjs", "Jasa KJS"], 0));
  const gajiJasaComputed =
    jasaPoli + jasaMadco + jasaRawatInap + jasaOperasi + jasaRadiologi > 0
      ? jasaPoli + jasaMadco + jasaRawatInap + jasaOperasi + jasaRadiologi
      : gajiJasaInput;

  return {
    institution: pickValue(input, ["institution", "Institusi"], process.env.INSTITUSI_DEFAULT || "Mulia Rakan Membangun"),
    month: toNumber(pickValue(input, ["month", "bulan", "Bulan"], dayjs().month() + 1)),
    year: toNumber(pickValue(input, ["year", "tahun", "Tahun"], dayjs().year())),
    employeeName: pickValue(input, ["employeeName", "nama", "Nama Karyawan", "Nama", "Nama Lengkap"], ""),
    position: normalizePosition(rawPosition),
    slipVariant,
    accountNumber: pickValue(input, ["accountNumber", "noRekening", "No Rekening", "rekening"], "-"),
    email: pickValue(input, ["email", "Email", "Email Karyawan"], ""),
    gajiPokok: toNumber(pickValue(input, ["gajiPokok", "Gaji Pokok"], 0)),
    gajiJasa: gajiJasaComputed,
    jasaKjs,
    gajiJaga: toNumber(pickValue(input, ["gajiJaga", "Gaji Jaga"], 0)),
    tunjangan: toNumber(pickValue(input, ["tunjangan", "Tunjangan"], 0)),
    tunjanganJabaran: toNumber(pickValue(input, ["tunjanganJabaran", "Tunjangan Jabaran"], 0)),
    tunjanganHariRaya: toNumber(pickValue(input, ["tunjanganHariRaya", "Tunjangan Hari Raya"], 0)),
    bpjsKetenagakerjaanPendapatan: toNumber(
      pickValue(
        input,
        ["bpjsKetenagakerjaanPendapatan", "BPJS Ketenagakerjaan Pendapatan", "BPJS Ketenagakerjaan (Pendapatan)"],
        0
      )
    ),
    bpjsKesehatanPendapatan: toNumber(
      pickValue(input, ["bpjsKesehatanPendapatan", "BPJS Kesehatan Pendapatan", "BPJS Kesehatan (Pendapatan)"], 0)
    ),
    bonus: toNumber(pickValue(input, ["bonus", "Bonus"], 0)),
    bpjsKetenagakerjaanPotongan: toNumber(
      pickValue(
        input,
        [
          "bpjsKetenagakerjaanPotongan",
          "BPJS Ketenagakerjaan Potongan",
          "BPJS Ketenagakerjaan (Potongan)",
          "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)"
        ],
        0
      )
    ),
    bpjsKesehatanPotongan: toNumber(
      pickValue(
        input,
        [
          "bpjsKesehatanPotongan",
          "BPJS Kesehatan Potongan",
          "BPJS Kesehatan (Potongan)",
          "BPJS Kesehatan (Potongan Ditanggung Rumah Sakit)"
        ],
        0
      )
    ),
    potonganLain: toNumber(pickValue(input, ["potonganLain", "Potongan Lain"], 0)) + pphPasal21,
    jasaMedisPasienPoli: jasaPoli,
    jasaMedisPasienMadco: jasaMadco,
    jasaMedisRawatInap: jasaRawatInap,
    jasaMedisTindakanOperasi: jasaOperasi,
    jasaPasienRadiologi: jasaRadiologi,
    pphPasal21
  };
}

async function createSlipRecord(payload, shouldSendEmail, indexHint = 0) {
  const normalized = normalizePayload(payload);

  if (!normalized.employeeName || !normalized.position || !normalized.email) {
    throw new Error("Nama, posisi, dan email wajib diisi");
  }

  if (normalized.month < 1 || normalized.month > 12 || normalized.year < 2000) {
    throw new Error("Bulan/tahun tidak valid");
  }

  if (!ALLOWED_SLIP_VARIANTS.includes(normalized.slipVariant)) {
    throw new Error("Tipe slip wajib: Dokter Umum, Dokter Spesialis, atau Karyawan");
  }

  const displayOptions = resolveDisplayOptions(payload, normalized.position, normalized.slipVariant);

  const salary = {
    gajiPokok: normalized.gajiPokok,
    gajiJasa: displayOptions.showGajiJasa ? normalized.gajiJasa : 0,
    jasaKjs: displayOptions.showJasaKjs ? normalized.jasaKjs : 0,
    gajiJaga: displayOptions.showGajiJaga ? normalized.gajiJaga : 0,
    tunjangan: displayOptions.showTunjangan ? normalized.tunjangan : 0,
    tunjanganJabaran: displayOptions.showTunjanganJabaran ? normalized.tunjanganJabaran : 0,
    tunjanganHariRaya: displayOptions.showTunjanganHariRaya ? normalized.tunjanganHariRaya : 0,
    bpjsKetenagakerjaanPendapatan: displayOptions.showBpjsPendapatan
      ? normalized.bpjsKetenagakerjaanPendapatan
      : 0,
    bpjsKesehatanPendapatan: displayOptions.showBpjsPendapatan ? normalized.bpjsKesehatanPendapatan : 0,
    bonus: displayOptions.showBonus ? normalized.bonus : 0,
    bpjsKetenagakerjaanPotongan: displayOptions.showPengurangan ? normalized.bpjsKetenagakerjaanPotongan : 0,
    bpjsKesehatanPotongan: displayOptions.showPengurangan ? normalized.bpjsKesehatanPotongan : 0,
    potonganLain: displayOptions.showPengurangan && displayOptions.showPotonganLain ? normalized.potonganLain : 0
  };

  const totalPendapatan =
    salary.gajiPokok +
    salary.gajiJasa +
    salary.jasaKjs +
    salary.gajiJaga +
    salary.tunjangan +
    salary.tunjanganJabaran +
    salary.tunjanganHariRaya +
    salary.bpjsKetenagakerjaanPendapatan +
    salary.bpjsKesehatanPendapatan +
    salary.bonus;

  const totalPengurangan =
    salary.bpjsKetenagakerjaanPotongan + salary.bpjsKesehatanPotongan + salary.potonganLain;

  const totalDiterima = totalPendapatan - totalPengurangan;

  const employee = await Employee.findOneAndUpdate(
    { email: String(normalized.email).toLowerCase() },
    {
      name: normalized.employeeName,
      position: normalized.position,
      accountNumber: normalized.accountNumber || "-",
      email: String(normalized.email).toLowerCase()
    },
    { upsert: true, new: true, setDefaultsOnInsert: true }
  );

  const periodLabel = buildPeriodLabel(normalized.month, normalized.year);
  const slipNumber = generateSlipNumber(indexHint);

  const pdfPath = await generateSlipPdf({
    slipNumber,
    institution: normalized.institution,
    periodLabel,
    employee,
    slipVariant: normalized.slipVariant,
    displayOptions,
    salary,
    totalPendapatan,
    totalPengurangan,
    totalDiterima
  });

  const slip = await Slip.create({
    slipNumber,
    institution: normalized.institution,
    periodLabel,
    periodMonth: normalized.month,
    periodYear: normalized.year,
    slipVariant: normalized.slipVariant,
    displayOptions,
    employee: employee._id,
    salary,
    totalPendapatan,
    totalPengurangan,
    totalDiterima,
    pdfPath,
    recipientEmail: employee.email,
    emailStatus: shouldSendEmail ? "pending" : "not-requested"
  });

  if (shouldSendEmail) {
    const result = await sendEmailWithRetry(
      {
        to: employee.email,
        employeeName: employee.name,
        periodLabel,
        pdfPath
      },
      EMAIL_RETRY_LIMIT
    );

    if (result.ok) {
      slip.emailStatus = "sent";
      slip.emailSentAt = new Date();
      slip.emailError = "";
    } else {
      slip.emailStatus = "failed";
      slip.emailError = result.error ? result.error.message : "Email gagal dikirim";
    }

    await slip.save();
  }

  return slip;
}

router.get("/login", (req, res) => {
  if (req.session && req.session.isAuthenticated) {
    return res.redirect("/");
  }

  return res.render("login", {
    error: req.query.error || ""
  });
});

router.post("/login", (req, res) => {
  const username = String(req.body.username || "").trim();
  const password = String(req.body.password || "").trim();

  if (username === LOGIN_USER && password === LOGIN_PASS) {
    req.session.isAuthenticated = true;
    req.session.username = username;
    return res.redirect("/");
  }

  return res.redirect("/login?error=Username%20atau%20password%20salah");
});

router.post("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect("/login");
  });
});

router.use(requireAuth);

router.get("/", async (req, res, next) => {
  try {
    const tab = String(req.query.tab || "gagal").trim().toLowerCase();

    // If user requests the 'gagal' tab, show all slips with emailStatus 'failed' (no limit).
    // Otherwise show the latest 20 slips (default behaviour).
    const baseQuery = {};
    if (tab === "gagal") {
      baseQuery.emailStatus = "failed";
    }

    let slipsQuery = Slip.find(baseQuery).populate("employee").sort({ createdAt: -1 });
    if (tab !== "gagal") {
      slipsQuery = slipsQuery.limit(20);
    }

    const slips = await slipsQuery.lean();

    res.render("index", {
      slips,
      currentTab: tab,
      success: req.query.success || "",
      error: req.query.error || "",
      smtpSuccess: req.query.smtpSuccess || "",
      smtpError: req.query.smtpError || "",
      formatRupiah,
      smtpReady: smtpConfigured(),
      loggedInUser: req.session.username || LOGIN_USER,
      defaultMonth: dayjs().month() + 1,
      defaultYear: dayjs().year()
    });
  } catch (error) {
    next(error);
  }
});

router.post("/slips/create", async (req, res) => {
  try {
    await createSlipRecord(req.body, req.body.sendEmail === "on", 0);

    return res.redirect("/?success=Slip%20berhasil%20dibuat");
  } catch (error) {
    return res.redirect(`/?error=${encodeURIComponent(error.message)}`);
  }
});

router.get("/slips/template.xlsx", (req, res) => {
  const commonHeaders = [
    "Institusi",
    "Bulan",
    "Tahun",
    "Nama Karyawan",
    "Posisi",
    "Tipe Slip",
    "No Rekening",
    "Email Karyawan",
    "Gaji Pokok",
    "Gaji Jasa",
    "Jasa KJS",
    "Gaji Jaga",
    "Tunjangan",
    "Tunjangan Jabaran",
    "Tunjangan Hari Raya",
    "BPJS Ketenagakerjaan (Pendapatan)",
    "BPJS Kesehatan (Pendapatan)",
    "Bonus",
    "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)",
    "BPJS Kesehatan (Potongan Ditanggung Rumah Sakit)",
    "Potongan Lain",
    "Tampilkan Gaji Jasa",
    "Tampilkan Jasa KJS",
    "Tampilkan Gaji Jaga",
    "Tampilkan Tunjangan",
    "Tampilkan Tunjangan Jabaran",
    "Tampilkan Tunjangan Hari Raya",
    "Tampilkan Pemotongan",
    "Tampilkan BPJS Pendapatan",
    "Tampilkan Bonus",
    "Tampilkan Potongan Lain",
    "Kirim Email"
  ];

  const dokterSpesialisHeaders = [
    "Institusi",
    "Bulan",
    "Tahun",
    "Nama Karyawan",
    "Posisi",
    "No Rekening",
    "Email Karyawan",
    "Gaji Pokok",
    "Jasa Medis Pasien Poli",
    "Jasa Medis Pasien Madco",
    "Jasa Medis Rawat Inap",
    "Jasa Medis Tindakan Operasi",
    "Jasa KJS",
    "Tunjangan",
    "Tunjangan Jabaran",
    "Tunjangan Hari Raya",
    "Bonus",
    "BPJS Ketenagakerjaan (Pendapatan)",
    "PPH Pasal 21",
    "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)",
    "Tampilkan Tunjangan",
    "Tampilkan Tunjangan Jabaran",
    "Tampilkan Tunjangan Hari Raya",
    "Tampilkan Pemotongan",
    "Kirim Email"
  ];

  const radiologiHeaders = [
    "Institusi",
    "Bulan",
    "Tahun",
    "Nama Karyawan",
    "Posisi",
    "No Rekening",
    "Email Karyawan",
    "Gaji Pokok",
    "Jasa Pasien Radiologi",
    "Jasa Medis Pasien Poli",
    "Jasa Medis Pasien Madco",
    "Jasa Medis Rawat Inap",
    "Jasa Medis Tindakan Operasi",
    "Jasa KJS",
    "Tunjangan",
    "Tunjangan Jabaran",
    "Tunjangan Hari Raya",
    "Bonus",
    "BPJS Ketenagakerjaan (Pendapatan)",
    "PPH Pasal 21",
    "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)",
    "Tampilkan Tunjangan",
    "Tampilkan Tunjangan Jabaran",
    "Tampilkan Tunjangan Hari Raya",
    "Tampilkan Pemotongan",
    "Kirim Email"
  ];

  const sampleRows = [
    {
      Institusi: "Mulia Rakan Membangun",
      Bulan: dayjs().month() + 1,
      Tahun: dayjs().year(),
      "Nama Karyawan": "Contoh Dokter Umum",
      Posisi: "Dokter Umum",
      "Tipe Slip": "Dokter Umum",
      "No Rekening": "1234567890",
      "Email Karyawan": "dokterumum@email.com",
      "Gaji Pokok": 3500000,
      "Gaji Jasa": 1200000,
      "Jasa KJS": 0,
      "Gaji Jaga": 450000,
      Tunjangan: 350000,
      "Tunjangan Jabaran": 150000,
      "Tunjangan Hari Raya": 0,
      "BPJS Ketenagakerjaan (Pendapatan)": 0,
      "BPJS Kesehatan (Pendapatan)": 0,
      Bonus: 0,
      "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)": 19904,
      "BPJS Kesehatan (Potongan Ditanggung Rumah Sakit)": 0,
      "Potongan Lain": 0,
      "Tampilkan Gaji Jasa": "yes",
      "Tampilkan Jasa KJS": "no",
      "Tampilkan Gaji Jaga": "yes",
      "Tampilkan Tunjangan": "yes",
      "Tampilkan Tunjangan Jabaran": "yes",
      "Tampilkan Tunjangan Hari Raya": "yes",
      "Tampilkan Pemotongan": "yes",
      "Tampilkan BPJS Pendapatan": "yes",
      "Tampilkan Bonus": "yes",
      "Tampilkan Potongan Lain": "no",
      "Kirim Email": "yes"
    },
    {
      Institusi: "Mulia Rakan Membangun",
      Bulan: dayjs().month() + 1,
      Tahun: dayjs().year(),
      "Nama Karyawan": "Contoh Dokter Spesialis",
      Posisi: "Dokter Spesialis",
      "Tipe Slip": "Dokter Spesialis",
      "No Rekening": "1234567891",
      "Email Karyawan": "dokterspesialis@email.com",
      "Gaji Pokok": 5500000,
      "Gaji Jasa": 2200000,
      "Jasa KJS": 0,
      "Gaji Jaga": 0,
      Tunjangan: 500000,
      "Tunjangan Jabaran": 250000,
      "Tunjangan Hari Raya": 0,
      "BPJS Ketenagakerjaan (Pendapatan)": 0,
      "BPJS Kesehatan (Pendapatan)": 0,
      Bonus: 0,
      "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)": 19904,
      "BPJS Kesehatan (Potongan Ditanggung Rumah Sakit)": 0,
      "Potongan Lain": 0,
      "Tampilkan Gaji Jasa": "yes",
      "Tampilkan Jasa KJS": "no",
      "Tampilkan Gaji Jaga": "no",
      "Tampilkan Tunjangan": "yes",
      "Tampilkan Tunjangan Jabaran": "yes",
      "Tampilkan Tunjangan Hari Raya": "yes",
      "Tampilkan Pemotongan": "yes",
      "Tampilkan BPJS Pendapatan": "yes",
      "Tampilkan Bonus": "yes",
      "Tampilkan Potongan Lain": "yes",
      "Kirim Email": "yes"
    },
    {
      Institusi: "Mulia Rakan Membangun",
      Bulan: dayjs().month() + 1,
      Tahun: dayjs().year(),
      "Nama Karyawan": "Contoh Karyawan",
      Posisi: "Karyawan",
      "Tipe Slip": "Karyawan",
      "No Rekening": "1234567892",
      "Email Karyawan": "karyawan@email.com",
      "Gaji Pokok": 2800000,
      "Gaji Jasa": 0,
      "Jasa KJS": 0,
      "Gaji Jaga": 0,
      Tunjangan: 250000,
      "Tunjangan Jabaran": 0,
      "Tunjangan Hari Raya": 0,
      "BPJS Ketenagakerjaan (Pendapatan)": 0,
      "BPJS Kesehatan (Pendapatan)": 0,
      Bonus: 0,
      "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)": 19904,
      "BPJS Kesehatan (Potongan Ditanggung Rumah Sakit)": 0,
      "Potongan Lain": 0,
      "Tampilkan Gaji Jasa": "no",
      "Tampilkan Jasa KJS": "no",
      "Tampilkan Gaji Jaga": "no",
      "Tampilkan Tunjangan": "yes",
      "Tampilkan Tunjangan Jabaran": "no",
      "Tampilkan Tunjangan Hari Raya": "no",
      "Tampilkan Pemotongan": "yes",
      "Tampilkan BPJS Pendapatan": "yes",
      "Tampilkan Bonus": "yes",
      "Tampilkan Potongan Lain": "no",
      "Kirim Email": "yes"
    }
  ];

  const sampleKaryawan = [sampleRows[2]];
  const sampleDokterUmum = [sampleRows[0]];

  const sampleDokterSpesialis = [
    {
      Institusi: "Mulia Rakan Membangun",
      Bulan: dayjs().month() + 1,
      Tahun: dayjs().year(),
      "Nama Karyawan": "Contoh Dokter Spesialis",
      Posisi: "Dokter Spesialis",
      "No Rekening": "9988776655",
      "Email Karyawan": "spesialis@email.com",
      "Gaji Pokok": 0,
      "Jasa Medis Pasien Poli": 670200,
      "Jasa Medis Pasien Madco": 288000,
      "Jasa Medis Rawat Inap": 0,
      "Jasa Medis Tindakan Operasi": 0,
      "Jasa KJS": 0,
      Tunjangan: 500000,
      "Tunjangan Jabaran": 200000,
      "Tunjangan Hari Raya": 0,
      Bonus: 0,
      "BPJS Ketenagakerjaan (Pendapatan)": 0,
      "PPH Pasal 21": 0,
      "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)": 0,
      "Tampilkan Tunjangan": "yes",
      "Tampilkan Tunjangan Jabaran": "yes",
      "Tampilkan Tunjangan Hari Raya": "yes",
      "Tampilkan Pemotongan": "yes",
      "Kirim Email": "yes"
    }
  ];

  const sampleRadiologi = [
    {
      Institusi: "Mulia Rakan Membangun",
      Bulan: dayjs().month() + 1,
      Tahun: dayjs().year(),
      "Nama Karyawan": "Contoh Spesialis Radiologi",
      Posisi: "Dokter Spesialis",
      "No Rekening": "7766554433",
      "Email Karyawan": "radiologi@email.com",
      "Gaji Pokok": 0,
      "Jasa Pasien Radiologi": 550000,
      "Jasa Medis Pasien Poli": 120000,
      "Jasa Medis Pasien Madco": 220000,
      "Jasa Medis Rawat Inap": 0,
      "Jasa Medis Tindakan Operasi": 0,
      "Jasa KJS": 0,
      Tunjangan: 300000,
      "Tunjangan Jabaran": 100000,
      "Tunjangan Hari Raya": 0,
      Bonus: 0,
      "BPJS Ketenagakerjaan (Pendapatan)": 0,
      "PPH Pasal 21": 0,
      "BPJS Ketenagakerjaan (Potongan Ditanggung Rumah Sakit)": 0,
      "Tampilkan Tunjangan": "yes",
      "Tampilkan Tunjangan Jabaran": "yes",
      "Tampilkan Tunjangan Hari Raya": "yes",
      "Tampilkan Pemotongan": "yes",
      "Kirim Email": "yes"
    }
  ];

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(sampleRows, { header: commonHeaders });
  worksheet["!cols"] = [
    { wch: 24 },
    { wch: 8 },
    { wch: 8 },
    { wch: 26 },
    { wch: 12 },
    { wch: 20 },
    { wch: 14 },
    { wch: 16 },
    { wch: 28 },
    { wch: 14 },
    { wch: 12 },
    { wch: 20 },
    { wch: 12 },
    { wch: 34 },
    { wch: 28 },
    { wch: 10 },
    { wch: 32 },
    { wch: 27 },
    { wch: 14 },
    { wch: 20 },
    { wch: 20 },
    { wch: 20 },
    { wch: 26 },
    { wch: 18 },
    { wch: 24 },
    { wch: 12 }
  ];
  XLSX.utils.book_append_sheet(workbook, worksheet, "template_semua");

  const karyawanSheet = XLSX.utils.json_to_sheet(sampleKaryawan, { header: commonHeaders });
  XLSX.utils.book_append_sheet(workbook, karyawanSheet, TEMPLATE_SHEETS.KARYAWAN);

  const dokumSheet = XLSX.utils.json_to_sheet(sampleDokterUmum, { header: commonHeaders });
  XLSX.utils.book_append_sheet(workbook, dokumSheet, TEMPLATE_SHEETS.DOKTER_UMUM);

  const dSpSheet = XLSX.utils.json_to_sheet(sampleDokterSpesialis, { header: dokterSpesialisHeaders });
  XLSX.utils.book_append_sheet(workbook, dSpSheet, TEMPLATE_SHEETS.DOKTER_SPESIALIS);

  const radioSheet = XLSX.utils.json_to_sheet(sampleRadiologi, { header: radiologiHeaders });
  XLSX.utils.book_append_sheet(workbook, radioSheet, TEMPLATE_SHEETS.RADIOLOGI);

  const referensiRows = [
    { tipeSlipValid: "Dokter Umum" },
    { tipeSlipValid: "Dokter Spesialis" },
    { tipeSlipValid: "Karyawan" }
  ];
  const referensiSheet = XLSX.utils.json_to_sheet(referensiRows);
  XLSX.utils.book_append_sheet(workbook, referensiSheet, "referensi_tipe_slip");

  const panduanSheet = XLSX.utils.json_to_sheet([
    {
      catatan:
        "Isi salah satu sheet template sesuai kebutuhan. Posisi bebas diisi apa saja, sistem akan menampilkan sesuai input.",
      nilai_boolean: "Gunakan yes/no untuk kolom Kirim Email atau Tampilkan ...",
      tipe_slip: "Tipe Slip wajib salah satu: Dokter Umum, Dokter Spesialis, atau Karyawan",
      kolom_baru:
        "Gunakan kolom Tunjangan, Tunjangan Jabaran, Tunjangan Hari Raya, dan Tampilkan Pemotongan untuk komponen/toggle tambahan"
    }
  ]);
  XLSX.utils.book_append_sheet(workbook, panduanSheet, "panduan");

  const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });

  res.setHeader(
    "Content-Disposition",
    `attachment; filename="template-slip-batch-${dayjs().format("YYYYMMDD")}.xlsx"`
  );
  res.type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  return res.send(buffer);
});

router.post("/slips/batch-upload", upload.single("batchFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.redirect("/?error=File%20Excel%20wajib%20diupload");
    }

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const dataSheets = workbook.SheetNames.filter(
      (name) => !["referensi_posisi", "referensi_tipe_slip", "panduan"].includes(String(name).toLowerCase())
    );

    if (!dataSheets.length) {
      return res.redirect("/?error=Sheet%20Excel%20tidak%20ditemukan");
    }

    const rowsWithSheet = [];
    dataSheets.forEach((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
      rows.forEach((row) => rowsWithSheet.push({ ...row, __sheetName: sheetName }));
    });

    if (!rowsWithSheet.length) {
      return res.redirect("/?error=Isi%20Excel%20kosong");
    }

    const sendAllEmail = req.body.sendAllEmail === "on";
    let created = 0;
    let sent = 0;
    let failed = 0;
    const errors = [];

    for (let index = 0; index < rowsWithSheet.length; index += 1) {
      const row = rowsWithSheet[index];
      try {
        const sendFlag = sendAllEmail || parseBoolean(pickValue(row, ["sendEmail", "Kirim Email", "kirimEmail"], ""));
        const slip = await createSlipRecord(row, sendFlag, index + 1);
        created += 1;
        if (slip.emailStatus === "sent") {
          sent += 1;
        }
        if (slip.emailStatus === "failed") {
          failed += 1;
        }

        if (sendFlag && BATCH_EMAIL_DELAY_MS > 0 && index < rowsWithSheet.length - 1) {
          await sleep(BATCH_EMAIL_DELAY_MS);
        }
      } catch (error) {
        errors.push(`Baris ${index + 2}: ${error.message}`);
      }
    }

    if (!created) {
      return res.redirect(`/?error=${encodeURIComponent(`Batch gagal. ${errors[0] || "Tidak ada data valid"}`)}`);
    }

    const summary = `Batch selesai. Berhasil: ${created}, Email terkirim: ${sent}, Email gagal: ${failed}, Baris error: ${errors.length}`;

    if (errors.length) {
      return res.redirect(
        `/?success=${encodeURIComponent(summary)}&error=${encodeURIComponent(errors.slice(0, 2).join(" | "))}`
      );
    }

    return res.redirect(`/?success=${encodeURIComponent(summary)}`);
  } catch (error) {
    return res.redirect(`/?error=${encodeURIComponent(error.message)}`);
  }
});

router.post("/smtp/test", async (req, res) => {
  try {
    await verifySmtpConnection();
    return res.redirect("/?smtpSuccess=Koneksi%20SMTP%20berhasil.%20Sistem%20siap%20kirim%20email.");
  } catch (error) {
    return res.redirect(`/?smtpError=${encodeURIComponent(error.message)}`);
  }
});

router.get("/slips/:id/download", async (req, res) => {
  try {
    const slip = await Slip.findById(req.params.id).lean();

    if (!slip) {
      return res.status(404).send("Slip tidak ditemukan");
    }

    return res.download(path.resolve(slip.pdfPath));
  } catch (error) {
    return res.status(500).send(error.message);
  }
});

module.exports = router;
