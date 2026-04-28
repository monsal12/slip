const mongoose = require("mongoose");

const salaryComponentSchema = new mongoose.Schema(
  {
    gajiPokok: { type: Number, default: 0 },
    gajiJasa: { type: Number, default: 0 },
    jasaKjs: { type: Number, default: 0 },
    gajiJaga: { type: Number, default: 0 },
    tunjangan: { type: Number, default: 0 },
    tunjanganJabaran: { type: Number, default: 0 },
    tunjanganHariRaya: { type: Number, default: 0 },
    bpjsKetenagakerjaanPendapatan: { type: Number, default: 0 },
    bpjsKesehatanPendapatan: { type: Number, default: 0 },
    bonus: { type: Number, default: 0 },
    bpjsKetenagakerjaanPotongan: { type: Number, default: 0 },
    bpjsKesehatanPotongan: { type: Number, default: 0 },
    potonganLain: { type: Number, default: 0 }
  },
  { _id: false }
);

const displayOptionsSchema = new mongoose.Schema(
  {
    showGajiJasa: { type: Boolean, default: true },
    showJasaKjs: { type: Boolean, default: true },
    showGajiJaga: { type: Boolean, default: false },
    showTunjangan: { type: Boolean, default: true },
    showTunjanganJabaran: { type: Boolean, default: true },
    showTunjanganHariRaya: { type: Boolean, default: true },
    showPengurangan: { type: Boolean, default: true },
    showBpjsPendapatan: { type: Boolean, default: true },
    showBonus: { type: Boolean, default: true },
    showPotonganLain: { type: Boolean, default: true }
  },
  { _id: false }
);

const slipSchema = new mongoose.Schema(
  {
    slipNumber: { type: String, required: true, unique: true },
    institution: { type: String, required: true, default: "Mulia Rakan Membangun" },
    periodLabel: { type: String, required: true },
    periodMonth: { type: Number, required: true },
    periodYear: { type: Number, required: true },
    slipVariant: {
      type: String,
      enum: ["dokter_umum", "dokter_spesialis", "karyawan", "otomatis"],
      default: "dokter_spesialis"
    },
    displayOptions: { type: displayOptionsSchema, default: () => ({}) },
    employee: { type: mongoose.Schema.Types.ObjectId, ref: "Employee", required: true },
    salary: { type: salaryComponentSchema, default: () => ({}) },
    totalPendapatan: { type: Number, required: true },
    totalPengurangan: { type: Number, required: true },
    totalDiterima: { type: Number, required: true },
    pdfPath: { type: String, required: true },
    recipientEmail: { type: String, required: true },
    emailStatus: {
      type: String,
      enum: ["pending", "sent", "failed", "not-requested"],
      default: "pending"
    },
    emailError: { type: String, default: "" },
    emailSentAt: { type: Date }
  },
  { timestamps: true }
);

module.exports = mongoose.model("Slip", slipSchema);
