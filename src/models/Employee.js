const mongoose = require("mongoose");

const employeeSchema = new mongoose.Schema(
  {
    name: { type: String, required: true, trim: true },
    position: { type: String, required: true, trim: true },
    accountNumber: { type: String, trim: true, default: "-" },
    email: { type: String, required: true, trim: true, lowercase: true, unique: true }
  },
  { timestamps: true }
);

module.exports = mongoose.model("Employee", employeeSchema);
