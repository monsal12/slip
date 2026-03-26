function toNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatRupiah(value) {
  return new Intl.NumberFormat("id-ID").format(toNumber(value));
}

module.exports = { toNumber, formatRupiah };
