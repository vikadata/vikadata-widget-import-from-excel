export function parseImportNumber(value: unknown): number | null {
  if (value == null) return null;
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new Error("Invalid numeric value");
    return value;
  }
  if (typeof value !== "string") throw new Error("Invalid numeric value");

  let text = value.trim();
  if (!text) return null;

  const accountingNegative = text.startsWith("(") && text.endsWith(")");
  if (accountingNegative) text = text.slice(1, -1).trim();
  text = text
    .replace(/\uFF0C/g, ",")
    .replace(/[\u00A0\u202F]/g, " ")
    .replace(/^([+-]?)\s*[$\u00A3\u00A5\u20AC\u20A9\u20B9\uFFE5]\s*/, "$1")
    .replace(/\s*[$\u00A3\u00A5\u20AC\u20A9\u20B9\uFFE5]$/, "")
    .trim();

  // Validate thousands grouping before removing separators to avoid changing malformed values.
  const numericText = /^[+-]?(?:(?:\d+|\d{1,3}(?:,\d{3})+|\d{1,3}(?: \d{3})+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
  if (!numericText.test(text) || (accountingNegative && /^[+-]/.test(text))) {
    throw new Error("Invalid numeric value");
  }

  const number = Number(text.replace(/[, ]/g, ""));
  if (!Number.isFinite(number)) throw new Error("Invalid numeric value");
  return accountingNegative ? -number : number;
}
