(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.CocinaTicketTools = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  const normalize = value => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase('es').replace(/[^\p{L}\p{N}]+/gu, ' ').trim();
  const ignored = /^(?:total|subtotal|iva|impuestos?|cambio|efectivo|tarjeta|visa|mastercard|pago|gracias|vuelva|cliente|fecha|hora|ticket|factura|base imponible|descuento|ahorro|cashback|numero|nif|cif|www|http)/i;
  const money = /(?:\d+[,.]\d{2}\s*(?:€|eur)?\s*$|€\s*\d)/i;
  function parseReceiptText(text) {
    const seen = new Set();
    return String(text || '').split(/\r?\n/).map(line => line.replace(/[|]/g, ' ').replace(/\s+/g, ' ').trim())
      .map(line => line.replace(/\s+(?:\d+[,.]\d{2}\s*€?|€\s*\d+[,.]\d{2})\s*$/i, '').replace(/\s+\d+\s*(?:ud|uds|kg|g|gr|lt|l)\b.*$/i, '').trim())
      .filter(line => line.length >= 3 && line.length <= 90 && !ignored.test(line) && !money.test(line) && !/^\d[\d\s./:-]*$/.test(line))
      .map(line => line.replace(/^[\d.,xX*\-\s]+/, '').trim())
      .filter(line => line.length >= 3 && !ignored.test(line))
      .filter(line => { const key = normalize(line); if (!key || seen.has(key)) return false; seen.add(key); return true; });
  }
  function ingredientLines(recipe, curatedModels = []) {
    const model = curatedModels.find(item => item.name === recipe.name);
    if (model?.blocks) {
      const lines = model.blocks.filter(block => block.type === 'ingredients').flatMap(block => block.items || []).map(item => typeof item === 'string' ? item : item.text);
      if (lines.length) return { items: [...new Set(lines)], source: 'estructurada' };
    }
    const lines = String(recipe.body || '').split(/\r?\n/).map(line => line.trim()).filter(Boolean);
    const start = lines.findIndex(line => /^ingredientes?(?:\s|:|$)/i.test(line));
    if (start < 0) return { items: [], source: 'pendiente' };
    const end = lines.findIndex((line, i) => i > start && /^(?:preparad[oa]|preparaci[oó]n|elaboraci[oó]n|procedimiento|pasos?|nota|recomendaciones?)\b/i.test(line));
    const candidates = lines.slice(start + 1, end < 0 ? undefined : end).filter(line => !/^(?:ingredientes?|para la|para el|salsa|relleno|opcional|base|guarnici[oó]n)\b/i.test(line) && line.length < 100);
    return candidates.length ? { items: [...new Set(candidates)], source: 'texto-importado' } : { items: [], source: 'pendiente' };
  }
  function matchesProduct(product, ingredient) {
    const a = normalize(product), b = normalize(ingredient);
    if (!a || !b) return false;
    if (a === b) return true;
    const wordsA = new Set(a.split(' ')), wordsB = b.split(' ');
    const shared = [...wordsA].filter(word => word.length >= 3 && wordsB.has(word));
    return shared.length > 0 && [...wordsA].every(word => word.length < 3 || wordsB.has(word)) && [...wordsB].every(word => word.length < 3 || wordsA.has(word));
  }
  function compareRecipe(recipe, available, curatedModels = []) {
    const source = ingredientLines(recipe, curatedModels);
    if (!source.items.length) return { ...source, matched: [], missing: [], coverage: 'sin-datos' };
    const matched = source.items.filter(item => available.some(product => matchesProduct(product, item)));
    return { ...source, matched, missing: source.items.filter(item => !matched.includes(item)), coverage: 'parcial' };
  }
  return { normalize, parseReceiptText, ingredientLines, matchesProduct, compareRecipe };
});
