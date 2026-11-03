(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.CocinaProfileStore = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  const key = 'cocina-user-profile-v1';
  const defaultProfile = () => ({
    schemaVersion: 1,
    profileName: 'Mi configuración',
    mode: 'advanced',
    periods: { monthly: true, weekly: false, daily: true },
    rules: [
      { id: 'carbonara-friday-monthly', type: 'frequency', period: 'monthly', target: { kind: 'dish', id: 'MEN-A0E0A499', name: 'Carbonara' }, operator: 'exact', count: 2, weekdays: [5], meal: 'Comida', gapWeeks: 1, strength: 'mandatory', enabled: true, applied: true, origin: 'personal' },
      { id: 'no-same-family-dinner', type: 'no-same-family-dinner', period: 'daily', scope: 'dinner', strength: 'mandatory', enabled: true, origin: 'personal' }
    ],
    recipes: []
  });
  function load(storage = globalThis.localStorage) {
    try {
      const saved = JSON.parse(storage.getItem(key) || 'null');
      if (!saved || typeof saved !== 'object') return defaultProfile();
      const defaults = defaultProfile();
      const rules = Array.isArray(saved.rules) ? [...saved.rules] : [...defaults.rules];
      const carbonara = rules.find(rule => rule.type === 'frequency' && (rule.target?.id === 'MEN-A0E0A499' || String(rule.target?.name || '').toLocaleLowerCase('es') === 'carbonara'));
      if (!carbonara) rules.unshift(defaults.rules[0]);
      else Object.assign(carbonara, defaults.rules[0], carbonara, { applied: true });
      const periods = { ...defaults.periods, ...(saved.periods || {}), monthly: saved.periods?.monthly === false ? false : true };
      return { ...defaults, ...saved, periods, rules, recipes: Array.isArray(saved.recipes) ? saved.recipes : [] };
    } catch { return defaultProfile(); }
  }
  function save(profile, storage = globalThis.localStorage) {
    const clean = { ...profile, schemaVersion: 1, periods: profile.periods || defaultProfile().periods, rules: profile.rules || [], recipes: profile.recipes || [] };
    storage.setItem(key, JSON.stringify(clean));
    return clean;
  }
  function createId() { return `USR-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`; }
  return { key, defaultProfile, load, save, createId };
});
