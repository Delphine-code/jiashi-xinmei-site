export function withBase(path: string) {
  const base = import.meta.env.BASE_URL || '/';
  if (!path) return base;
  if (/^(https?:|mailto:|#|data:)/.test(path)) return path;
  const normalizedBase = base.replace(/\/$/, '');
  if (normalizedBase && (path === normalizedBase || path.startsWith(normalizedBase + '/'))) return path;
  return base.replace(/\/$/, '') + '/' + path.replace(/^\//, '');
}
