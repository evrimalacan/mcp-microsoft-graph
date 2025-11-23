/**
 * Utilities for working with OData queries
 */

/**
 * Escapes a string for use in OData filter expressions.
 * Single quotes are escaped by doubling them (' → '').
 *
 * @param value - The string to escape
 * @returns The escaped string safe for OData filters
 *
 * @example
 * // "Let's Talk" → "Let''s Talk"
 * const escaped = escapeODataString("Let's Talk");
 * const filter = `contains(topic, '${escaped}')`;
 */
export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
