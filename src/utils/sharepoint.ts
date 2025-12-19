/**
 * SharePoint URL utilities for Microsoft Graph API.
 */

/**
 * Encode a SharePoint URL for use with the /shares endpoint.
 * Algorithm: "u!" + base64url(url)
 *
 * @see https://learn.microsoft.com/en-us/graph/api/shares-get
 */
export function encodeSharePointUrl(sharepointUrl: string): string {
  const base64Value = Buffer.from(sharepointUrl).toString('base64');
  return `u!${base64Value
    .replace(/=+$/, '') // Remove padding
    .replace(/\//g, '_') // Replace / with _
    .replace(/\+/g, '-')}`; // Replace + with -
}
