/* auth.js — API key authentication for Gorelo Public API.
   Replaces the previous MSAL B2C implementation. */

import { getSetting } from '../settings.js';

/**
 * Returns the stored API key, throwing a user-friendly error if absent.
 */
export function getApiKey() {
  const key = getSetting('apiKey');
  if (!key || !key.trim()) {
    throw new Error('No Gorelo API key configured. Open Settings (⚙) to add your key.');
  }
  return key.trim();
}

/**
 * Returns true when an API key has been saved (no network call needed).
 */
export function isAuthenticated() {
  const key = getSetting('apiKey');
  return !!(key && key.trim().length > 0);
}

/**
 * Validates an API key against the live Gorelo API.
 * Uses GET /v1/clients?PageSize=1 — a lightweight authenticated endpoint.
 * Returns { valid: true } or { valid: false, error: string }.
 */
export async function validateApiKey(key) {
  if (!key || !key.trim()) {
    return { valid: false, error: 'API key cannot be empty.' };
  }
  try {
    const response = await fetch('https://api.usw.gorelo.io/v1/clients?PageSize=1', {
      headers: {
        'X-API-Key': key.trim(),
        'Accept': 'application/json',
      },
    });
    if (response.ok) {
      return { valid: true };
    } else if (response.status === 401 || response.status === 403) {
      return { valid: false, error: 'Invalid API key — please check and try again.' };
    } else {
      return { valid: false, error: `Gorelo API returned ${response.status}. Try again shortly.` };
    }
  } catch (err) {
    return { valid: false, error: 'Could not reach Gorelo API: ' + err.message };
  }
}
