/* gorelo-api.js — All calls to the Gorelo Public API (api.usw.gorelo.io).
   Authentication: X-API-Key header using the key stored in RoamingSettings.

   Real API shape (confirmed from live responses):
   - Response wrapper: { StatusCode, IsSuccess, Data: [...], DataContext: { Pagination: { NextCursor, HasMore } } }
   - Pagination: cursor-based using ?Cursor=<token>&PageSize=<n>  (no PageNumber)
   - Clients:   Data[].{ Id, Name, Status, Domains[] }
   - Contacts:  Data[].{ Id, FirstName, LastName, ClientId, PrimaryEmail }
   - Users:     Data[].{ Id, FirstName, LastName, Email }
   - Groups:    /v1/organization/groups  (no pagination params)
   - Tags:      /v1/tickets/tags  (requires Tickets read scope)
   - Tickets:   POST /v1/tickets
*/

import { getApiKey } from './auth/auth.js';

const BASE = 'https://gorelo-proxy.chris-rogan.workers.dev';

function headers() {
  return {
    'X-API-Key': getApiKey(),
    'Content-Type': 'application/json',
    'Accept': 'application/json',
  };
}

async function apiGet(path) {
  const res = await fetch(BASE + path, { headers: headers() });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Gorelo API ${res.status}: ${text || res.statusText}`);
  }
  const json = await res.json();
  if (!json.IsSuccess) {
    const msg = json.Notifications?.[0]?.Message || `Status ${json.StatusCode}`;
    throw new Error(`Gorelo: ${msg}`);
  }
  return json;
}

async function apiPost(path, body) {
  const res = await fetch(BASE + path, {
    method: 'POST',
    headers: headers(),
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Gorelo API ${res.status}: ${text || res.statusText}`);
  }
  const ct = res.headers.get('content-type') || '';
  if (res.status === 204 || !ct.includes('application/json')) return null;
  const json = await res.json();
  if (json.IsSuccess === false) {
    const msg = json.Notifications?.[0]?.Message || `Status ${json.StatusCode}`;
    throw new Error(`Gorelo: ${msg}`);
  }
  return json.Data ?? json;
}

/** Unwrap Data array from standard response */
function unwrap(json) {
  return Array.isArray(json.Data) ? json.Data : [];
}

// ─── Lookups ─────────────────────────────────────────────────────────────────

export async function getClients(search = '') {
  const q = search ? `&Search=${encodeURIComponent(search)}` : '';
  const json = await apiGet(`/v1/clients?PageSize=200${q}`);
  // Map to consistent shape: { id, name }
  return unwrap(json).map(c => ({ id: c.Id, name: c.Name }));
}

export async function getContacts(clientId) {
  const q = clientId ? `?ClientId=${clientId}&PageSize=200` : '?PageSize=200';
  const json = await apiGet(`/v1/contacts${q}`);
  // Map to consistent shape: { id, name, clientId, email }
  return unwrap(json).map(c => ({
    id: c.Id,
    name: [c.FirstName, c.LastName].filter(Boolean).join(' ') || c.PrimaryEmail || `Contact #${c.Id}`,
    clientId: c.ClientId,
    email: c.PrimaryEmail || '',
  }));
}

export async function getTechnicianGroups() {
  // /v1/organization/groups does not accept PageSize/PageNumber
  const json = await apiGet('/v1/organization/groups');
  return unwrap(json).map(g => ({ id: g.Id, name: g.Name }));
}

export async function getTechnicians(groupId) {
  // Users endpoint — filter by group if supported, otherwise return all
  const q = groupId ? `?GroupId=${groupId}&PageSize=200` : '?PageSize=200';
  try {
    const json = await apiGet(`/v1/organization/users${q}`);
    return unwrap(json).map(u => ({
      id: u.Id,
      name: [u.FirstName, u.LastName].filter(Boolean).join(' ') || u.Email || `User #${u.Id}`,
      email: u.Email || '',
    }));
  } catch {
    // GroupId filter may not be supported — fall back to all users
    const json = await apiGet('/v1/organization/users?PageSize=200');
    return unwrap(json).map(u => ({
      id: u.Id,
      name: [u.FirstName, u.LastName].filter(Boolean).join(' ') || u.Email || `User #${u.Id}`,
      email: u.Email || '',
    }));
  }
}

export async function getTags() {
  // Requires 'Tickets' read scope on the API key.
  // Returns empty array gracefully if the key lacks that permission.
  try {
    const json = await apiGet('/v1/tickets/tags?PageSize=200');
    return unwrap(json).map(t => ({ id: t.Id, name: t.Name || t.TagName || '' }));
  } catch {
    return [];
  }
}

export async function getEmailConfigurations() {
  // Not present in the Public API — return empty array.
  return [];
}

// Static lookups — values confirmed from Gorelo UI
export const PRIORITIES = [
  { id: 1, name: 'High' },
  { id: 2, name: 'Normal' },
  { id: 3, name: 'Low' },
  { id: 4, name: 'Urgent' },
  { id: 5, name: 'No Priority' },
];

export const SOURCES = [
  { id: 1, name: 'Web' },
  { id: 2, name: 'Email' },
  { id: 3, name: 'Phone' },
  { id: 4, name: 'Chat' },
  { id: 5, name: 'Alert' },
  { id: 6, name: 'API' },
];

// ─── Ticket creation ──────────────────────────────────────────────────────────

/**
 * Creates a ticket via POST /v1/tickets.
 *
 * @param {object} opts
 * @param {string}   opts.title
 * @param {string}   opts.description        Plain text / HTML body
 * @param {string}   opts.markdownComment    Markdown version of body
 * @param {number}   [opts.clientId]
 * @param {string}   [opts.clientName]
 * @param {number}   [opts.contactId]
 * @param {string}   [opts.outboundContactEmail]
 * @param {number[]} [opts.technicianGroupIds]
 * @param {number}   [opts.assignedToId]
 * @param {number}   [opts.priorityId]       1=High 2=Normal 3=Low 4=Urgent
 * @param {number}   [opts.sourceId]         2=Email
 * @param {string[]} [opts.newTags]          Tag names to create
 * @param {string[]} [opts.existingTagIds]   Existing tag IDs
 * @param {number}   [opts.emailConfigurationId]
 * @param {boolean}  [opts.sendNotification]
 */
export async function createTicket(opts) {
  const payload = {
    TicketTitle: opts.title,
    TicketDescription: opts.description || '',
    markdownComment: opts.markdownComment || opts.description || '',
    ClientId: opts.clientId || 0,
    ClientName: opts.clientName || '',
    ContactId: opts.contactId || 0,
    outboundContactEmail: opts.outboundContactEmail || '',
    technicianGroupIds: opts.technicianGroupIds || [],
    AssignedToId: opts.assignedToId || 0,
    TicketPriorityId: opts.priorityId || 2,
    TicketSeverityId: null,
    TicketSourceId: opts.sourceId || 2,
    TicketStatusId: 1,
    emailConfigurationId: opts.emailConfigurationId || null,
    ticketLocation: opts.ticketLocation || 0,
    NewTicketTags: opts.newTags || [],
    OldTicketTags: opts.existingTagIds || [],
    TicketDeviceIds: [],
    sendTicketCreatedEmail: opts.sendNotification !== false,
    isFromWeb: true,
  };

  return apiPost('/v1/tickets', payload);
}
