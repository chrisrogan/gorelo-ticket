/**
 * Gorelo Ticket - API Client
 * All calls to gw.usw.gorelo.tech use the same headers and Bearer token.
 *
 * NOTE: This uses Gorelo's internal API gateway, reverse-engineered from
 * browser traffic. Gorelo has not officially documented these endpoints.
 * They may change without notice. See README for details.
 */

import { getToken } from './auth/auth.js';

const BASE_URL = 'https://gw.usw.gorelo.tech';
const API_VERSION = '1.0';

/**
 * Core fetch wrapper with Gorelo auth headers.
 */
async function gorelofetch(path, options = {}) {
  const token = await getToken();

  const url = `${BASE_URL}${path}${path.includes('?') ? '&' : '?'}api-version=${API_VERSION}&AppType=Web`;

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'apptype': 'Web',
      'iscustomdomain': '1',
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Gorelo API error ${response.status}: ${text}`);
  }

  return response.json();
}

// ─── Lookup endpoints ────────────────────────────────────────────────────────

export async function getClients() {
  // Returns clientEntity array from Gorelo
  return gorelofetch('/client/v1/client/list');
}

export async function getContacts(clientId = null) {
  const path = clientId
    ? `/contact/v1/contact/list?ClientId=${clientId}`
    : '/contact/v1/contact/list';
  return gorelofetch(path);
}

export async function getTechnicianGroups() {
  return gorelofetch('/ticket/v1/techniciangroup/list');
}

export async function getTechnicians() {
  return gorelofetch('/ticket/v1/technician/dropdown');
}

export async function getTags() {
  return gorelofetch('/ticket/v1/tag/list');
}

export async function getAssets(clientId = null) {
  const path = clientId
    ? `/asset/v1/asset/list?ClientId=${clientId}`
    : '/asset/v1/asset/list';
  return gorelofetch(path);
}

export async function getLocations() {
  return gorelofetch('/ticket/v1/location/list');
}

export async function getPriorities() {
  return gorelofetch('/ticket/v1/ticket/priority/list');
}

export async function getSources() {
  return gorelofetch('/ticket/v1/ticket/source/list');
}

export async function getEmailConfigurations() {
  return gorelofetch('/ticket/v1/emailconfiguration/list');
}

// ─── Ticket creation ─────────────────────────────────────────────────────────

/**
 * Create a ticket in Gorelo.
 *
 * @param {object} params
 * @param {string}   params.title              - Ticket title/subject
 * @param {string}   params.description        - HTML body (ticket description)
 * @param {string}   params.markdownComment    - Plain text initial comment
 * @param {number}   params.clientId
 * @param {string}   params.clientName
 * @param {number}   params.contactId
 * @param {string}   params.outboundContactEmail - Original sender email address
 * @param {number[]} params.technicianGroupIds
 * @param {number}   params.assignedToId
 * @param {number}   params.priorityId          - 1=High, 2=Normal, 3=Low, 4=Urgent
 * @param {number}   params.sourceId            - 1=Web, 2=Email, 3=Phone etc.
 * @param {number}   params.emailConfigurationId
 * @param {number}   params.ticketLocation
 * @param {number[]} params.tagIds
 * @param {string[]} params.deviceIds           - Asset GUIDs
 * @param {boolean}  params.sendTicketCreatedEmail
 * @param {boolean}  params.onSite
 */
export async function createTicket(params) {
  const payload = {
    // Core fields
    TicketTitle: params.title,
    TicketDescription: params.description || `<p>${params.markdownComment || ''}</p>`,
    markdownComment: params.markdownComment || '',

    // Client & contact
    ClientId: params.clientId,
    ClientName: params.clientName,
    ContactId: params.contactId || 0,
    outboundContactEmail: params.outboundContactEmail || '',

    // Assignment
    technicianGroupIds: params.technicianGroupIds || [],
    AssignedToId: params.assignedToId || 0,

    // Ticket metadata
    TicketPriorityId: params.priorityId || 2,       // Default: Normal
    TicketSeverityId: params.severityId || 3,        // Default: Medium
    TicketSourceId: params.sourceId || 2,            // Default: Email
    TicketStatusId: 1,                               // Open
    TypeId: null,

    // Email & location
    emailConfigurationId: params.emailConfigurationId,
    ticketLocation: params.ticketLocation || 0,

    // Tags & assets
    NewTicketTags: params.tagIds || [],
    OldTicketTags: params.tagIds || [],
    TicketDeviceIds: params.deviceIds || [],
    TicketCustomAssetIds: [],
    TicketNodeChecks: [],

    // Notification options
    sendTicketCreatedEmail: params.sendTicketCreatedEmail !== false,
    doNotEmailCheck: false,
    sendSelfMentionNotification: false,
    isFromWeb: true,
    isTicketCreatedThroughAlert: false,
    onSite: params.onSite || false,

    // Relations
    mentions: [],
    mentionedTechnicianGroupIds: [],
    secondaryContacts: [],
    secondaryTechnicians: [],
    watchers: [],
    attachments: [],
    deviceId: '',
  };

  return gorelofetch('/ticket/v1/ticket/add', {
    method: 'POST',
    body: JSON.stringify(payload),
  });
}
