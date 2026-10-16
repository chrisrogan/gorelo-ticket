/* commands.js — Quick Create ribbon button handler.
   Creates a ticket with defaults and minimal prompting. */

import { isAuthenticated, getApiKey } from '../auth/auth.js';
import { createTicket } from '../gorelo-api.js';
import { getSetting } from '../settings.js';

Office.onReady(() => {
  // Register the function called by the manifest ExecuteFunction action
  if (typeof Office.actions !== 'undefined') {
    Office.actions.associate('quickCreateTicket', quickCreateTicket);
  }
});

/**
 * One-click Quick Create: reads the current email and fires a ticket
 * with the user's saved defaults. Shows a brief notification with the result.
 */
async function quickCreateTicket(event) {
  // Guard: API key must be configured
  if (!isAuthenticated()) {
    showNotification(
      'Gorelo — Setup needed',
      'Open the Gorelo Ticket task pane (⚙ button) to add your API key first.',
      'errorMessage'
    );
    event.completed();
    return;
  }

  const item = Office.context.mailbox.item;

  try {
    // Read email fields
    const subject = item.subject || '(No subject)';
    const from = item.from || {};
    const fromName = from.displayName || from.emailAddress || '';
    const fromEmail = from.emailAddress || '';

    // Get body text (plain text is fine for the description)
    const body = await getBodyText(item);

    // Build description with original sender info preserved
    const description = buildDescription(fromName, fromEmail, body);

    const result = await createTicket({
      title: subject,
      description: description,
      markdownComment: description,
      outboundContactEmail: fromEmail,
      priorityId: getSetting('defaultPriorityId') || 2,
      sourceId: getSetting('defaultSourceId') || 2,
      technicianGroupIds: getSetting('defaultTechnicianGroupId')
        ? [Number(getSetting('defaultTechnicianGroupId'))]
        : [],
      assignedToId: getSetting('defaultAssignedToId') || 0,
      emailConfigurationId: getSetting('defaultEmailConfigurationId') || null,
      ticketLocation: getSetting('defaultTicketLocation') || 0,
      sendNotification: true,
    });

    const ticketNum = result?.ticketNumberWithPrefix || result?.ticketNumber || '';
    showNotification(
      'Gorelo — Ticket created',
      ticketNum ? `Ticket ${ticketNum} created from this email.` : 'Ticket created successfully.',
      'informationalMessage'
    );
  } catch (err) {
    showNotification(
      'Gorelo — Error',
      err.message || 'Could not create ticket. Check your settings.',
      'errorMessage'
    );
  }

  event.completed();
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getBodyText(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, { asyncContext: null }, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '');
      } else {
        resolve('');
      }
    });
  });
}

function buildDescription(fromName, fromEmail, bodyText) {
  const senderLine = fromEmail
    ? `From: ${fromName ? fromName + ' <' + fromEmail + '>' : fromEmail}\n\n`
    : '';
  return senderLine + (bodyText || '').trim();
}

function showNotification(title, message, type) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync('gorelo', {
      type: Office.MailboxEnums.ItemNotificationMessageType[type] || type,
      message: message,
      icon: 'icon16',
      persistent: false,
    });
  } catch {
    // Notification API may not be available in all hosts
    console.log('[Gorelo]', title, message);
  }
}
