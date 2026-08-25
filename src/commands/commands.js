/**
 * Gorelo Ticket - Quick Create Command
 * Fires when the user clicks "Create Gorelo Ticket" in the Outlook ribbon.
 * Creates a ticket immediately with defaults, no task pane interaction needed.
 */

/* global Office */

Office.onReady(() => {
  // Register the command handler
  if (Office.context.requirements.isSetSupported('Mailbox', '1.1')) {
    Office.actions.associate('quickCreateTicket', quickCreateTicket);
  }
});

async function quickCreateTicket(event) {
  try {
    const item = Office.context.mailbox.item;

    // Gather email data
    const subject = item.subject;
    const from = item.from;
    const body = await getBodyAsync(item);

    // Load settings
    const { getSetting } = await import('../settings.js');
    const { getToken } = await import('../auth/auth.js');
    const { createTicket } = await import('../gorelo-api.js');

    // Ensure we have a token (may trigger login popup)
    await getToken();

    // Build ticket description with original email context
    const description = buildDescription(from, subject, body);

    // Create ticket with defaults
    await createTicket({
      title: subject || '(No subject)',
      description: description,
      markdownComment: `Email from: ${from.displayName} <${from.emailAddress}>\n\n${body.slice(0, 500)}`,
      clientId: getSetting('defaultClientId') || 0,
      clientName: getSetting('defaultClientName') || '',
      contactId: getSetting('defaultContactId') || 0,
      outboundContactEmail: from.emailAddress,
      technicianGroupIds: getSetting('defaultTechnicianGroupId')
        ? [getSetting('defaultTechnicianGroupId')]
        : [],
      assignedToId: getSetting('defaultAssignedToId') || 0,
      priorityId: getSetting('defaultPriorityId') || 2,
      sourceId: getSetting('defaultSourceId') || 2,
      emailConfigurationId: getSetting('defaultEmailConfigurationId') || 0,
      ticketLocation: getSetting('defaultTicketLocation') || 0,
      sendTicketCreatedEmail: true,
    });

    // Show success notification
    Office.context.mailbox.item.notificationMessages.addAsync('gorelo-success', {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: '✓ Gorelo ticket created successfully',
      icon: 'icon-16',
      persistent: false,
    });

  } catch (error) {
    console.error('Quick create failed:', error);

    Office.context.mailbox.item.notificationMessages.addAsync('gorelo-error', {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: 'Failed to create ticket: ' + error.message,
    });
  } finally {
    event.completed();
  }
}

/**
 * Build HTML description preserving original sender context.
 */
function buildDescription(from, subject, body) {
  return `
    <div style="border-left: 3px solid #1DAF94; padding-left: 12px; margin-bottom: 16px;">
      <p><strong>Original email from:</strong> ${escapeHtml(from.displayName)} &lt;${escapeHtml(from.emailAddress)}&gt;</p>
      <p><strong>Subject:</strong> ${escapeHtml(subject)}</p>
    </div>
    <div>${body}</div>
  `.trim();
}

function escapeHtml(str) {
  return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Get email body as HTML, falling back to text.
 */
function getBodyAsync(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, { asyncContext: 'html' }, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '');
      } else {
        // Fall back to plain text
        item.body.getAsync(Office.CoercionType.Text, {}, textResult => {
          resolve(textResult.value || '');
        });
      }
    });
  });
}
