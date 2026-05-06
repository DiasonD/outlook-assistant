/**
 * Update event functionality.
 *
 * Wraps the Microsoft Graph PATCH /me/events/{id} endpoint. Sends only
 * the fields the caller provides — anything left out is preserved
 * server-side. Useful for re-scheduling, adding/removing attendees,
 * editing the body, or tweaking the subject without rebuilding the
 * event from scratch (which loses RSVP state).
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { DEFAULT_TIMEZONE } = require('../config');

/**
 * Update event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateEvent(args) {
  const { eventId, subject, start, end, attendees, body, location } = args;

  if (!eventId) {
    return {
      content: [
        {
          type: 'text',
          text: 'Event ID is required to update an event.',
        },
      ],
    };
  }

  // Build the patch body from only the fields the caller actually provided.
  // Graph treats absent properties as "no change", so we never overwrite
  // something the user didn't intend to touch.
  const patch = {};

  if (subject !== undefined) patch.subject = subject;

  if (start !== undefined) {
    patch.start = {
      dateTime: start.dateTime || start,
      timeZone: start.timeZone || DEFAULT_TIMEZONE,
    };
  }

  if (end !== undefined) {
    patch.end = {
      dateTime: end.dateTime || end,
      timeZone: end.timeZone || DEFAULT_TIMEZONE,
    };
  }

  if (attendees !== undefined) {
    // Replaces the full attendee list — Graph PATCH on this property is
    // not additive. Caller must pass the desired complete list.
    patch.attendees = (attendees || []).map((email) => ({
      emailAddress: { address: email },
      type: 'required',
    }));
  }

  if (body !== undefined) {
    patch.body = { contentType: 'HTML', content: body };
  }

  if (location !== undefined) {
    patch.location = { displayName: location };
  }

  if (Object.keys(patch).length === 0) {
    return {
      content: [
        {
          type: 'text',
          text: 'No fields to update — provide at least one of: subject, start, end, attendees, body, location.',
        },
      ],
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const endpoint = `me/events/${eventId}`;

    const response = await callGraphAPI(accessToken, 'PATCH', endpoint, patch);

    const output = [
      `Event '${response.subject || eventId}' updated successfully.`,
    ];
    if (response.id) {
      output.push(`**ID**: \`${response.id}\``);
    }
    if (response.start) {
      output.push(
        `**Start**: ${response.start.dateTime} (${response.start.timeZone})`
      );
    }
    if (response.end) {
      output.push(
        `**End**: ${response.end.dateTime} (${response.end.timeZone})`
      );
    }
    if (response.webLink) {
      output.push(`**Link**: ${response.webLink}`);
    }
    output.push(`\nFields changed: ${Object.keys(patch).join(', ')}`);

    return {
      content: [
        {
          type: 'text',
          text: output.join('\n'),
        },
      ],
      _meta: {
        eventId: response.id,
        subject: response.subject,
        start: response.start,
        end: response.end,
        fieldsChanged: Object.keys(patch),
      },
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [
          {
            type: 'text',
            text: "Authentication required. Please use the 'authenticate' tool first.",
          },
        ],
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: `Error updating event: ${error.message}`,
        },
      ],
    };
  }
}

module.exports = handleUpdateEvent;
