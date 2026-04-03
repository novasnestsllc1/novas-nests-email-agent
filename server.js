const express = require('express');
const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const MONDAY_API_KEY = process.env.MONDAY_API_KEY;
const MS_TENANT_ID = process.env.MS_TENANT_ID;
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
const RESERVATIONS_EMAIL = 'reservations@novasnestsgov.com';

// ─── Contract Configuration ───────────────────────────────────────────────────

const CONTRACTS = {
  portland: {
    name: 'Portland VA Contract',
    boardId: '18406634526',
    vaLocation: 'Portland VA Health Care System',
    keywords: ['portland', 'self-care lodging', 'fisher house', 'best western', '503-220-8262'],
    notifiedColumnId: 'color_mm217f68',
    hotel: 'Best Western',
    statusColumnId: 'color_mm1d6vgs'
  },
  wrj: {
    name: 'WRJ VA Contract',
    boardId: '18403874769',
    vaLocation: 'White River Junction VA Medical Center',
    keywords: ['white river junction', 'vermont', 'wrj', 'vt '],
    notifiedColumnId: 'color_mm21sgv8',
    hotel: 'Comfort Inn',
    statusColumnId: 'color_mm1d6vgs'
  },
  slc_heart: {
    name: 'Heart Transplant SLC VA Contract',
    boardId: '18406339737',
    vaLocation: 'Salt Lake City VA Medical Center',
    keywords: ['heart transplant', 'lvad', 'residence inn', 'salt lake', 'slc'],
    notifiedColumnId: null,
    hotel: 'Residence Inn',
    statusColumnId: 'color_mm1d6vgs'
  },
  // Hoptel SLC is handled manually — Excel spreadsheet workflow
  // not compatible with automated item creation. Team processes manually.
};

// Shared Monday column IDs
const COLS = {
  phone:    'phone_mm1rs0ge',
  hotel:    'color_mm1dtr12',
  status:   'color_mm1d6vgs',
  checkin:  'date_mm1daepf',
  checkout: 'date_mm1dfe8r',
  guests:   'numeric_mm1dz7va',
  confirm:  'text_mm1d20vm',
  notes:    'long_text_mm1dppk1'
};

function log(msg) {
  console.log(`[${new Date().toISOString()}] ${msg}`);
}

// ─── Phone formatting — always +1 ────────────────────────────────────────────

function formatPhone(rawPhone) {
  if (!rawPhone) return null;
  const digits = String(rawPhone).replace(/\D/g, '');
  if (digits.length === 10) return '+1' + digits;
  if (digits.length === 11 && digits.startsWith('1')) return '+' + digits;
  if (digits.length > 10) return '+1' + digits.slice(-10);
  return null;
}

// ─── Microsoft Graph API ──────────────────────────────────────────────────────

let msToken = null;
let msTokenExpiry = 0;

async function getMSToken() {
  if (msToken && Date.now() < msTokenExpiry) return msToken;
  const resp = await fetch(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: MS_CLIENT_ID,
        client_secret: MS_CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    }
  );
  const data = await resp.json();
  if (!data.access_token) throw new Error(`MS Auth failed: ${JSON.stringify(data)}`);
  msToken = data.access_token;
  msTokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
  return msToken;
}

async function getUnreadEmails() {
  const token = await getMSToken();
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/users/${RESERVATIONS_EMAIL}/messages?$filter=isRead eq false&$top=20&$select=id,subject,body,from,receivedDateTime,hasAttachments`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await resp.json();
  return data.value || [];
}

async function getEmailAttachments(emailId) {
  const token = await getMSToken();
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/users/${RESERVATIONS_EMAIL}/messages/${emailId}/attachments`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await resp.json();
  return data.value || [];
}

async function markEmailAsRead(emailId) {
  const token = await getMSToken();
  await fetch(
    `https://graph.microsoft.com/v1.0/users/${RESERVATIONS_EMAIL}/messages/${emailId}`,
    {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ isRead: true })
    }
  );
}

// ─── Attachment extraction ────────────────────────────────────────────────────

function extractTextFromAttachment(att) {
  const name = (att.name || '').toLowerCase();
  if (att.contentBytes) {
    const mediaType = name.endsWith('.pdf') ? 'application/pdf'
      : (name.endsWith('.docx') || name.endsWith('.doc'))
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        : null;
    if (mediaType) return { type: 'document', mediaType, data: att.contentBytes, name: att.name };
  }
  if (att.body?.content) return { type: 'text', content: att.body.content, name: att.name };
  return null;
}

// ─── Claude AI ────────────────────────────────────────────────────────────────

async function callClaude(messages, systemPrompt, maxTokens = 2000) {
  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: maxTokens,
      system: systemPrompt,
      messages
    })
  });
  if (!resp.ok) {
    const err = await resp.json();
    throw new Error(`Claude API error: ${err.error?.message || JSON.stringify(err)}`);
  }
  const data = await resp.json();
  return data.content?.[0]?.text || '';
}

async function classifyEmail(emailSubject, emailBody, attachmentNames) {
  // If email has a PDF or DOCX attachment it's almost certainly a reservation
  const hasReservationAttachment = attachmentNames.some(name => {
    const lower = name.toLowerCase();
    return lower.endsWith('.pdf') || lower.endsWith('.docx') || lower.endsWith('.doc');
  });

  if (hasReservationAttachment) {
    log(`📎 Reservation attachment detected — treating as new_reservation`);
    return 'new_reservation';
  }

  // Check body for cancellation/extension keywords before calling Claude
  const combined = `${emailSubject} ${emailBody}`.toLowerCase();

  const cancellationWords = ['cancel', 'cancelled', 'cancellation', 'no longer needs', 'no show', 'no-show', 'will not be', 'won\'t be'];
  const extensionWords = ['extend', 'extension', 'extended', 'stay longer', 'additional night', 'additional nights', 'extra night'];
  const updateWords = ['update', 'change', 'modify', 'reschedule', 'rescheduled', 'correction', 'corrected'];
  const reservationWords = ['reservation', 'lodging', 'check-in', 'check in', 'arrival', 'veteran', 'voucher', 'booking', 'hotel conf', 'ck in', 'ck out'];

  if (cancellationWords.some(w => combined.includes(w))) return 'cancellation';
  if (extensionWords.some(w => combined.includes(w))) return 'extension';
  if (updateWords.some(w => combined.includes(w))) return 'update';
  if (reservationWords.some(w => combined.includes(w))) return 'new_reservation';

  // Ask Claude only if keywords didn't match
  const text = await callClaude([{
    role: 'user',
    content: `Classify this VA lodging email. Reply with ONLY one word:
- new_reservation: new booking request or lodging authorization
- cancellation: veteran cancelling or no longer needs lodging
- extension: extending an existing stay
- update: change to existing reservation
- other: clearly not related to VA lodging at all

Subject: ${emailSubject}
Body excerpt: ${emailBody.substring(0, 800)}`
  }], 'You classify VA lodging emails. When in doubt lean toward new_reservation. Reply with only one word.', 50);

  const result = text.trim().toLowerCase();
  const valid = ['new_reservation', 'cancellation', 'extension', 'update', 'other'];
  return valid.includes(result) ? result : 'new_reservation'; // default to new_reservation if unclear
}

async function identifyContract(emailSubject, emailBody, attachmentNames, attachments) {
  const combined = `${emailSubject} ${emailBody} ${attachmentNames.join(' ')}`.toLowerCase();

  // Fast keyword matching first — no Claude call needed
  for (const [key, contract] of Object.entries(CONTRACTS)) {
    for (const keyword of contract.keywords) {
      if (combined.includes(keyword.toLowerCase())) return key;
    }
  }

  // Keywords didn't match — send to Claude including attachment content so it can read the doc
  const messageContent = [];

  // Include document attachments for Claude to read and identify contract
  for (const att of (attachments || [])) {
    if (att.type === 'document') {
      log(`📎 Using ${att.name} to identify contract`);
      messageContent.push({
        type: 'document',
        source: { type: 'base64', media_type: att.mediaType, data: att.data }
      });
    }
  }

  messageContent.push({
    type: 'text',
    text: `Which VA contract is this email for? Reply with ONLY one word: portland, wrj, slc_heart, or unknown

Email subject: ${emailSubject}
Email body: ${emailBody.substring(0, 500)}
Attachment names: ${attachmentNames.join(', ')}

Contracts:
- portland: Portland Oregon VA Health Care System, Self-Care Lodging, Best Western, 503 area code, Christine Morgan, Fisher House, SW US Veterans Hospital
- wrj: White River Junction Vermont VA Medical Center, 802 area code, Comfort Inn
- slc_heart: Salt Lake City Heart Transplant or LVAD, Residence Inn, flight info, caregiver, hotel conf number in table format
- unknown: cannot determine from any content

Read any attached documents carefully to identify which VA facility and contract this belongs to.`
  });

  const result = await callClaude(
    [{ role: 'user', content: messageContent }],
    'You identify which VA contract an email belongs to by reading the email and any attached documents. Reply with only one word: portland, wrj, slc_heart, or unknown.',
    20
  );

  const r = result.trim().toLowerCase();
  return ['portland', 'wrj', 'slc_heart'].includes(r) ? r : 'unknown';
}

async function extractReservations(emailSubject, emailBody, contractKey, attachments, emailType) {
  const contract = CONTRACTS[contractKey] || { name: 'Unknown' };

  const systemPrompt = `You are a VA reservation data extraction specialist for Nova's Nests LLC.

Extract ALL reservations and return ONLY a valid JSON array. No explanation, no markdown, just raw JSON.

Each object must have exactly these fields (null if not found):
{
  "veteran_name": "First Last — convert Last, First to First Last",
  "veteran_phone": "10 digits only, no formatting, no country code, no +1",
  "checkin_date": "YYYY-MM-DD",
  "checkout_date": "YYYY-MM-DD",
  "num_guests": number or null,
  "hotel_confirmation": "string or null",
  "room_type": "string or null",
  "special_notes": "ADA, oxygen, service dog, ground floor, medical companion, etc or null",
  "caregiver_name": "string or null",
  "caregiver_phone": "10 digits only or null",
  "flight_info": "all flight details as one string or null"
}

Contract: ${contract.name}
Email type: ${emailType}

Rules:
- Convert Last, First → First Last
- Extract EVERY veteran if multiple appear
- Phone: strip all non-digits, remove leading 1 if 11 digits
- Dates: convert any format to YYYY-MM-DD
- Return [] if no valid reservations found`;

  const messageContent = [];

  // Include document attachments
  for (const att of attachments) {
    if (att.type === 'document') {
      log(`📎 Sending ${att.name} to Claude as document`);
      messageContent.push({
        type: 'document',
        source: { type: 'base64', media_type: att.mediaType, data: att.data }
      });
    }
  }

  const cleanBody = emailBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  messageContent.push({
    type: 'text',
    text: `Subject: ${emailSubject}\n\nBody:\n${cleanBody.substring(0, 3000)}\n\nExtract all VA reservations.`
  });

  const text = await callClaude(
    [{ role: 'user', content: messageContent }],
    systemPrompt
  );

  try {
    return JSON.parse(text.replace(/```json|```/g, '').trim());
  } catch (e) {
    log(`⚠ Parse error: ${e.message} — Raw: ${text.substring(0, 200)}`);
    return [];
  }
}

// ─── Monday API ───────────────────────────────────────────────────────────────

async function mondayQuery(query) {
  const resp = await fetch('https://api.monday.com/v2', {
    method: 'POST',
    headers: {
      Authorization: MONDAY_API_KEY,
      'Content-Type': 'application/json',
      'API-Version': '2024-01'
    },
    body: JSON.stringify({ query })
  });
  return resp.json();
}

async function findExistingItem(boardId, veteranName, checkinDate) {
  // Search board for matching veteran + checkin date
  const query = `
    query {
      boards(ids: [${boardId}]) {
        items_page(limit: 200) {
          items {
            id
            name
            column_values(ids: ["${COLS.checkin}", "${COLS.checkout}", "${COLS.status}"]) {
              id
              text
            }
          }
        }
      }
    }
  `;
  const result = await mondayQuery(query);
  const items = result?.data?.boards?.[0]?.items_page?.items || [];

  // Match on first name + checkin date for robustness
  const firstName = veteranName.toLowerCase().split(' ')[0];

  return items.find(item => {
    const nameMatch = item.name.toLowerCase().includes(firstName);
    const checkinCol = item.column_values?.find(c => c.id === COLS.checkin);
    const dateMatch = checkinDate ? checkinCol?.text === checkinDate : true;
    return nameMatch && dateMatch;
  }) || null;
}

async function findExistingItemByName(boardId, veteranName) {
  // Find by name only — for cancellations where we may not have the exact date
  const query = `
    query {
      boards(ids: [${boardId}]) {
        items_page(limit: 200) {
          items {
            id
            name
            column_values(ids: ["${COLS.checkin}", "${COLS.checkout}", "${COLS.status}"]) {
              id
              text
            }
          }
        }
      }
    }
  `;
  const result = await mondayQuery(query);
  const items = result?.data?.boards?.[0]?.items_page?.items || [];
  const firstName = veteranName.toLowerCase().split(' ')[0];
  const lastName = veteranName.toLowerCase().split(' ').slice(-1)[0];

  // Return most recently created match
  return items.find(item => {
    const nameLower = item.name.toLowerCase();
    return nameLower.includes(firstName) || nameLower.includes(lastName);
  }) || null;
}

async function createMondayItem(boardId, veteranName, reservation, contract) {
  const phone = formatPhone(reservation.veteran_phone);

  const colValues = {
    [COLS.status]: { label: 'Working on it' }
  };

  if (phone) colValues[COLS.phone] = phone;
  if (reservation.checkin_date) colValues[COLS.checkin] = { date: reservation.checkin_date };
  if (reservation.checkout_date) colValues[COLS.checkout] = { date: reservation.checkout_date };
  if (reservation.num_guests) colValues[COLS.guests] = reservation.num_guests;
  if (reservation.hotel_confirmation) colValues[COLS.confirm] = reservation.hotel_confirmation;
  if (contract.hotel) colValues[COLS.hotel] = { label: contract.hotel };

  const notesParts = [];
  if (reservation.special_notes) notesParts.push(reservation.special_notes);
  if (reservation.caregiver_name) notesParts.push(`Caregiver: ${reservation.caregiver_name}`);
  if (reservation.caregiver_phone) notesParts.push(`Caregiver Ph: +1${reservation.caregiver_phone}`);
  if (reservation.flight_info) notesParts.push(`Flight: ${reservation.flight_info}`);
  if (reservation.room_type) notesParts.push(`Room: ${reservation.room_type}`);
  if (notesParts.length > 0) colValues[COLS.notes] = notesParts.join(' | ');

  const mutation = `
    mutation {
      create_item(
        board_id: ${boardId},
        item_name: "${veteranName.replace(/"/g, '\\"')}",
        column_values: ${JSON.stringify(JSON.stringify(colValues))}
      ) { id name }
    }
  `;

  const result = await mondayQuery(mutation);
  return result?.data?.create_item;
}

async function cancelMondayItem(boardId, itemId, veteranName) {
  // Set status to Stuck (closest to cancelled on existing boards)
  // Note: "Stuck" is index 2 on all boards
  const colValues = { [COLS.status]: { label: 'Stuck' } };
  const mutation = `
    mutation {
      change_multiple_column_values(
        board_id: ${boardId},
        item_id: ${itemId},
        column_values: ${JSON.stringify(JSON.stringify(colValues))}
      ) { id name }
    }
  `;
  const result = await mondayQuery(mutation);
  log(`✓ Marked ${veteranName} as Stuck/Cancelled (item ${itemId})`);
  return result?.data?.change_multiple_column_values;
}

async function extendMondayItem(boardId, itemId, newCheckoutDate, veteranName) {
  const colValues = {
    [COLS.checkout]: { date: newCheckoutDate },
    [COLS.status]: { label: 'Working on it' }
  };
  const mutation = `
    mutation {
      change_multiple_column_values(
        board_id: ${boardId},
        item_id: ${itemId},
        column_values: ${JSON.stringify(JSON.stringify(colValues))}
      ) { id name }
    }
  `;
  const result = await mondayQuery(mutation);
  log(`✓ Extended ${veteranName} checkout to ${newCheckoutDate} (item ${itemId})`);
  return result?.data?.change_multiple_column_values;
}

async function updateMondayItem(boardId, itemId, reservation, veteranName) {
  const colValues = {};
  if (reservation.checkin_date) colValues[COLS.checkin] = { date: reservation.checkin_date };
  if (reservation.checkout_date) colValues[COLS.checkout] = { date: reservation.checkout_date };
  if (reservation.hotel_confirmation) colValues[COLS.confirm] = reservation.hotel_confirmation;
  if (reservation.num_guests) colValues[COLS.guests] = reservation.num_guests;
  colValues[COLS.status] = { label: 'Working on it' };

  const mutation = `
    mutation {
      change_multiple_column_values(
        board_id: ${boardId},
        item_id: ${itemId},
        column_values: ${JSON.stringify(JSON.stringify(colValues))}
      ) { id name }
    }
  `;
  const result = await mondayQuery(mutation);
  log(`✓ Updated ${veteranName} reservation (item ${itemId})`);
  return result?.data?.change_multiple_column_values;
}

// ─── Main email processor ─────────────────────────────────────────────────────

async function processEmail(email) {
  log(`📧 Processing: "${email.subject}" from ${email.from?.emailAddress?.address}`);

  const emailBody = email.body?.content || '';
  const cleanBody = emailBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

  // Get attachments
  let processedAttachments = [];
  let attachmentNames = [];

  if (email.hasAttachments) {
    try {
      const attachments = await getEmailAttachments(email.id);
      attachmentNames = attachments.map(a => a.name || '');
      for (const att of attachments) {
        const extracted = extractTextFromAttachment(att);
        if (extracted) {
          processedAttachments.push(extracted);
          log(`📎 Attachment: ${att.name} (${extracted.type})`);
        }
      }
    } catch (e) {
      log(`⚠ Could not read attachments: ${e.message} — continuing with email body`);
    }
  }

  // Step 1: Classify email type
  let emailType;
  try {
    emailType = await classifyEmail(email.subject, cleanBody, attachmentNames);
    log(`📋 Email type: ${emailType}`);
  } catch (e) {
    log(`✗ Classification failed: ${e.message} — leaving UNREAD`);
    return;
  }

  if (emailType === 'other') {
    log(`ℹ Not a reservation email — leaving UNREAD for manual review`);
    return;
  }

  // Step 2: Identify contract
  let contractKey;
  try {
    contractKey = await identifyContract(email.subject, cleanBody, attachmentNames, processedAttachments);
  } catch (e) {
    log(`✗ Contract ID failed: ${e.message} — leaving UNREAD`);
    return;
  }

  await new Promise(r => setTimeout(r, 3000)); // pause between Claude calls

  if (contractKey === 'unknown') {
    log(`⚠ Unknown contract: "${email.subject}" — leaving UNREAD for manual review`);
    return;
  }

  // Hoptel is manual — leave unread for team to process
  if (contractKey === 'hoptel') {
    log(`📋 Hoptel SLC email detected: "${email.subject}" — leaving UNREAD for manual processing`);
    return;
  }

  const contract = CONTRACTS[contractKey];
  log(`✓ Contract: ${contract.name}`);

  // Step 3: Extract reservation data
  let reservations = [];
  try {
    log(`🤖 Extracting with Claude (${emailType})...`);
    reservations = await extractReservations(email.subject, cleanBody, contractKey, processedAttachments, emailType);
  } catch (e) {
    log(`✗ Extraction failed: ${e.message} — leaving UNREAD`);
    return;
  }

  if (reservations.length === 0) {
    log(`⚠ No reservations extracted — leaving UNREAD for manual review`);
    return;
  }

  log(`✓ Extracted ${reservations.length} reservation(s) — type: ${emailType}`);

  let created = 0;
  let updated = 0;
  let cancelled = 0;
  let skipped = 0;
  let failed = 0;

  for (const res of reservations) {
    if (!res.veteran_name) {
      log(`⚠ No veteran name — skipping`);
      skipped++;
      continue;
    }

    try {
      // ── CANCELLATION ──────────────────────────────────────────────────────
      if (emailType === 'cancellation') {
        const existing = await findExistingItemByName(contract.boardId, res.veteran_name);
        if (existing) {
          await cancelMondayItem(contract.boardId, existing.id, res.veteran_name);
          log(`🚫 Cancelled: ${res.veteran_name} (item ${existing.id})`);
          cancelled++;
        } else {
          log(`⚠ No existing item found for cancellation: ${res.veteran_name} — leaving UNREAD`);
          failed++;
        }
        continue;
      }

      // ── EXTENSION ─────────────────────────────────────────────────────────
      if (emailType === 'extension') {
        const existing = await findExistingItem(contract.boardId, res.veteran_name, res.checkin_date)
          || await findExistingItemByName(contract.boardId, res.veteran_name);

        if (existing && res.checkout_date) {
          await extendMondayItem(contract.boardId, existing.id, res.checkout_date, res.veteran_name);
          log(`📅 Extended: ${res.veteran_name} → checkout ${res.checkout_date}`);
          updated++;
        } else if (!existing) {
          // Extension for reservation not yet in Monday — create it
          log(`⚠ No existing item for extension — creating new item for ${res.veteran_name}`);
          const item = await createMondayItem(contract.boardId, res.veteran_name, res, contract);
          if (item) { created++; }
        } else {
          log(`⚠ Extension found existing but no checkout date — skipping ${res.veteran_name}`);
          skipped++;
        }
        continue;
      }

      // ── UPDATE ────────────────────────────────────────────────────────────
      if (emailType === 'update') {
        const existing = await findExistingItem(contract.boardId, res.veteran_name, res.checkin_date)
          || await findExistingItemByName(contract.boardId, res.veteran_name);

        if (existing) {
          await updateMondayItem(contract.boardId, existing.id, res, res.veteran_name);
          log(`✏️ Updated: ${res.veteran_name} (item ${existing.id})`);
          updated++;
        } else {
          // Not found — treat as new
          log(`⚠ No existing item for update — creating new for ${res.veteran_name}`);
          const item = await createMondayItem(contract.boardId, res.veteran_name, res, contract);
          if (item) { created++; }
        }
        continue;
      }

      // ── NEW RESERVATION ───────────────────────────────────────────────────
      if (!res.checkin_date) {
        log(`⚠ No check-in date for ${res.veteran_name} — skipping`);
        skipped++;
        continue;
      }

      // Check for exact duplicate (same veteran + same checkin)
      const existing = await findExistingItem(contract.boardId, res.veteran_name, res.checkin_date);
      if (existing) {
        log(`⏭ Duplicate: ${res.veteran_name} on ${res.checkin_date} — skipping`);
        skipped++;
        continue;
      }

      // Create new item
      const item = await createMondayItem(contract.boardId, res.veteran_name, res, contract);
      if (item) {
        log(`✓ Created: ${item.name} (${item.id}) — phone: ${formatPhone(res.veteran_phone) || 'none'}`);
        created++;
      } else {
        log(`✗ Monday create returned null for ${res.veteran_name}`);
        failed++;
      }

    } catch (e) {
      log(`✗ Error processing ${res.veteran_name}: ${e.message}`);
      failed++;
    }
  }

  // Only mark as read if something succeeded
  if (failed > 0 && created === 0 && updated === 0 && cancelled === 0) {
    log(`✗ Nothing succeeded — leaving email UNREAD for manual review`);
    return;
  }

  log(`✓ Complete — ${created} created, ${updated} updated, ${cancelled} cancelled, ${skipped} skipped, ${failed} failed`);
  await markEmailAsRead(email.id);
}

// ─── Polling ──────────────────────────────────────────────────────────────────

async function pollEmails() {
  log('📬 Checking for new emails...');
  try {
    const emails = await getUnreadEmails();
    if (emails.length === 0) { log('✓ No new emails'); return; }
    log(`Found ${emails.length} unread email(s)`);
    for (const email of emails) {
      await processEmail(email);
      await new Promise(r => setTimeout(r, 15000)); // 15 second pause between emails to avoid rate limits
    }
  } catch (e) {
    log(`✗ Poll error: ${e.message}`);
  }
}

function startPolling() {
  log('🔄 Email polling started — every 5 minutes');
  pollEmails();
  setInterval(pollEmails, 5 * 60 * 1000);
}

// ─── Routes ───────────────────────────────────────────────────────────────────

app.post('/webhook', async (req, res) => {
  if (req.body?.challenge) {
    log('✓ Webhook challenge verified');
    return res.json({ challenge: req.body.challenge });
  }
  res.json({ status: 'received' });
});

app.get('/', (req, res) => {
  res.json({
    status: "Nova's Nests Email Agent running",
    mailbox: RESERVATIONS_EMAIL,
    contracts: Object.keys(CONTRACTS),
    time: new Date().toISOString()
  });
});

app.post('/run-manual', async (req, res) => {
  res.json({ status: 'Manual check started' });
  log('▶ Manual run triggered');
  await pollEmails();
});

// ─── Start ────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  log(`Nova's Nests Email Agent listening on port ${PORT}`);
  log(`Mailbox: ${RESERVATIONS_EMAIL}`);
  log(`Contracts: ${Object.values(CONTRACTS).map(c => c.name).join(', ')}`);
  startPolling();
});
