const express = require('express');
const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3001;
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
    notifiedColumnId: 'color_mm217f68'
  },
  wrj: {
    name: 'WRJ VA Contract',
    boardId: '18403874769',
    vaLocation: 'White River Junction VA Medical Center',
    keywords: ['white river junction', 'vermont', 'wrj', 'vt '],
    notifiedColumnId: 'color_mm21sgv8'
  },
  slc_heart: {
    name: 'Heart Transplant SLC VA Contract',
    boardId: '18406339737',
    vaLocation: 'Salt Lake City VA Medical Center',
    keywords: ['heart transplant', 'lvad', 'residence inn', 'salt lake', 'slc'],
    notifiedColumnId: null // will be created
  },
  hoptel: {
    name: 'Hoptel SLC',
    boardId: '18406338444',
    vaLocation: 'Salt Lake City VA Medical Center',
    keywords: ['crystal inn', 'hoptel', 'west valley', 'offsite va master'],
    notifiedColumnId: null // excel based, different flow
  }
};

// Shared Monday column IDs (same across all boards)
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
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ isRead: true })
    }
  );
}

// ─── Claude AI Extraction ─────────────────────────────────────────────────────

async function extractWithClaude(content, contentType, contractHint) {
  const systemPrompt = `You are a VA reservation data extraction specialist for Nova's Nests LLC, a veteran-owned federal lodging broker.

Extract reservation data from the provided content and return ONLY a JSON array of reservation objects. No explanation, no markdown, just the JSON array.

Each reservation object must have these fields (use null if not found):
{
  "veteran_name": "First Last format",
  "veteran_phone": "digits only, no formatting",
  "checkin_date": "YYYY-MM-DD",
  "checkout_date": "YYYY-MM-DD",
  "num_guests": number or null,
  "hotel_confirmation": "string or null",
  "room_type": "string or null",
  "special_notes": "any special requests, ADA, oxygen, service dog, etc or null",
  "caregiver_name": "string or null",
  "caregiver_phone": "digits only or null",
  "flight_info": "string or null"
}

Content type: ${contentType}
Contract hint: ${contractHint}

Rules:
- If veteran name is in "Last, First" format, convert to "First Last"
- Extract ALL veterans if multiple are in one email
- For SLC Heart Transplant emails with a table, extract each row as a separate reservation
- For Hoptel Excel content, extract each veteran row as a separate reservation
- Phone numbers: strip all non-digits, if 11 digits starting with 1 strip the leading 1
- Dates: convert any format to YYYY-MM-DD
- Return empty array [] if no valid reservations found`;

  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 2000,
      system: systemPrompt,
      messages: [{ role: 'user', content: `Extract reservations from this content:\n\n${content}` }]
    })
  });

  const data = await resp.json();
  const text = data.content?.[0]?.text || '[]';

  try {
    const clean = text.replace(/```json|```/g, '').trim();
    return JSON.parse(clean);
  } catch (e) {
    log(`⚠ Claude extraction parse error: ${e.message}`);
    return [];
  }
}

async function identifyContract(emailSubject, emailBody, attachmentNames) {
  const combined = `${emailSubject} ${emailBody} ${attachmentNames.join(' ')}`.toLowerCase();

  // Check each contract's keywords
  for (const [key, contract] of Object.entries(CONTRACTS)) {
    for (const keyword of contract.keywords) {
      if (combined.includes(keyword.toLowerCase())) {
        return key;
      }
    }
  }

  // If unclear, ask Claude
  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 100,
      messages: [{
        role: 'user',
        content: `Which VA contract is this email for? Reply with ONLY one of: portland, wrj, slc_heart, hoptel, unknown

Email subject: ${emailSubject}
Email body excerpt: ${emailBody.substring(0, 500)}
Attachments: ${attachmentNames.join(', ')}

Contracts:
- portland: Portland Oregon VA Health Care System, Self-Care Lodging, Best Western
- wrj: White River Junction Vermont VA Medical Center
- slc_heart: Salt Lake City Heart Transplant/LVAD, Residence Inn
- hoptel: Salt Lake City Crystal Inn West Valley or SLC, Hoptel, Excel spreadsheet`
      }]
    })
  });

  const data = await resp.json();
  const result = data.content?.[0]?.text?.trim().toLowerCase() || 'unknown';
  return ['portland', 'wrj', 'slc_heart', 'hoptel'].includes(result) ? result : 'unknown';
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

async function createMondayItem(boardId, veteranName, reservation, hotelName) {
  const columnValues = {
    [COLS.status]: { label: 'Working on it' }
  };

  if (reservation.veteran_phone) columnValues[COLS.phone] = reservation.veteran_phone;
  if (reservation.checkin_date) columnValues[COLS.checkin] = { date: reservation.checkin_date };
  if (reservation.checkout_date) columnValues[COLS.checkout] = { date: reservation.checkout_date };
  if (reservation.num_guests) columnValues[COLS.guests] = reservation.num_guests;
  if (reservation.hotel_confirmation) columnValues[COLS.confirm] = reservation.hotel_confirmation;
  if (hotelName) columnValues[COLS.hotel] = { label: hotelName };

  // Build notes from special fields
  const notesParts = [];
  if (reservation.special_notes) notesParts.push(reservation.special_notes);
  if (reservation.caregiver_name) notesParts.push(`Caregiver: ${reservation.caregiver_name}`);
  if (reservation.caregiver_phone) notesParts.push(`Caregiver Ph: ${reservation.caregiver_phone}`);
  if (reservation.flight_info) notesParts.push(`Flight: ${reservation.flight_info}`);
  if (reservation.room_type) notesParts.push(`Room: ${reservation.room_type}`);
  if (notesParts.length > 0) columnValues[COLS.notes] = notesParts.join(' | ');

  const mutation = `
    mutation {
      create_item(
        board_id: ${boardId},
        item_name: "${veteranName.replace(/"/g, '\\"')}",
        column_values: ${JSON.stringify(JSON.stringify(columnValues))}
      ) { id name }
    }
  `;

  const result = await mondayQuery(mutation);
  return result?.data?.create_item;
}

async function checkDuplicateItem(boardId, veteranName, checkinDate) {
  const query = `
    query {
      boards(ids: [${boardId}]) {
        items_page(limit: 100) {
          items {
            id
            name
            column_values(ids: ["${COLS.checkin}"]) { text }
          }
        }
      }
    }
  `;

  const result = await mondayQuery(query);
  const items = result?.data?.boards?.[0]?.items_page?.items || [];

  return items.some(item => {
    const nameMatch = item.name.toLowerCase().includes(veteranName.toLowerCase().split(' ')[0].toLowerCase());
    const dateMatch = item.column_values?.[0]?.text === checkinDate;
    return nameMatch && dateMatch;
  });
}

// ─── Main Email Processor ─────────────────────────────────────────────────────

async function processEmail(email) {
  log(`📧 Processing: "${email.subject}" from ${email.from?.emailAddress?.address}`);

  const emailBody = email.body?.content || '';
  // Strip HTML tags for cleaner text
  const cleanBody = emailBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

  // Get attachments if any
  let attachments = [];
  let attachmentNames = [];
  let attachmentContent = '';

  if (email.hasAttachments) {
    attachments = await getEmailAttachments(email.id);
    attachmentNames = attachments.map(a => a.name || '');

    for (const att of attachments) {
      const name = (att.name || '').toLowerCase();
      // Process PDF, DOCX, XLSX, TXT attachments
      if (name.endsWith('.pdf') || name.endsWith('.docx') || name.endsWith('.doc')) {
        // For Word/PDF - use the text content if available
        if (att.contentBytes) {
          attachmentContent += `\n[Attachment: ${att.name}]\n`;
          // Note: Binary content - Claude will use email body context
          // Full PDF/DOCX parsing would require additional libraries on server
        }
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        attachmentContent += `\n[Excel attachment: ${att.name} - process as Hoptel SLC spreadsheet]\n`;
      } else if (att.body?.content) {
        attachmentContent += `\n[Attachment: ${att.name}]\n${att.body.content}`;
      }
    }
  }

  // Identify which contract this is for
  const contractKey = await identifyContract(email.subject, cleanBody, attachmentNames);

  if (contractKey === 'unknown') {
    log(`⚠ Could not identify contract for email: "${email.subject}" — skipping`);
    await markEmailAsRead(email.id);
    return;
  }

  const contract = CONTRACTS[contractKey];
  log(`✓ Identified contract: ${contract.name}`);

  // Build full content for Claude
  const fullContent = `
Subject: ${email.subject}
From: ${email.from?.emailAddress?.address}
Body: ${cleanBody}
${attachmentContent}
  `.trim();

  // Extract reservations with Claude
  const reservations = await extractWithClaude(fullContent, 'email', contract.name);

  if (reservations.length === 0) {
    log(`⚠ No reservations extracted from email: "${email.subject}"`);
    await markEmailAsRead(email.id);
    return;
  }

  log(`✓ Extracted ${reservations.length} reservation(s)`);

  // Determine hotel name from contract
  let hotelName = null;
  if (contractKey === 'portland') hotelName = 'Best Western';
  else if (contractKey === 'wrj') hotelName = 'Comfort Inn';
  else if (contractKey === 'slc_heart') hotelName = 'Residence Inn';
  else if (contractKey === 'hoptel') hotelName = 'Crystal Inn';

  // Create Monday items for each reservation
  let created = 0;
  let skipped = 0;

  for (const res of reservations) {
    if (!res.veteran_name || !res.checkin_date) {
      log(`⚠ Skipping incomplete reservation: ${JSON.stringify(res)}`);
      skipped++;
      continue;
    }

    // Check for duplicate
    const isDuplicate = await checkDuplicateItem(contract.boardId, res.veteran_name, res.checkin_date);
    if (isDuplicate) {
      log(`⏭ Duplicate found for ${res.veteran_name} on ${res.checkin_date} — skipping`);
      skipped++;
      continue;
    }

    const item = await createMondayItem(contract.boardId, res.veteran_name, res, hotelName);
    if (item) {
      log(`✓ Created Monday item: ${item.name} (ID: ${item.id}) on ${contract.name}`);
      created++;
    }
  }

  log(`✓ Email processed — ${created} created, ${skipped} skipped`);

  // Mark email as read
  await markEmailAsRead(email.id);
}

// ─── Polling ──────────────────────────────────────────────────────────────────

async function pollEmails() {
  log('📬 Checking reservations@novasnestsgov.com for new emails...');

  try {
    const emails = await getUnreadEmails();

    if (emails.length === 0) {
      log('✓ No new emails');
      return;
    }

    log(`Found ${emails.length} unread email(s)`);

    for (const email of emails) {
      await processEmail(email);
      await new Promise(r => setTimeout(r, 1000)); // brief pause between emails
    }

  } catch (e) {
    log(`✗ Poll error: ${e.message}`);
  }
}

function startPolling() {
  log('🔄 Starting email polling every 5 minutes...');
  pollEmails();
  setInterval(pollEmails, 5 * 60 * 1000);
}

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/', (req, res) => {
  res.json({
    status: "Nova's Nests Email Agent running",
    mailbox: RESERVATIONS_EMAIL,
    contracts: Object.keys(CONTRACTS),
    time: new Date().toISOString()
  });
});

app.post('/run-manual', async (req, res) => {
  res.json({ status: 'Manual email check started' });
  log('▶ Manual email check triggered');
  await pollEmails();
});

// ─── Start ────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  log(`Nova's Nests Email Agent listening on port ${PORT}`);
  log(`Monitoring: ${RESERVATIONS_EMAIL}`);
  log(`Contracts: ${Object.values(CONTRACTS).map(c => c.name).join(', ')}`);
  startPolling();
});
