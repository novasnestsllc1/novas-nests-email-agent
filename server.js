const express = require('express');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
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
  // Hoptel must be FIRST so crystal inn/west valley keywords match before slc_heart's 'slc' keyword
  hoptel: {
    name: 'Hoptel SLC',
    boardId: '18407764258',
    vaLocation: 'Salt Lake City VA Medical Center',
    keywords: ['crystal inn', 'hoptel', 'west valley', 'offsite va master', 'crystalinns.com'],
    notifiedColumnId: null,
    hotel: null,
    statusColumnId: 'color_mm27ywsm',
    manualOnly: false,  // ← change this to false
    // Hoptel-specific column IDs
    cols: {
      hotelName:    'color_mm27wqn1',
      status:       'color_mm27ywsm',
      checkin:      'date_mm27f0bk',
      checkout:     'date_mm27pjmd',
      nights:       'numeric_mm27n3m8',
      confirm:      'text_mm27ra9h',
      billingRate:  'numeric_mm27kx8m',
      totalBilled:  'numeric_mm27s0h6',
      hotelCost:    'numeric_mm27nfgf',
      totalCost:    'numeric_mm27ykq3',
      profitNight:  'numeric_mm27f7w3',
      totalProfit:  'numeric_mm278s8y',
      authorizedBy: 'color_mm2796fr',
      notes:        'long_text_mm273xnx',
      invoiceMonth: 'text_mm27xt1k'
    }
  },
  portland: {
    name: 'Portland VA Contract',
    boardId: '18406634526',
    vaLocation: 'Portland VA Health Care System',
    keywords: ['portland', 'self-care lodging', 'fisher house', 'best western', '503-220-8262'],
    notifiedColumnId: 'color_mm217f68',
    hotel: 'Best Western Portland West Beaverton',
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
    // Removed 'salt lake', 'slc' — too generic, causes false matches with Crystal Inn emails
    keywords: ['heart transplant', 'lvad', 'residence inn'],
    notifiedColumnId: null,
    hotel: 'Residence Inn',
    statusColumnId: 'color_mm1d6vgs'
  }
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

// ─── Microsoft Graph API — Delegated Auth ────────────────────────────────────

let msToken = null;
let msTokenExpiry = 0;
let currentRefreshToken = process.env.MS_REFRESH_TOKEN;

async function getMSToken() {
  // Return cached token if still valid
  if (msToken && Date.now() < msTokenExpiry) return msToken;

  if (!currentRefreshToken) {
    throw new Error('MS_REFRESH_TOKEN not set — run the OAuth setup tool first');
  }

  log('🔐 Refreshing Microsoft access token...');

  const resp = await fetch(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: MS_CLIENT_ID,
        client_secret: MS_CLIENT_SECRET,
        refresh_token: currentRefreshToken,
        scope: 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access',
        grant_type: 'refresh_token'
      })
    }
  );

  const data = await resp.json();

  if (!data.access_token) {
    throw new Error(`MS Token refresh failed: ${JSON.stringify(data)}`);
  }

  // Microsoft may rotate the refresh token — always store the latest one
  if (data.refresh_token) {
    currentRefreshToken = data.refresh_token;
    log('🔐 Refresh token rotated — using new token');
  }

  msToken = data.access_token;
  msTokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
  log('✓ Microsoft access token refreshed');
  return msToken;
}

async function getUnreadEmails() {
  const token = await getMSToken();
  // Use /me endpoint since we're now authenticated as the mailbox owner
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false&$top=20&$select=id,subject,body,from,receivedDateTime,hasAttachments`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await resp.json();
  if (data.error) throw new Error(`Graph API error: ${data.error.message}`);
  return data.value || [];
}

async function getEmailAttachments(emailId) {
  const token = await getMSToken();
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await resp.json();
  if (data.error) throw new Error(`Graph API error: ${data.error.message}`);
  return data.value || [];
}

async function markEmailAsRead(emailId) {
  const token = await getMSToken();
  await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${emailId}`,
    {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ isRead: true })
    }
  );
}

async function flagAndMarkRead(emailId, reason) {
  // Flag the email for manual review AND mark as read in one call
  // This stops the agent from ever retrying it while alerting your team
  const token = await getMSToken();
  await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${emailId}`,
    {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        isRead: true,
        flag: {
          flagStatus: 'flagged'
        }
      })
    }
  );
  log(`🚩 Email flagged for manual review — reason: ${reason}`);
}

// ─── Attachment extraction ────────────────────────────────────────────────────

async function extractTextFromAttachment(att) {
  const name = (att.name || '').toLowerCase();

  // PDF — send as native document to Claude
  if (att.contentBytes && name.endsWith('.pdf')) {
    log(`📎 PDF: ${att.name} (${att.contentBytes.length} base64 chars)`);
    return { type: 'document', mediaType: 'application/pdf', data: att.contentBytes, name: att.name };
  }

  // DOCX/DOC — extract raw text using mammoth
  if (att.contentBytes && (name.endsWith('.docx') || name.endsWith('.doc'))) {
    try {
      const buffer = Buffer.from(att.contentBytes, 'base64');
      // Check file header — valid DOCX (zip) starts with PK = 504b0304
      const header = buffer.slice(0, 4).toString('hex');
      log(`📎 DOCX: ${att.name} | size: ${buffer.length} bytes | header: ${header}`);

      if (header !== '504b0304') {
        log(`⚠ Attachment header is ${header} — not a valid DOCX. Likely still OME encrypted.`);
        return { type: 'text', content: `[Encrypted Word doc: ${att.name} — OME encryption prevented reading. Veteran details must be entered manually.]`, name: att.name };
      }

      const result = await mammoth.extractRawText({ buffer });
      const text = result.value || '';
      log(`📎 Extracted ${text.length} chars from ${att.name}`);
      return { type: 'text', content: text, name: att.name };
    } catch (e) {
      log(`⚠ mammoth failed for ${att.name}: ${e.message}`);
      return { type: 'text', content: `[Word document: ${att.name} — extraction failed]`, name: att.name };
    }
  }

  // Plain text/HTML body
  if (att.body?.content) {
    return { type: 'text', content: att.body.content, name: att.name };
  }

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

async function classifyAndIdentify(emailSubject, emailBody, attachmentNames, attachments) {
  // Fast keyword checks first — no Claude needed
  const combined = `${emailSubject} ${emailBody} ${attachmentNames.join(' ')}`.toLowerCase();

  // Check for Hoptel first before any SLC matching
  const hoptelKeywords = ['crystal inn', 'hoptel', 'west valley', 'offsite va master', 'crystalinns'];
  if (hoptelKeywords.some(k => combined.includes(k))) {
    return { emailType: 'new_reservation', contractKey: 'hoptel' };
  }

  // Cancellation/extension/update keywords
  const cancellationWords = ['cancel', 'cancelled', 'cancellation', 'no longer needs', 'no show', 'no-show', 'will not be', "won't be"];
  const extensionWords = ['extend', 'extension', 'extended', 'stay longer', 'additional night', 'additional nights', 'extra night'];
  const updateWords = ['update', 'change', 'modify', 'reschedule', 'rescheduled', 'correction', 'corrected'];

  let emailType = null;
  if (cancellationWords.some(w => combined.includes(w))) emailType = 'cancellation';
  else if (extensionWords.some(w => combined.includes(w))) emailType = 'extension';
  else if (updateWords.some(w => combined.includes(w))) emailType = 'update';

  // Contract keyword matching
  const contractKeywords = {
    portland: ['portland', 'self-care lodging', 'fisher house', 'best western', '503-220-8262'],
    wrj: ['white river junction', 'vermont', 'wrj'],
    slc_heart: ['heart transplant', 'lvad', 'residence inn']
  };

  let contractKey = null;
  for (const [key, keywords] of Object.entries(contractKeywords)) {
    if (keywords.some(k => combined.includes(k))) {
      contractKey = key;
      break;
    }
  }

  // If attachment present and no emailType yet — it's a new reservation
  const hasDocAttachment = attachmentNames.some(n => {
    const l = n.toLowerCase();
    return l.endsWith('.pdf') || l.endsWith('.docx') || l.endsWith('.doc');
  });
  if (!emailType && hasDocAttachment) emailType = 'new_reservation';

  // If we have both from keywords — no Claude needed
  if (emailType && contractKey) {
    log(`✓ Classified via keywords: ${emailType} / ${contractKey}`);
    return { emailType, contractKey };
  }

  // Single Claude call for both classification and contract ID
  log('🤖 Calling Claude for classification + contract ID...');

  const messageContent = [];

  // Include PDF attachments
  for (const att of (attachments || [])) {
    if (att.type === 'document' && att.mediaType === 'application/pdf') {
      messageContent.push({
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: att.data }
      });
    }
  }

  // Build attachment text context from DOCX
  let attContext = '';
  for (const att of (attachments || [])) {
    if (att.type === 'text' && att.content) {
      attContext += `\n\n[Attachment: ${att.name}]\n${att.content.substring(0, 2000)}`;
    }
  }

  const reservationWords = ['reservation', 'lodging', 'check-in', 'check in', 'arrival', 'veteran', 'voucher', 'booking', 'hotel conf', 'ck in', 'ck out', 'name of veteran', 'date of arrival'];
  if (!emailType) {
    emailType = reservationWords.some(w => combined.includes(w)) ? 'new_reservation' : null;
  }

  messageContent.push({
    type: 'text',
    text: `Analyze this VA lodging email and reply with ONLY a JSON object — no explanation:
{"type": "new_reservation|cancellation|extension|update|other", "contract": "portland|wrj|slc_heart|hoptel|unknown"}

Email subject: ${emailSubject}
Email body: ${emailBody.substring(0, 500)}${attContext}

Contract definitions:
- portland: Portland Oregon VA, Self-Care Lodging, Best Western, 503 area code, SW US Veterans Hospital
- wrj: White River Junction Vermont VA, 802 area code, Comfort Inn
- slc_heart: Salt Lake City Heart Transplant or LVAD, Residence Inn, flight info
- hoptel: Crystal Inn, West Valley, crystalinns — manual only
- unknown: cannot determine

Type definitions:
- new_reservation: new booking
- cancellation: cancelling stay
- extension: extending checkout
- update: changing details
- other: not VA lodging related

When in doubt on type, use new_reservation.`
  });

  const text = await callClaude(
    [{ role: 'user', content: messageContent }],
    'You analyze VA lodging emails. Reply with only a JSON object with "type" and "contract" fields.',
    100
  );

  try {
    const clean = text.replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(clean);
    const validTypes = ['new_reservation', 'cancellation', 'extension', 'update', 'other'];
    const validContracts = ['portland', 'wrj', 'slc_heart', 'hoptel', 'unknown'];

    return {
      emailType: emailType || (validTypes.includes(parsed.type) ? parsed.type : 'new_reservation'),
      contractKey: contractKey || (validContracts.includes(parsed.contract) ? parsed.contract : 'unknown')
    };
  } catch (e) {
    log(`⚠ Could not parse classification response: ${text.substring(0, 100)}`);
    return { emailType: emailType || 'new_reservation', contractKey: contractKey || 'unknown' };
  }
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

  // Include PDF attachments as native documents and DOCX as extracted text
  let attachmentTextContext = '';
  for (const att of attachments) {
    if (att.type === 'document' && att.mediaType === 'application/pdf') {
      log(`📎 Sending PDF ${att.name} to Claude`);
      messageContent.push({
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: att.data }
      });
    } else if (att.type === 'text' && att.content) {
      log(`📎 Including extracted text from ${att.name}`);
      attachmentTextContext += `\n\n[Document: ${att.name}]\n${att.content.substring(0, 3000)}`;
    }
  }

  const cleanBody = emailBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  messageContent.push({
    type: 'text',
    text: `Subject: ${emailSubject}\n\nEmail body:\n${cleanBody.substring(0, 2000)}${attachmentTextContext}\n\nExtract all VA reservations from the above content.`
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
// ─── Hoptel Spreadsheet Parser ────────────────────────────────────────────────

function parseHoptelSpreadsheet(buffer, filename) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  
  // Determine location from filename
  const isWestValley = filename.toLowerCase().includes('west valley') || 
                       filename.toLowerCase().includes('west_valley');
  const location = isWestValley ? 'Crystal Inn West Valley' : 'Crystal Inn SLC';
  const groupId = isWestValley ? 'group_mm275sdz' : 'group_mm27c6kr';
  
  // Rates
  const hotelCost = isWestValley ? 132.23 : 132.80;
  const billingRate = 138.72;
  const profitPerNight = parseFloat((billingRate - hotelCost).toFixed(2));

  const validBookings = [];
  const flaggedRows = [];

  // Process current month sheet only (first sheet with actual data)
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  // Data starts at row index 14 (0-based)
  for (let i = 14; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 10) continue;

    const confNum  = row[1];
    const name     = row[5];
    const checkinRaw  = row[6];
    const checkoutRaw = row[7];
    const nights   = row[9];
    const authBy   = row[12];
    const noshow   = row[15];

    // Skip empty slots — no name means unused room slot
    if (!name || String(name).trim() === '') continue;

    // Skip no-shows and cancellations
    if (noshow && String(noshow).trim() !== '') continue;

    const nameClean = String(name).trim();
    const confClean = confNum ? String(confNum).trim() : null;

    // Parse dates
    const checkin  = parseHoptelDate(checkinRaw);
    const checkout = parseHoptelDate(checkoutRaw);

    // Flag rows with a name but bad dates
    if (!checkin || !checkout) {
      flaggedRows.push({
        name: nameClean,
        confNum: confClean,
        checkinRaw: checkinRaw,
        checkoutRaw: checkoutRaw,
        row: i + 1,
        sheet: sheetName,
        file: filename
      });
      continue;
    }

    // Additional sanity checks — flag logically impossible dates
    const checkinDate  = new Date(checkin);
    const checkoutDate = new Date(checkout);
    const nightsCalc   = Math.round((checkoutDate - checkinDate) / (1000 * 60 * 60 * 24));

    if (checkoutDate <= checkinDate) {
      flaggedRows.push({
        name: nameClean,
        confNum: confClean,
        checkinRaw,
        checkoutRaw,
        row: i + 1,
        sheet: sheetName,
        file: filename,
        reason: 'Checkout is not after check-in'
      });
      continue;
    }

    const nightsActual = nights ? parseInt(nights) : nightsCalc;

    validBookings.push({
      location,
      groupId,
      confNum: confClean,
      name: nameClean,
      checkin,
      checkout,
      nights: nightsActual,
      authBy: authBy ? String(authBy).trim() : null,
      hotelCost,
      billingRate,
      profitPerNight,
      totalHotelCost: parseFloat((hotelCost * nightsActual).toFixed(2)),
      totalBilled: parseFloat((billingRate * nightsActual).toFixed(2)),
      totalProfit: parseFloat((profitPerNight * nightsActual).toFixed(2))
    });
  }

  log(`📊 ${filename}: ${validBookings.length} valid bookings, ${flaggedRows.length} flagged rows`);
  return { validBookings, flaggedRows, location };
}

function parseHoptelDate(val) {
  if (!val) return null;

  // Already a JS Date object (xlsx parsed it)
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return val.toISOString().split('T')[0];
  }

  const str = String(val).trim();
  if (!str) return null;

  // MM.DD.YY or MM.DD.YYYY
  const dotMatch = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (dotMatch) {
    let [_, m, d, y] = dotMatch;
    if (y.length === 2) y = '20' + y;
    const date = new Date(`${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`);
    if (!isNaN(date.getTime())) return date.toISOString().split('T')[0];
  }

  // ISO format
  const isoMatch = str.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoMatch) return isoMatch[1];

  return null;
}
// ─── Main email processor ─────────────────────────────────────────────────────

async function processEmail(email) {
  log(`📧 Processing: "${email.subject}" from ${email.from?.emailAddress?.address}`);

  const emailBody = email.body?.content || '';

  // Improved HTML cleaning — preserve table structure by converting to readable text
  let cleanBody = emailBody
    // Convert table rows to newlines to preserve structure
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/td>/gi, ' | ')
    .replace(/<\/th>/gi, ' | ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    // Strip remaining HTML tags
    .replace(/<[^>]+>/g, '')
    // Decode HTML entities
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .replace(/&#160;/g, ' ')
    // Clean up whitespace but preserve newlines
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  log(`📧 Body length: ${cleanBody.length} chars`);

  // Get attachments
  let processedAttachments = [];
  let attachmentNames = [];

  if (email.hasAttachments) {
    try {
      const attachments = await getEmailAttachments(email.id);
      attachmentNames = attachments.map(a => a.name || '');
      for (const att of attachments) {
        const extracted = await extractTextFromAttachment(att);
        if (extracted) {
          processedAttachments.push(extracted);
          log(`📎 Attachment: ${att.name} (${extracted.type})`);
        }
      }
    } catch (e) {
      log(`⚠ Could not read attachments: ${e.message} — continuing with email body`);
    }
  }

  // Step 1 & 2 combined: Classify AND identify contract in one Claude call
  let emailType, contractKey;
  try {
    const classified = await classifyAndIdentify(email.subject, cleanBody, attachmentNames, processedAttachments);
    emailType = classified.emailType;
    contractKey = classified.contractKey;
    log(`📋 Type: ${emailType} | Contract: ${contractKey}`);
  } catch (e) {
    log(`✗ Classification failed: ${e.message}`);
    await flagAndMarkRead(email.id, `Classification error: ${e.message}`);
    return;
  }

  if (emailType === 'other') {
    log(`ℹ Not a reservation email — flagging for manual review`);
    await flagAndMarkRead(email.id, 'Not identified as a VA reservation email');
    return;
  }

  // Hoptel is manual — flag and mark read
  // ── HOPTEL SPREADSHEET ────────────────────────────────────────────────────
  if (contractKey === 'hoptel') {
    // Find xlsx attachments
    const xlsxAtts = email.hasAttachments ? 
      (await getEmailAttachments(email.id)).filter(a => 
        (a.name || '').toLowerCase().endsWith('.xlsx')
      ) : [];

    if (xlsxAtts.length === 0) {
      log(`📋 Hoptel email has no xlsx — flagging for manual review`);
      await flagAndMarkRead(email.id, 'Hoptel email with no xlsx attachment');
      return;
    }

    const contract = CONTRACTS.hoptel;
    let totalCreated = 0;
    let totalFlagged = [];

    for (const att of xlsxAtts) {
      const buffer = Buffer.from(att.contentBytes, 'base64');
      const { validBookings, flaggedRows, location } = parseHoptelSpreadsheet(buffer, att.name);

      // Push each valid booking to Monday — deduplicate on confirmation number
      for (const booking of validBookings) {
        try {
          // Check if confirmation number already exists on the board
          if (booking.confNum) {
            const existing = await mondayQuery(`
              query {
                boards(ids: [${contract.boardId}]) {
                  items_page(limit: 500) {
                    items {
                      id
                      column_values(ids: ["${contract.cols.confirm}"]) { text }
                    }
                  }
                }
              }
            `);
            const items = existing?.data?.boards?.[0]?.items_page?.items || [];
            const duplicate = items.find(item =>
              item.column_values?.[0]?.text === booking.confNum
            );
            if (duplicate) {
              log(`⏭ Duplicate conf# ${booking.confNum} (${booking.name}) — skipping`);
              continue;
            }
          }

          // Get invoice month from check-in date
          const invoiceMonth = booking.checkin ? 
            new Date(booking.checkin + 'T12:00:00').toLocaleString('en-US', { month: 'long', year: 'numeric' }) : null;

          const colValues = {
            [contract.cols.status]:       { label: 'Working on it' },
            [contract.cols.hotelName]:    { label: location },
            [contract.cols.checkin]:      { date: booking.checkin },
            [contract.cols.checkout]:     { date: booking.checkout },
            [contract.cols.nights]:       booking.nights,
            [contract.cols.billingRate]:  booking.billingRate,
            [contract.cols.totalBilled]:  booking.totalBilled,
            [contract.cols.hotelCost]:    booking.hotelCost,
            [contract.cols.totalCost]:    booking.totalHotelCost,
            [contract.cols.profitNight]:  booking.profitPerNight,
            [contract.cols.totalProfit]:  booking.totalProfit,
          };

          if (booking.confNum)    colValues[contract.cols.confirm]      = booking.confNum;
          if (booking.authBy)     colValues[contract.cols.authorizedBy]  = { label: booking.authBy };
          if (invoiceMonth)       colValues[contract.cols.invoiceMonth]  = invoiceMonth;

          const mutation = `
            mutation {
              create_item(
                board_id: ${contract.boardId},
                group_id: "${booking.groupId}",
                item_name: "${booking.name.replace(/"/g, '\\"')}",
                column_values: ${JSON.stringify(JSON.stringify(colValues))}
              ) { id name }
            }
          `;
          const result = await mondayQuery(mutation);
          const item = result?.data?.create_item;
          if (item) {
            log(`✓ Hoptel created: ${item.name} (${item.id}) — ${location} conf#${booking.confNum}`);
            totalCreated++;
          }

          await new Promise(r => setTimeout(r, 300)); // small pause between creates

        } catch (e) {
          log(`✗ Hoptel Monday create error for ${booking.name}: ${e.message}`);
        }
      }

      totalFlagged.push(...flaggedRows);
    }

    // Send flagged rows summary email if any
    if (totalFlagged.length > 0) {
      log(`⚠ Sending flagged rows summary — ${totalFlagged.length} rows need review`);
      const flagLines = totalFlagged.map(f =>
        `• ${f.name} | Conf#: ${f.confNum || 'N/A'} | Check-in: "${f.checkinRaw}" | Check-out: "${f.checkoutRaw}" | File: ${f.file} | Row: ${f.row}${f.reason ? ' | Reason: ' + f.reason : ''}`
      ).join('\n');

      const token = await getMSToken();
      await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: {
            subject: `⚠ Hoptel SLC — ${totalFlagged.length} booking(s) could not be imported`,
            body: {
              contentType: 'Text',
              content: `The following veterans were in the Crystal Inn spreadsheet but could not be imported due to date errors. Please correct the spreadsheet and the agent will pick them up on the next send.\n\n${flagLines}\n\n— Nova's Nests Email Agent`
            },
            toRecipients: [{ emailAddress: { address: 'reservations@novasnestsgov.com' } }]
          }
        })
      });
    }

    log(`✓ Hoptel complete — ${totalCreated} created, ${totalFlagged.length} flagged`);
    await markEmailAsRead(email.id);
    return;
  }

  if (contractKey === 'unknown') {
    log(`⚠ Unknown contract: "${email.subject}" — flagging for manual review`);
    await flagAndMarkRead(email.id, 'Could not identify VA contract');
    return;
  }

  await new Promise(r => setTimeout(r, 3000)); // pause before extraction

  const contract = CONTRACTS[contractKey];
  log(`✓ Contract: ${contract.name}`);

  // Step 3: Extract reservation data
  let reservations = [];
  try {
    log(`🤖 Extracting with Claude (${emailType}) — body: ${cleanBody.length} chars, attachments: ${processedAttachments.length}`);
    if (cleanBody.length < 50 && processedAttachments.length === 0) {
      log(`⚠ Email body is very short (${cleanBody.length} chars) and no attachments — likely encrypted body`);
    }
    reservations = await extractReservations(email.subject, cleanBody, contractKey, processedAttachments, emailType);
  } catch (e) {
    log(`✗ Extraction failed: ${e.message}`);
    await flagAndMarkRead(email.id, `Extraction error: ${e.message}`);
    return;
  }

  if (reservations.length === 0) {
    log(`⚠ No reservations extracted — flagging for manual review`);
    await flagAndMarkRead(email.id, 'Claude could not extract reservation data — check attachment or email format');
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

  // Flag and mark read if everything failed — stops retrying
  if (failed > 0 && created === 0 && updated === 0 && cancelled === 0) {
    log(`✗ All Monday creates failed — flagging for manual review`);
    await flagAndMarkRead(email.id, 'Monday item creation failed — check Render logs');
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
    contracts: Object.keys(CONTRACTS).filter(k => !CONTRACTS[k].manualOnly),
    time: new Date().toISOString()
  });
});

app.post('/run-manual', async (req, res) => {
  res.json({ status: 'Manual check started' });
  log('▶ Manual run triggered');
  await pollEmails();
});

// ─── OAuth Setup Routes ───────────────────────────────────────────────────────
// Visit /auth to start the OAuth flow and capture your refresh token

app.get('/auth', (req, res) => {
  const params = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    response_type: 'code',
    redirect_uri: `https://novas-nests-email-agent.onrender.com/auth/callback`,
    scope: 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access',
    response_mode: 'query',
    login_hint: RESERVATIONS_EMAIL
  });

  const authUrl = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/authorize?${params}`;
  log('▶ OAuth flow started — redirecting to Microsoft login');
  res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
  const { code, error, error_description } = req.query;

  if (error) {
    log(`✗ OAuth error: ${error} — ${error_description}`);
    return res.send(`
      <h2 style="color:red;font-family:monospace">OAuth Error: ${error}</h2>
      <p style="font-family:monospace">${error_description}</p>
      <p style="font-family:monospace"><a href="/auth">Try again</a></p>
    `);
  }

  if (!code) {
    return res.send(`<p style="font-family:monospace">No code received. <a href="/auth">Try again</a></p>`);
  }

  try {
    log('🔐 Exchanging auth code for refresh token...');

    const resp = await fetch(
      `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: MS_CLIENT_ID,
          client_secret: MS_CLIENT_SECRET,
          code,
          redirect_uri: `https://novas-nests-email-agent.onrender.com/auth/callback`,
          grant_type: 'authorization_code',
          scope: 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access'
        })
      }
    );

    const data = await resp.json();

    if (!data.refresh_token) {
      log(`✗ Token exchange failed: ${JSON.stringify(data)}`);
      return res.send(`
        <h2 style="color:red;font-family:monospace">Token exchange failed</h2>
        <pre style="font-family:monospace">${JSON.stringify(data, null, 2)}</pre>
        <p><a href="/auth">Try again</a></p>
      `);
    }

    // Immediately use the new refresh token
    currentRefreshToken = data.refresh_token;
    log('✓ Refresh token captured and active — agent will now read encrypted attachments');

    res.send(`
      <!DOCTYPE html>
      <html>
      <head><style>
        body { background:#0a1628; color:#f0f4ff; font-family:monospace; padding:40px; }
        h1 { color:#c9a84c; letter-spacing:3px; }
        .token { background:#050d1a; border:1px solid rgba(201,168,76,0.3); padding:16px; border-radius:6px; word-break:break-all; font-size:12px; margin:16px 0; }
        .success { color:#2dd4a0; font-size:14px; margin-bottom:16px; }
        .steps { color:#7a8ba8; line-height:2; font-size:12px; }
        .steps b { color:#f0f4ff; }
        .copy-btn { background:#c9a84c; color:#0a1628; border:none; padding:10px 20px; border-radius:4px; cursor:pointer; font-family:monospace; font-weight:bold; margin-top:8px; }
      </style></head>
      <body>
        <h1>NOVA'S NESTS — OAUTH SETUP COMPLETE</h1>
        <div class="success">✓ Refresh token captured successfully — agent is now using delegated auth</div>
        <p style="color:#7a8ba8;font-size:12px;margin-bottom:8px;">Copy this token and add it to Render as MS_REFRESH_TOKEN:</p>
        <div class="token" id="token">${data.refresh_token}</div>
        <button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById('token').textContent).then(()=>this.textContent='✓ COPIED')">COPY TOKEN</button>
        <br><br>
        <div class="steps">
          <b>Add to Render:</b><br>
          1. Go to render.com → novas-nests-email-agent → Environment<br>
          2. Add variable: <b>MS_REFRESH_TOKEN</b><br>
          3. Paste the token above<br>
          4. Click Save Changes<br>
          5. Done — encrypted attachments now work permanently
        </div>
      </body>
      </html>
    `);

  } catch (e) {
    log(`✗ OAuth callback error: ${e.message}`);
    res.send(`<h2 style="color:red;font-family:monospace">Error: ${e.message}</h2><p><a href="/auth">Try again</a></p>`);
  }
});

// ─── Start ────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  log(`Nova's Nests Email Agent listening on port ${PORT}`);
  log(`Mailbox: ${RESERVATIONS_EMAIL}`);
  log(`Contracts: ${Object.values(CONTRACTS).map(c => c.name).join(', ')}`);
  startPolling();
});
