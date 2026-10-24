/**
 * IMC Connects — live-capacity submission backend (v14 self-sync + cancellation archive)
 *
 * Paste this entire file into Code.gs in a BOUND Apps Script project
 * attached to a blank Google Sheet.
 *
 * FIRST RUN:
 *   1. Save.
 *   2. Run setupSheets() manually.
 *   3. Approve permissions.
 *   4. Deploy as a Web app: Execute as Me; access = Anyone (if available).
 *
 * WEBSITE API:
 *   GET  ?action=capacity
 *   GET  ?action=capacity&callback=myCallback        // browser-safe JSONP
 *   GET  ?action=submissionStatus&token=CLIENT_TOKEN&email=EMAIL&callback=myCallback
 *        -> returns the exact POST result once complete; falls back to the
 *           Submissions sheet for successful ACTIVE writes
 *   GET  ?action=ping
 *
 *   POST text/plain JSON (server-to-server/direct test), OR ordinary
 *   browser form POST with fields transport=iframe, responseToken=...,
 *   and payload=<JSON string>:
 *   {
 *     "action": "submit",
 *     "name": "Student Name",
 *     "email": "student@u.northwestern.edu",
 *     "guaranteed": [
 *       {"alumniKey":"...", "name":"...", "company":"...", "timestamp":"..."}
 *     ],
 *     "wishlist": [
 *       {"alumniKey":"...", "name":"...", "company":"...", "timestamp":"..."}
 *     ]
 *   }
 *
 *   POST text/plain JSON:
 *   {
 *     "action": "feedback",
 *     "submissionId": "...",
 *     "email": "student@u.northwestern.edu",
 *     "rating": 5,
 *     "feedback": "Optional text"
 *   }
 *
 * IMPORTANT:
 * - Guaranteed selections are NOT reserved until submit succeeds.
 * - My List can contain up to 8 alumni total.
 * - Up to 3 guaranteed selections are allowed.
 * - Everyone else is stored as wishlist; wishlist entries never consume capacity.
 * - Duplicate email checking is case-insensitive and whitespace-trimmed.
 * - CANCELLED/ERROR submissions do not block the same email from submitting again.
 * - Organizer-cancelled submissions are moved to Test Submissions Archive.
 * - Missing guaranteed alumni trigger one automatic roster sync before rejection.
 */

const CONFIG = Object.freeze({
  BACKEND_VERSION: 14,
  ALUMNI_CSV_URL:
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vQzxeSvTj4rQrR9zgUJHcR6iWvXidfCJUTRhSdD2NfiyUj_gfH3uVr331McW6CJwWfgobE8WRvEO1Ws/pub?gid=1834383749&single=true&output=csv',

  MAX_GUARANTEED: 3,
  MAX_MY_LIST: 8,
  ASSUMED_CAPACITY_IF_BLANK: 2,
  INITIAL_WISHLIST_COLUMNS: 7,
  CAPACITY_CACHE_SECONDS: 120,

  SHEETS: {
    SUBMISSIONS: 'Submissions',
    ARCHIVE: 'Test Submissions Archive',
    FEEDBACK: 'Feedback',
    CAPACITY: 'Alumni Capacity',
    RESERVATIONS: 'Reservations',
    MATCHES: 'Guaranteed Matches'
  },

  PROP_SPREADSHEET_ID: 'CMI_CONNECTS_SPREADSHEET_ID'
});

const SOURCE_HEADERS = Object.freeze({
  TIMESTAMP: 'Timestamp',
  NAME: 'Name',
  ROLE: 'Current role. (If in between jobs, please state your most recent role and your incoming role.)',
  COMPANY: 'Current company. (If in between jobs, please share the last company at which you worked.)',
  CAPACITY: 'How many coffee chats with students can you support this fall? (Please enter a number. If you are unsure of your bandwidth, you can give a conservative estimate.)'
});

const SUBMISSION_CORE_HEADERS = Object.freeze([
  'Timestamp',
  'Submission ID',
  'Client Token',
  'Status',
  'Student Name',
  'Student Email',
  'Normalized Email',

  'Alum 1 Key',
  'Alum 1',
  'Alum 1 Company',
  'Alum 2 Key',
  'Alum 2',
  'Alum 2 Company',
  'Alum 3 Key',
  'Alum 3',
  'Alum 3 Company',

  'Receipt Email Status',
  'Receipt Email Sent At',
  'Cancelled At',
  'Error'
]);

const CAPACITY_HEADERS = Object.freeze([
  'Alumni Key',
  'Name',
  'Company',
  'Role',
  'Source Timestamp',
  'Starting Capacity',
  'Capacity Limit',
  'Spots Reserved',
  'Spots Remaining',
  'Overbooked By',
  'Source Active',
  'Last Synced'
]);

const RESERVATION_HEADERS = Object.freeze([
  'Reservation ID',
  'Reserved At',
  'Submission ID',
  'Student Name',
  'Student Email',
  'Alumni Key',
  'Alumni Name',
  'Company',
  'Status',
  'Cancelled At'
]);

const MATCH_HEADERS = Object.freeze([
  'Submission ID',
  'Student Name',
  'Student Email',
  'Alumni Key',
  'Alumni Name',
  'Company',
  'Reserved At'
]);

const FEEDBACK_HEADERS = Object.freeze([
  'Timestamp',
  'Feedback ID',
  'Submission ID',
  'Student Email',
  'Normalized Email',
  'Rating',
  'Feedback'
]);


/* ========================================================================== */
/* SETUP + ORGANIZER MENU                                                     */
/* ========================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IMC Connects')
    .addItem('Setup / sync alumni', 'setupSheets')
    .addItem('Sync alumni only', 'syncAlumniOnly')
    .addSeparator()
    .addItem('Recalculate spots', 'recalculateSpots')
    .addItem('Rebuild guaranteed matches', 'rebuildGuaranteedMatches')
    .addSeparator()
    .addItem('Cancel a submission…', 'cancelSubmissionPrompt')
    .addItem('Archive cancelled submissions', 'archiveCancelledSubmissions')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the bound spreadsheet and run setupSheets() from its Apps Script project.');

  PropertiesService.getScriptProperties().setProperty(
    CONFIG.PROP_SPREADSHEET_ID,
    ss.getId()
  );

  const submissionsSheet = ensureSheet_(
    ss,
    CONFIG.SHEETS.SUBMISSIONS,
    submissionHeaders_(CONFIG.INITIAL_WISHLIST_COLUMNS)
  );
  normalizeGuaranteedColumnsToThree_(submissionsSheet);
  normalizeWishlistColumnsToSeven_(submissionsSheet);
  ensureArchiveSheet_(ss, submissionsSheet);

  ensureSheet_(ss, CONFIG.SHEETS.FEEDBACK, FEEDBACK_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.CAPACITY, CAPACITY_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.RESERVATIONS, RESERVATION_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.MATCHES, MATCH_HEADERS);

  formatSheets_(ss);
  syncAlumni_();
  recalculateSpots_();
  invalidateCapacityCache_();
  rebuildGuaranteedMatches_();

  SpreadsheetApp.getUi().alert(
    'IMC Connects setup complete.\n\n' +
    'Created/verified: Submissions, Test Submissions Archive, Feedback, Alumni Capacity, Reservations, and Guaranteed Matches.'
  );
}


/**
 * One-time repair for a blank sheet created by an older 5-alumni setup.
 * Run this manually if Alum 4 / Alum 5 columns are visible.
 */
function repairThreeAlumniSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || getDb_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SUBMISSIONS);
  if (!sheet) throw new Error('Submissions sheet not found.');

  normalizeGuaranteedColumnsToThree_(sheet);

  SpreadsheetApp.getUi().alert(
    'Submission setup repaired. Guaranteed selections are now capped at exactly 3 alumni.'
  );
}

function syncAlumniOnly() {
  syncAlumni_();
  recalculateSpots_();
  rebuildGuaranteedMatches_();
  SpreadsheetApp.getUi().alert('Alumni sync complete.');
}

function recalculateSpots() {
  withScriptLock_(function () {
    recalculateSpots_();
    invalidateCapacityCache_();
    rebuildGuaranteedMatches_();
  });
  SpreadsheetApp.getUi().alert('Spots recalculated.');
}

function rebuildGuaranteedMatches() {
  rebuildGuaranteedMatches_();
  SpreadsheetApp.getUi().alert('Guaranteed Matches rebuilt.');
}

function cancelSubmissionPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Cancel submission',
    'Enter a Submission ID or the student email address. Capacity will be restored and the student will be able to submit again.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const identifier = String(response.getResponseText() || '').trim();
  if (!identifier) {
    ui.alert('Nothing entered.');
    return;
  }

  const result = withScriptLock_(function () {
    return cancelSubmission_(identifier);
  });

  ui.alert(result.message);
}

function archiveCancelledSubmissions() {
  const result = withScriptLock_(function () {
    return archiveCancelledSubmissions_();
  });

  SpreadsheetApp.getUi().alert(result.message);
}


/* ========================================================================== */
/* WEB APP                                                                    */
/* ========================================================================== */

function doGet(e) {
  try {
    const action = normalizeText_(e && e.parameter && e.parameter.action) || 'capacity';
    let result;

    if (action === 'ping') {
      result = {
        ok: true,
        service: 'IMC Connects submission backend',
        backendVersion: CONFIG.BACKEND_VERSION,
        now: new Date().toISOString()
      };
    } else if (action === 'capacity') {
      result = capacityResponse_();
    } else if (action === 'submissionfaststatus') {
      result = submissionFastStatusResponse_(
        cleanSingleLine_(e && e.parameter && e.parameter.token)
      );
    } else if (action === 'submissionstatus') {
      result = submissionStatusResponse_(
        cleanSingleLine_(e && e.parameter && e.parameter.token),
        cleanSingleLine_(e && e.parameter && e.parameter.email)
      );
    } else {
      result = {
        ok: false,
        error: 'UNKNOWN_ACTION',
        message: 'Unknown GET action.'
      };
    }

    // Browser reads use JSONP because Apps Script ContentService responses
    // are cross-origin. Plain JSON remains available for direct/manual tests.
    return respondGet_(e, result);
  } catch (err) {
    console.error(err);
    return respondGet_(e, {
      ok: false,
      error: 'SERVER_ERROR',
      message: String(err && err.message || err)
    });
  }
}

function doPost(e) {
  let payload = {};
  let requestToken = cleanSingleLine_(
    e && e.parameter && e.parameter.responseToken
  );

  try {
    payload = parseRequestBody_(e);
    requestToken = requestToken || cleanSingleLine_(payload.clientToken);

    const action = normalizeText_(payload.action);
    let result;

    if (action === 'submit') {
      result = submitRequest_(payload);
    } else if (action === 'healthcheck') {
      result = {
        ok: true,
        backendVersion: CONFIG.BACKEND_VERSION,
        postTransport: true,
        checkedAt: new Date().toISOString()
      };
    } else if (action === 'feedback') {
      result = saveFeedback_(payload);
    } else {
      result = {
        ok: false,
        error: 'UNKNOWN_ACTION',
        message: 'Unknown POST action.'
      };
    }

    rememberRequestResult_(requestToken, result);
    return respondPost_(e, payload, result);
  } catch (err) {
    console.error(err);

    const result = {
      ok: false,
      error: 'SERVER_ERROR',
      message: String(err && err.message || err)
    };

    rememberRequestResult_(requestToken, result);
    return respondPost_(e, payload, result);
  }
}


/* ========================================================================== */
/* SUBMISSION TRANSACTION                                                     */
/* ========================================================================== */

function submitRequest_(payload) {
  const studentName = cleanSingleLine_(payload.name || (payload.student && payload.student.name));
  const studentEmail = cleanSingleLine_(payload.email || (payload.student && payload.student.email));
  const normalizedEmail = normalizeEmail_(studentEmail);
  const clientToken = cleanSingleLine_(payload.clientToken);

  if (!studentName) {
    return fail_('INVALID_NAME', 'Please enter your name.');
  }
  if (!isValidEmail_(normalizedEmail)) {
    return fail_('INVALID_EMAIL', 'Please enter a valid email address.');
  }
  if (!clientToken || clientToken.length < 12) {
    return fail_('INVALID_CLIENT_TOKEN', 'Submission token is missing. Refresh the page and try again.');
  }

  const requestedGuaranteed = arrayOrEmpty_(payload.guaranteed);
  const requestedWishlist = arrayOrEmpty_(payload.wishlist);

  if (requestedGuaranteed.length < 1 || requestedGuaranteed.length > CONFIG.MAX_GUARANTEED) {
    return fail_(
      'INVALID_SELECTION_COUNT',
      'Choose between 1 and ' + CONFIG.MAX_GUARANTEED + ' guaranteed alumni.'
    );
  }

  if (requestedGuaranteed.length + requestedWishlist.length > CONFIG.MAX_MY_LIST) {
    return fail_(
      'MY_LIST_TOO_LARGE',
      'My List can contain up to ' + CONFIG.MAX_MY_LIST + ' alumni total.'
    );
  }

  // New alumni can appear in the public directory before the organizer has
  // manually refreshed Alumni Capacity. If any guaranteed selection is
  // missing, refresh the roster once BEFORE entering the reservation lock,
  // then continue the same submission.
  try {
    ensureRequestedAlumniSynced_(requestedGuaranteed);
  } catch (err) {
    return fail_(
      'ALUMNI_SYNC_FAILED',
      'The alumni roster could not be refreshed. No spots were reserved. Please try again.',
      { detail: String(err && err.message || err) }
    );
  }

  let transactionResult;

  try {
    transactionResult = withScriptLock_(function () {
      const ss = getDb_();
      const submissionSheet = ss.getSheetByName(CONFIG.SHEETS.SUBMISSIONS);
      const capacitySheet = ss.getSheetByName(CONFIG.SHEETS.CAPACITY);
      const reservationSheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);

      ensureWishlistColumns_(submissionSheet, requestedWishlist.length);

      // Capacity rows are maintained incrementally by every successful
      // reservation/cancellation. The final read still occurs inside the
      // shared script lock, so no two submissions can claim the same last spot.
      if (emailHasSubmitted_(submissionSheet, normalizedEmail)) {
        return fail_(
          'DUPLICATE_EMAIL',
          'A submission already exists for this email address. If you need to change it, reply to your confirmation email with your updated request.'
        );
      }

      const ledger = readCapacityLedger_(capacitySheet);

      let guaranteed;
      try {
        guaranteed = requestedGuaranteed.map(function (item) {
          return resolveRequestedAlum_(item, ledger);
        });
      } catch (err) {
        return fail_('ALUMNI_NOT_FOUND', err.message);
      }

      const uniqueKeys = new Set(guaranteed.map(function (a) { return a.alumniKey; }));
      if (uniqueKeys.size !== guaranteed.length) {
        return fail_('DUPLICATE_SELECTION', 'The same alum was selected more than once.');
      }

      // Final source-of-truth validation: count ACTIVE reservations for only
      // the selected alumni while holding the same global script lock.
      // This means a stale/manual edit to derived Spots Remaining values
      // cannot cause an overbooking. Reservations + Capacity Limit are the
      // authoritative inputs for the final transaction.
      const authoritativeReserved = countActiveReservationsForKeys_(
        reservationSheet,
        uniqueKeys
      );

      guaranteed.forEach(function (a) {
        const reserved = authoritativeReserved[a.alumniKey] || 0;
        const limit = Math.max(0, Number(a.capacityLimit) || 0);
        a.spotsReserved = reserved;
        a.spotsRemaining = Math.max(0, limit - reserved);
        a.overbookedBy = Math.max(0, reserved - limit);
      });

      const unavailable = guaranteed
        .filter(function (a) {
          return !a.sourceActive || Number(a.spotsRemaining) < 1;
        })
        .map(function (a) {
          return {
            alumniKey: a.alumniKey,
            name: a.name,
            company: a.company,
            spotsRemaining: Math.max(0, Number(a.spotsRemaining) || 0)
          };
        });

      if (unavailable.length) {
        return {
          ok: false,
          error: 'CAPACITY_CHANGED',
          message: 'One or more selected alumni just filled up. No spots were reserved.',
          unavailable: unavailable
        };
      }

      let wishlistResolved = [];
      requestedWishlist.forEach(function (item) {
        try {
          const alum = resolveRequestedAlum_(item, ledger);
          if (!uniqueKeys.has(alum.alumniKey) &&
              !wishlistResolved.some(function (x) { return x.alumniKey === alum.alumniKey; })) {
            wishlistResolved.push(alum);
          }
        } catch (err) {
          // Wishlist never reserves capacity. If a stale wishlist item no
          // longer exists in the source, keep a readable best-effort record.
          const fallbackName = cleanSingleLine_(item && item.name);
          if (fallbackName) {
            wishlistResolved.push({
              alumniKey: cleanSingleLine_(item && item.alumniKey),
              name: fallbackName,
              company: cleanSingleLine_(item && item.company)
            });
          }
        }
      });

      const now = new Date();
      const submissionId = makeId_('SUB');
      let submissionRow = 0;

      try {
        submissionRow = appendSubmission_(
          submissionSheet,
          now,
          submissionId,
          clientToken,
          'PENDING',
          studentName,
          studentEmail,
          normalizedEmail,
          guaranteed,
          wishlistResolved
        );

        const reservationRows = guaranteed.map(function (alum) {
          return [
            makeId_('RES'),
            now,
            submissionId,
            studentName,
            studentEmail,
            alum.alumniKey,
            alum.name,
            alum.company,
            'ACTIVE',
            ''
          ];
        });

        appendRows_(reservationSheet, reservationRows);

        // Only touch the 1–3 capacity rows involved in this transaction.
        // This keeps the global lock short even when many students submit.
        reserveSelectedCapacityRows_(capacitySheet, guaranteed);

        setCellByHeader_(submissionSheet, submissionRow, 'Status', 'ACTIVE');
        SpreadsheetApp.flush();

        // Publish the fresh shared snapshot while we already have the ledger
        // in memory; capacity GETs can now avoid rereading the sheet.
        cacheCapacityPayload_(capacityPayloadFromLedger_(ledger));

        return {
          ok: true,
          submissionId: submissionId,
          submittedAt: now.toISOString(),
          guaranteed: guaranteed.map(function (a) {
            return {
              alumniKey: a.alumniKey,
              name: a.name,
              company: a.company,
              spotsRemaining: Math.max(0, Number(a.spotsRemaining) || 0)
            };
          }),
          wishlist: wishlistResolved.map(function (a) {
            return {
              alumniKey: a.alumniKey || '',
              name: a.name || '',
              company: a.company || ''
            };
          })
        };
      } catch (err) {
        // Best-effort rollback: a partial failure must never intentionally
        // leave ACTIVE reservations consuming capacity.
        if (submissionRow) {
          setCellByHeader_(submissionSheet, submissionRow, 'Status', 'ERROR');
          setCellByHeader_(submissionSheet, submissionRow, 'Error', String(err && err.message || err));
        }
        cancelReservationsForSubmission_(reservationSheet, submissionId, 'SYSTEM_ROLLBACK');
        recalculateSpots_();
        invalidateCapacityCache_();
        rebuildGuaranteedMatches_();
        throw err;
      }
    });
  } catch (err) {
    return fail_('SUBMISSION_FAILED', 'Submission failed. No new reservation should be assumed. Please try again.', {
      detail: String(err && err.message || err)
    });
  }

  if (!transactionResult.ok) return transactionResult;

  // The reservation transaction is now complete and ACTIVE. Publish that
  // success to the request-status cache immediately so the browser can move
  // to the success screen without waiting for operational housekeeping or
  // MailApp. doPost() will overwrite this cache entry later with the final
  // receipt-email status.
  rememberRequestResult_(clientToken, Object.assign({}, transactionResult, {
    receiptEmailSent: null,
    reservationCommitted: true
  }));

  // Operational output is deliberately OUTSIDE the reservation lock. The
  // authoritative records are Submissions + Reservations; this sheet is a
  // regenerated convenience view and should not make the student wait.
  try {
    rebuildGuaranteedMatches_();
  } catch (err) {
    console.error('Could not rebuild Guaranteed Matches after submission: ' + err);
  }

  // Email is also outside the capacity lock. A mail quota or latency problem
  // must never undo a successfully reserved spot or delay browser confirmation.
  const emailResult = sendReceipt_(
    transactionResult.submissionId,
    studentName,
    studentEmail,
    transactionResult.guaranteed,
    transactionResult.wishlist
  );

  updateReceiptStatus_(
    transactionResult.submissionId,
    emailResult.ok ? 'SENT' : 'FAILED: ' + emailResult.message,
    emailResult.ok ? new Date() : ''
  );

  transactionResult.receiptEmailSent = emailResult.ok;
  if (!emailResult.ok) {
    transactionResult.receiptEmailWarning =
      'Your selections were reserved, but the confirmation email could not be sent.';
  }

  return transactionResult;
}


function submissionFastStatusResponse_(clientToken) {
  if (!clientToken) {
    return {
      ok: false,
      backendVersion: CONFIG.BACKEND_VERSION,
      error: 'MISSING_IDENTITY',
      message: 'Submission token is required.'
    };
  }

  // Fast path used only while the browser is waiting for a just-submitted
  // request. It checks the exact request-result cache and deliberately does
  // NOT scan the Submissions sheet when the transaction is still in flight.
  // Once submitRequest_ commits the ACTIVE reservation it publishes success
  // to this cache immediately, before email and operational housekeeping.
  const requestResult = readRequestResult_(clientToken);
  if (requestResult) {
    return {
      ok: true,
      backendVersion: CONFIG.BACKEND_VERSION,
      found: true,
      requestComplete: true,
      result: requestResult
    };
  }

  return {
    ok: true,
    backendVersion: CONFIG.BACKEND_VERSION,
    found: false,
    requestComplete: false
  };
}


function submissionStatusResponse_(clientToken, email) {
  const normalizedEmail = normalizeEmail_(email);

  if (!clientToken && !normalizedEmail) {
    return {
      ok: false,
      error: 'MISSING_IDENTITY',
      message: 'Submission token or email is required.'
    };
  }

  // First preference: the exact terminal result from the matching POST.
  // This captures backend rejections such as DUPLICATE_EMAIL and
  // CAPACITY_CHANGED, not just successful writes.
  if (clientToken) {
    const requestResult = readRequestResult_(clientToken);
    if (requestResult) {
      return {
        ok: true,
        backendVersion: CONFIG.BACKEND_VERSION,
        found: true,
        requestComplete: true,
        result: requestResult
      };
    }
  }

  const ss = getDb_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SUBMISSIONS);
  const rows = readObjects_(sheet);

  // Success fallback: the transaction writes the ACTIVE submission before
  // the receipt email is sent, so the browser can confirm success even while
  // MailApp is still finishing.
  let match = null;

  if (clientToken) {
    match = rows.find(function (row) {
      return cleanSingleLine_(row['Client Token']) === clientToken;
    }) || null;
  }

  if (!match && normalizedEmail) {
    const emailMatches = rows.filter(function (row) {
      const rowEmail = normalizeEmail_(row['Normalized Email'] || row['Student Email']);
      const status = String(row.Status || '').trim().toUpperCase();
      return rowEmail === normalizedEmail && status === 'ACTIVE';
    });

    if (emailMatches.length) {
      match = emailMatches[emailMatches.length - 1];
    }
  }

  if (!match) {
    return {
      ok: true,
      backendVersion: CONFIG.BACKEND_VERSION,
      found: false,
      requestComplete: false
    };
  }

  return {
    ok: true,
    backendVersion: CONFIG.BACKEND_VERSION,
    found: true,
    requestComplete: false,
    status: String(match.Status || '').trim().toUpperCase(),
    submissionId: cleanSingleLine_(match['Submission ID']),
    receiptEmailStatus: cleanSingleLine_(match['Receipt Email Status'])
  };
}


/* ========================================================================== */
/* FEEDBACK                                                                   */
/* ========================================================================== */

function saveFeedback_(payload) {
  const submissionId = cleanSingleLine_(payload.submissionId);
  const email = cleanSingleLine_(payload.email);
  const normalizedEmail = normalizeEmail_(email);
  const rating = Number(payload.rating);
  const feedback = String(payload.feedback || '').trim();

  if (!submissionId && !normalizedEmail) {
    return fail_('MISSING_IDENTITY', 'Submission ID or email is required.');
  }
  if (!Number.isInteger(rating) || rating < 1 || rating > 5) {
    return fail_('INVALID_RATING', 'Rating must be an integer from 1 to 5.');
  }

  const ss = getDb_();
  // Feedback should never depend on an organizer having remembered to run
  // setupSheets() after a backend upgrade. Create the dedicated sheet on
  // demand if it is missing; the normal path remains just one sheet lookup.
  let sheet = ss.getSheetByName(CONFIG.SHEETS.FEEDBACK);
  if (!sheet) {
    sheet = ensureSheet_(ss, CONFIG.SHEETS.FEEDBACK, FEEDBACK_HEADERS);
  }

  appendRows_(sheet, [[
    new Date(),
    makeId_('FDB'),
    submissionId,
    email,
    normalizedEmail,
    rating,
    feedback
  ]]);

  return {
    ok: true,
    message: 'Feedback saved.'
  };
}


/* ========================================================================== */
/* ALUMNI SYNC + CAPACITY                                                     */
/* ========================================================================== */

function syncAlumni_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CAPACITY);

    const response = UrlFetchApp.fetch(CONFIG.ALUMNI_CSV_URL, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    const status = response.getResponseCode();
    if (status < 200 || status >= 300) {
      throw new Error('Could not fetch alumni CSV. HTTP ' + status);
    }

    const csv = response.getContentText();
    const rows = Utilities.parseCsv(csv);
    if (!rows.length) throw new Error('Alumni CSV is empty.');

    const headers = rows[0].map(String);
    const required = [
      SOURCE_HEADERS.TIMESTAMP,
      SOURCE_HEADERS.NAME,
      SOURCE_HEADERS.COMPANY,
      SOURCE_HEADERS.CAPACITY
    ];

    required.forEach(function (header) {
      if (headers.indexOf(header) === -1) {
        throw new Error('Missing required alumni CSV column: ' + header);
      }
    });

    const source = rows.slice(1)
      .filter(function (row) {
        return row.some(function (cell) { return String(cell || '').trim() !== ''; });
      })
      .map(function (row) {
        const obj = {};
        headers.forEach(function (h, i) { obj[h] = row[i] || ''; });

        const timestamp = cleanSingleLine_(obj[SOURCE_HEADERS.TIMESTAMP]);
        const name = cleanSingleLine_(obj[SOURCE_HEADERS.NAME]);
        const company = cleanSingleLine_(obj[SOURCE_HEADERS.COMPANY]);
        const role = cleanSingleLine_(obj[SOURCE_HEADERS.ROLE]);
        const parsed = parseInt(String(obj[SOURCE_HEADERS.CAPACITY] || '').trim(), 10);
        const startingCapacity = Number.isFinite(parsed)
          ? Math.max(0, parsed)
          : CONFIG.ASSUMED_CAPACITY_IF_BLANK;

        return {
          alumniKey: makeAlumniKey_(timestamp, name, company),
          name: name,
          company: company,
          role: role,
          timestamp: timestamp,
          startingCapacity: startingCapacity
        };
      })
      .filter(function (a) { return a.name; });

    const existing = readCapacityLedger_(sheet);
    const byKey = existing.byKey;
    const now = new Date();

    // Preserve rows that disappeared from the source for historical integrity,
    // but mark them inactive so the website will not offer them.
    Object.keys(byKey).forEach(function (key) {
      byKey[key].sourceActive = false;
    });

    source.forEach(function (a) {
      const prior = byKey[a.alumniKey];

      if (prior) {
        const oldStarting = Number(prior.startingCapacity) || 0;
        const oldLimit = Number(prior.capacityLimit);

        // If the organizer never customized Capacity Limit, let a source
        // capacity correction flow through. If they DID customize it, preserve it.
        const nextLimit =
          Number.isFinite(oldLimit) && oldLimit !== oldStarting
            ? Math.max(0, oldLimit)
            : a.startingCapacity;

        prior.name = a.name;
        prior.company = a.company;
        prior.role = a.role;
        prior.sourceTimestamp = a.timestamp;
        prior.startingCapacity = a.startingCapacity;
        prior.capacityLimit = nextLimit;
        prior.sourceActive = true;
        prior.lastSynced = now;
      } else {
        byKey[a.alumniKey] = {
          alumniKey: a.alumniKey,
          name: a.name,
          company: a.company,
          role: a.role,
          sourceTimestamp: a.timestamp,
          startingCapacity: a.startingCapacity,
          capacityLimit: a.startingCapacity,
          spotsReserved: 0,
          spotsRemaining: a.startingCapacity,
          overbookedBy: 0,
          sourceActive: true,
          lastSynced: now
        };
      }
    });

    const output = Object.keys(byKey)
      .map(function (key) { return byKey[key]; })
      .sort(function (a, b) {
        return String(a.name || '').localeCompare(String(b.name || ''));
      })
      .map(function (a) {
        return [
          a.alumniKey,
          a.name,
          a.company,
          a.role || '',
          a.sourceTimestamp || '',
          a.startingCapacity,
          a.capacityLimit,
          a.spotsReserved || 0,
          a.spotsRemaining || 0,
          a.overbookedBy || 0,
          Boolean(a.sourceActive),
          a.lastSynced || now
        ];
      });

    replaceDataRows_(sheet, output, CAPACITY_HEADERS.length);
    formatCapacitySheet_(sheet);
    invalidateCapacityCache_();
  } finally {
    lock.releaseLock();
  }
}

function ensureRequestedAlumniSynced_(requestedGuaranteed) {
  if (!requestedGuaranteed || !requestedGuaranteed.length) return false;

  const ss = getDb_();
  const capacitySheet = ss.getSheetByName(CONFIG.SHEETS.CAPACITY);
  const ledger = readCapacityLedger_(capacitySheet);

  const needsSync = requestedGuaranteed.some(function (item) {
    try {
      resolveRequestedAlum_(item, ledger);
      return false;
    } catch (err) {
      return true;
    }
  });

  if (!needsSync) return false;

  syncAlumni_();
  recalculateSpots_();
  invalidateCapacityCache_();
  return true;
}

function recalculateSpots_() {
  const ss = getDb_();
  const capacitySheet = ss.getSheetByName(CONFIG.SHEETS.CAPACITY);
  const reservationSheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);

  if (capacitySheet.getLastRow() < 2) return;

  const reservationCounts = {};
  if (reservationSheet.getLastRow() >= 2) {
    const reservationData = readObjects_(reservationSheet);
    reservationData.forEach(function (r) {
      if (String(r.Status || '').toUpperCase() !== 'ACTIVE') return;
      const key = String(r['Alumni Key'] || '');
      if (!key) return;
      reservationCounts[key] = (reservationCounts[key] || 0) + 1;
    });
  }

  const headers = headerMap_(capacitySheet);
  const numRows = capacitySheet.getLastRow() - 1;
  const data = capacitySheet.getRange(2, 1, numRows, capacitySheet.getLastColumn()).getValues();

  data.forEach(function (row) {
    const key = String(row[headers['Alumni Key'] - 1] || '');
    const limit = Math.max(0, Number(row[headers['Capacity Limit'] - 1]) || 0);
    const reserved = reservationCounts[key] || 0;
    const remaining = Math.max(0, limit - reserved);
    const overbooked = Math.max(0, reserved - limit);

    row[headers['Spots Reserved'] - 1] = reserved;
    row[headers['Spots Remaining'] - 1] = remaining;
    row[headers['Overbooked By'] - 1] = overbooked;
  });

  capacitySheet
    .getRange(2, 1, data.length, data[0].length)
    .setValues(data);
}

function countActiveReservationsForKeys_(reservationSheet, keySet) {
  const counts = {};
  if (!keySet || !keySet.size || reservationSheet.getLastRow() < 2) return counts;

  readObjects_(reservationSheet).forEach(function (r) {
    if (String(r.Status || '').trim().toUpperCase() !== 'ACTIVE') return;
    const key = cleanSingleLine_(r['Alumni Key']);
    if (!key || !keySet.has(key)) return;
    counts[key] = (counts[key] || 0) + 1;
  });

  return counts;
}


function reserveSelectedCapacityRows_(capacitySheet, guaranteed) {
  if (!guaranteed.length) return;

  const headers = headerMap_(capacitySheet);
  const reservedCol = headers['Spots Reserved'];
  const remainingCol = headers['Spots Remaining'];
  const overbookedCol = headers['Overbooked By'];
  if (!reservedCol || !remainingCol || !overbookedCol) {
    throw new Error('Capacity sheet is missing reservation columns.');
  }

  guaranteed.forEach(function (alum) {
    const nextReserved = (Number(alum.spotsReserved) || 0) + 1;
    const nextRemaining = Math.max(0, (Number(alum.capacityLimit) || 0) - nextReserved);
    const nextOverbooked = Math.max(0, nextReserved - (Number(alum.capacityLimit) || 0));

    capacitySheet
      .getRange(alum.rowNumber, reservedCol, 1, 3)
      .setValues([[nextReserved, nextRemaining, nextOverbooked]]);

    // Keep the already-read ledger object synchronized so we can publish
    // the fresh cache snapshot without rereading the entire sheet.
    alum.spotsReserved = nextReserved;
    alum.spotsRemaining = nextRemaining;
    alum.overbookedBy = nextOverbooked;
  });
}

function capacityCacheKey_() {
  return 'imc-capacity-snapshot-v1';
}

function capacityPayloadFromLedger_(ledger) {
  const alumni = ledger.rows
    .filter(function (a) { return a.sourceActive; })
    .map(function (a) {
      return {
        alumniKey: a.alumniKey,
        name: a.name,
        company: a.company,
        timestamp: a.sourceTimestamp,
        startingCapacity: Number(a.startingCapacity) || 0,
        capacityLimit: Number(a.capacityLimit) || 0,
        spotsReserved: Number(a.spotsReserved) || 0,
        spotsRemaining: Math.max(0, Number(a.spotsRemaining) || 0),
        fullyBooked: (Number(a.spotsRemaining) || 0) <= 0
      };
    });

  return {
    ok: true,
    backendVersion: CONFIG.BACKEND_VERSION,
    updatedAt: new Date().toISOString(),
    alumni: alumni
  };
}

function cacheCapacityPayload_(payload) {
  try {
    CacheService.getScriptCache().put(
      capacityCacheKey_(),
      JSON.stringify(payload),
      CONFIG.CAPACITY_CACHE_SECONDS
    );
  } catch (err) {
    console.error('Could not cache capacity snapshot: ' + err);
  }
}

function invalidateCapacityCache_() {
  try {
    CacheService.getScriptCache().remove(capacityCacheKey_());
  } catch (err) {
    console.error('Could not invalidate capacity cache: ' + err);
  }
}

function capacityResponse_() {
  try {
    const cached = CacheService.getScriptCache().get(capacityCacheKey_());
    if (cached) return JSON.parse(cached);
  } catch (err) {
    console.error('Could not read capacity cache: ' + err);
  }

  const ss = getDb_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CAPACITY);
  const ledger = readCapacityLedger_(sheet);
  const payload = capacityPayloadFromLedger_(ledger);
  cacheCapacityPayload_(payload);
  return payload;
}


/* ========================================================================== */
/* GUARANTEED MATCHES                                                         */
/* ========================================================================== */

function rebuildGuaranteedMatches_() {
  const ss = getDb_();
  const reservationSheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  const matchesSheet = ss.getSheetByName(CONFIG.SHEETS.MATCHES);

  const rows = readObjects_(reservationSheet)
    .filter(function (r) {
      return String(r.Status || '').toUpperCase() === 'ACTIVE';
    })
    .map(function (r) {
      return [
        r['Submission ID'] || '',
        r['Student Name'] || '',
        r['Student Email'] || '',
        r['Alumni Key'] || '',
        r['Alumni Name'] || '',
        r.Company || '',
        r['Reserved At'] || ''
      ];
    });

  replaceDataRows_(matchesSheet, rows, MATCH_HEADERS.length);
}


/* ========================================================================== */
/* CANCELLATION                                                               */
/* ========================================================================== */

function cancelSubmission_(identifier) {
  const ss = getDb_();
  const submissionSheet = ss.getSheetByName(CONFIG.SHEETS.SUBMISSIONS);
  const reservationSheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  const archiveSheet = ensureArchiveSheet_(ss, submissionSheet);

  const normalizedIdentifier = normalizeEmail_(identifier);
  const objects = readObjectsWithRow_(submissionSheet);

  const matches = objects.filter(function (item) {
    const row = item.data;
    return String(row['Submission ID'] || '').trim() === identifier ||
      normalizeEmail_(row['Normalized Email'] || row['Student Email']) === normalizedIdentifier;
  });

  if (!matches.length) {
    const archivedMatch = readObjects_(archiveSheet).some(function (row) {
      return String(row['Submission ID'] || '').trim() === identifier ||
        normalizeEmail_(row['Normalized Email'] || row['Student Email']) === normalizedIdentifier;
    });

    if (archivedMatch) {
      return {
        ok: true,
        message: 'That submission was already cancelled and moved to Test Submissions Archive. The email is free to submit again.'
      };
    }

    return {
      ok: false,
      message: 'No submission found for: ' + identifier
    };
  }

  const active = matches.filter(function (item) {
    const status = String(item.data.Status || '').trim().toUpperCase();
    return status === 'ACTIVE' || status === 'PENDING';
  });

  if (!active.length) {
    const cancelledRows = matches.filter(function (item) {
      return String(item.data.Status || '').trim().toUpperCase() === 'CANCELLED';
    });

    if (cancelledRows.length) {
      archiveSubmissionRows_(
        submissionSheet,
        archiveSheet,
        cancelledRows.map(function (item) { return item.rowNumber; })
      );
    }

    return {
      ok: true,
      message: 'That submission was already cancelled and is now archived. The email is free to submit again.'
    };
  }

  const now = new Date();
  active.forEach(function (item) {
    setCellByHeader_(submissionSheet, item.rowNumber, 'Status', 'CANCELLED');
    setCellByHeader_(submissionSheet, item.rowNumber, 'Cancelled At', now);
    cancelReservationsForSubmission_(
      reservationSheet,
      String(item.data['Submission ID'] || ''),
      'ORGANIZER_CANCELLED'
    );
  });

  recalculateSpots_();
  invalidateCapacityCache_();
  rebuildGuaranteedMatches_();

  archiveSubmissionRows_(
    submissionSheet,
    archiveSheet,
    active.map(function (item) { return item.rowNumber; })
  );

  return {
    ok: true,
    message:
      'Cancelled and archived ' + active.length + ' submission(s), and restored their active reserved capacity.\n\n' +
      'The email can now submit again.'
  };
}

function archiveCancelledSubmissions_() {
  const ss = getDb_();
  const submissionSheet = ss.getSheetByName(CONFIG.SHEETS.SUBMISSIONS);
  const archiveSheet = ensureArchiveSheet_(ss, submissionSheet);

  const cancelled = readObjectsWithRow_(submissionSheet).filter(function (item) {
    return String(item.data.Status || '').trim().toUpperCase() === 'CANCELLED';
  });

  if (!cancelled.length) {
    return {
      ok: true,
      message: 'No cancelled submissions are waiting to be archived.'
    };
  }

  archiveSubmissionRows_(
    submissionSheet,
    archiveSheet,
    cancelled.map(function (item) { return item.rowNumber; })
  );

  return {
    ok: true,
    message: 'Moved ' + cancelled.length + ' cancelled submission(s) to Test Submissions Archive.'
  };
}

function ensureArchiveSheet_(ss, submissionSheet) {
  const headers = getHeaders_(submissionSheet);
  let archiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);
  if (!archiveSheet) archiveSheet = ss.insertSheet(CONFIG.SHEETS.ARCHIVE);

  if (archiveSheet.getMaxColumns() < headers.length) {
    archiveSheet.insertColumnsAfter(
      archiveSheet.getMaxColumns(),
      headers.length - archiveSheet.getMaxColumns()
    );
  }

  archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRange_(archiveSheet.getRange(1, 1, 1, headers.length));
  archiveSheet.setFrozenRows(1);
  return archiveSheet;
}

function archiveSubmissionRows_(submissionSheet, archiveSheet, rowNumbers) {
  const rowsToMove = Array.from(new Set(rowNumbers || []))
    .filter(function (rowNumber) { return Number(rowNumber) >= 2; })
    .sort(function (a, b) { return a - b; });

  if (!rowsToMove.length) return 0;

  archiveSheet = archiveSheet || ensureArchiveSheet_(getDb_(), submissionSheet);
  const width = submissionSheet.getLastColumn();
  const values = rowsToMove.map(function (rowNumber) {
    return submissionSheet.getRange(rowNumber, 1, 1, width).getValues()[0];
  });

  if (values.length) {
    archiveSheet
      .getRange(archiveSheet.getLastRow() + 1, 1, values.length, width)
      .setValues(values);
  }

  rowsToMove
    .slice()
    .sort(function (a, b) { return b - a; })
    .forEach(function (rowNumber) {
      submissionSheet.deleteRow(rowNumber);
    });

  return values.length;
}

function cancelReservationsForSubmission_(sheet, submissionId, reason) {
  if (!submissionId || sheet.getLastRow() < 2) return;

  const headers = headerMap_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  let changed = false;

  values.forEach(function (row) {
    const rowSubmissionId = String(row[headers['Submission ID'] - 1] || '');
    const status = String(row[headers.Status - 1] || '').toUpperCase();

    if (rowSubmissionId === submissionId && status === 'ACTIVE') {
      row[headers.Status - 1] = reason || 'CANCELLED';
      row[headers['Cancelled At'] - 1] = new Date();
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  }
}


/* ========================================================================== */
/* EMAIL RECEIPT                                                              */
/* ========================================================================== */

function sendReceipt_(submissionId, studentName, studentEmail, guaranteed, wishlist) {
  try {
    const firstName = cleanSingleLine_(studentName).split(/\s+/)[0] || 'there';

    const guaranteedLines = guaranteed.map(function (a, i) {
      return (i + 1) + '. ' + a.name + (a.company ? ' - ' + a.company : '');
    });
    const wishlistLines = wishlist.map(function (a, i) {
      return (i + 1) + '. ' + a.name + (a.company ? ' - ' + a.company : '');
    });

    const body =
      'Hi ' + firstName + ',\n\n' +
      'Your alumni coffee chat requests have been received.\n\n' +
      'Top 3\n' +
      guaranteedLines.join('\n') + '\n\n' +
      'Wish List\n' +
      (wishlistLines.length ? wishlistLines.join('\n') : 'None') + '\n\n' +
      'Submission ID: ' + submissionId + '\n\n' +
      'You will receive a separate introduction email for each alum as the matching process is completed.\n\n' +
      'If you need to change something, please reply to this email with your updated request.\n\n' +
      'Cheers,\n\n' +
      'Gaby\n' +
      'IMC Connects';

    const htmlGuaranteedList = guaranteed.map(function (a, i) {
      return '<div style="margin:0 0 4px;">' +
        (i + 1) + '. ' +
        escapeHtml_(a.name) +
        (a.company ? ' - ' + escapeHtml_(a.company) : '') +
        '</div>';
    }).join('');

    const htmlWishlistList = wishlist.length
      ? wishlist.map(function (a, i) {
          return '<div style="margin:0 0 4px;">' +
            (i + 1) + '. ' +
            escapeHtml_(a.name) +
            (a.company ? ' - ' + escapeHtml_(a.company) : '') +
            '</div>';
        }).join('')
      : '<div style="margin:0;">None</div>';

    const htmlBody =
      '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;color:#202124;">' +
        '<p style="margin:0 0 20px;">Hi ' + escapeHtml_(firstName) + ',</p>' +
        '<p style="margin:0 0 20px;">Your alumni coffee chat requests have been received.</p>' +
        '<div style="margin:0 0 18px;">' +
          '<div style="font-weight:700;margin:0 0 6px;">Top 3</div>' +
          htmlGuaranteedList +
        '</div>' +
        '<div style="margin:0 0 20px;">' +
          '<div style="font-weight:700;margin:0 0 6px;">Wish List</div>' +
          htmlWishlistList +
        '</div>' +
        '<p style="margin:0 0 20px;">Submission ID: ' + escapeHtml_(submissionId) + '</p>' +
        '<p style="margin:0 0 20px;">You will receive a separate introduction email for each alum as the matching process is completed.</p>' +
        '<p style="margin:0 0 20px;">If you need to change something, please reply to this email with your updated request.</p>' +
        '<p style="margin:0 0 2px;">Cheers,</p>' +
        '<p style="margin:0 0 2px;">Gaby</p>' +
        '<p style="margin:0;color:#8a8a8a;">IMC Connects</p>' +
      '</div>';

    MailApp.sendEmail({
      to: studentEmail,
      subject: 'IMC Connects: Submission Successful',
      body: body,
      htmlBody: htmlBody,
      name: 'IMC Connects'
    });

    return { ok: true, message: 'sent' };
  } catch (err) {
    return {
      ok: false,
      message: String(err && err.message || err)
    };
  }
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function updateReceiptStatus_(submissionId, status, sentAt) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const ss = getDb_();
      const sheet = ss.getSheetByName(CONFIG.SHEETS.SUBMISSIONS);
      const rows = readObjectsWithRow_(sheet);
      const match = rows.find(function (item) {
        return String(item.data['Submission ID'] || '') === submissionId;
      });

      if (!match) return;

      setCellByHeader_(sheet, match.rowNumber, 'Receipt Email Status', status);
      setCellByHeader_(sheet, match.rowNumber, 'Receipt Email Sent At', sentAt || '');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error('Could not update receipt status: ' + err);
  }
}


/* ========================================================================== */
/* SHEET WRITE HELPERS                                                        */
/* ========================================================================== */

function appendSubmission_(
  sheet,
  timestamp,
  submissionId,
  clientToken,
  status,
  studentName,
  studentEmail,
  normalizedEmail,
  guaranteed,
  wishlist
) {
  ensureWishlistColumns_(sheet, wishlist.length);

  const headers = getHeaders_(sheet);
  const row = new Array(headers.length).fill('');
  const map = {};
  headers.forEach(function (h, i) { map[h] = i; });

  function put(header, value) {
    if (Object.prototype.hasOwnProperty.call(map, header)) {
      row[map[header]] = value;
    }
  }

  put('Timestamp', timestamp);
  put('Submission ID', submissionId);
  put('Client Token', clientToken);
  put('Status', status);
  put('Student Name', studentName);
  put('Student Email', studentEmail);
  put('Normalized Email', normalizedEmail);

  for (let i = 0; i < CONFIG.MAX_GUARANTEED; i++) {
    const a = guaranteed[i];
    if (!a) continue;
    put('Alum ' + (i + 1) + ' Key', a.alumniKey);
    put('Alum ' + (i + 1), a.name);
    put('Alum ' + (i + 1) + ' Company', a.company);
  }

  wishlist.forEach(function (a, i) {
    put('Wishlist ' + (i + 1), a.name || '');
  });

  const rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  return rowNumber;
}

function ensureWishlistColumns_(sheet, needed) {
  const desired = Math.max(CONFIG.INITIAL_WISHLIST_COLUMNS, Number(needed) || 0);
  const headers = getHeaders_(sheet);
  const existing = headers.filter(function (h) {
    return /^Wishlist \d+$/.test(String(h));
  }).length;

  if (existing >= desired) return;

  const newHeaders = [];
  for (let i = existing + 1; i <= desired; i++) {
    newHeaders.push('Wishlist ' + i);
  }

  if (newHeaders.length) {
    const startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, newHeaders.length).setValues([newHeaders]);
    styleHeaderRange_(sheet.getRange(1, startCol, 1, newHeaders.length));
  }
}

function emailHasSubmitted_(sheet, normalizedEmail) {
  return readObjects_(sheet).some(function (row) {
    const email = normalizeEmail_(row['Normalized Email'] || row['Student Email']);
    const status = String(row.Status || '').trim().toUpperCase();

    // Only a live submission blocks the email. Once the organizer cancels
    // a submission, its capacity is restored and that student may submit again.
    return email === normalizedEmail && (status === 'ACTIVE' || status === 'PENDING');
  });
}

function setCellByHeader_(sheet, rowNumber, header, value) {
  const map = headerMap_(sheet);
  const col = map[header];
  if (!col) throw new Error('Missing sheet header "' + header + '" in ' + sheet.getName());
  sheet.getRange(rowNumber, col).setValue(value);
}

function appendRows_(sheet, rows) {
  if (!rows || !rows.length) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function replaceDataRows_(sheet, rows, width) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), width);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, width).setValues(rows);
  }
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existing = getHeaders_(sheet);

    // Do not destructively reorder a populated sheet. Add missing headers.
    headers.forEach(function (header) {
      if (existing.indexOf(header) === -1) {
        const col = sheet.getLastColumn() + 1;
        sheet.getRange(1, col).setValue(header);
        existing.push(header);
      }
    });
  }

  styleHeaderRange_(sheet.getRange(1, 1, 1, sheet.getLastColumn()));
  sheet.setFrozenRows(1);
  return sheet;
}


/**
 * Removes legacy Alum 4 / Alum 5 guaranteed-selection columns from an
 * EMPTY Submissions sheet. The live product supports exactly 3 guaranteed
 * alumni. This cleanup is deliberately conservative: once submission data
 * exists, it will not delete columns automatically.
 */
function normalizeGuaranteedColumnsToThree_(sheet) {
  if (!sheet) return;

  const headers = getHeaders_(sheet);
  const legacyCols = [];

  headers.forEach(function (header, index) {
    if (/^Alum [4-9](?: Key| Company)?$/.test(String(header || '').trim())) {
      legacyCols.push(index + 1);
    }
  });

  if (!legacyCols.length) return;

  if (sheet.getLastRow() > 1) {
    throw new Error(
      'The Submissions sheet still contains legacy Alum 4/5 columns and also has data. ' +
      'For safety, the script did not delete them automatically.'
    );
  }

  legacyCols
    .sort(function (a, b) { return b - a; })
    .forEach(function (col) {
      sheet.deleteColumn(col);
    });
}

/**
 * The live product now caps My List at 8 total alumni. Since at least one
 * guaranteed selection is required, the largest possible wishlist is 7.
 *
 * Older script versions created Wishlist 1-10. Remove legacy Wishlist 8+
 * columns only when those columns contain no data, so setup remains safe
 * on a sheet that has already been used.
 */
function normalizeWishlistColumnsToSeven_(sheet) {
  if (!sheet) return;

  const headers = getHeaders_(sheet);
  const legacyCols = [];

  headers.forEach(function (header, index) {
    const match = String(header || '').trim().match(/^Wishlist (\d+)$/);
    if (match && Number(match[1]) > 7) {
      legacyCols.push(index + 1);
    }
  });

  if (!legacyCols.length) return;

  const deletable = legacyCols.filter(function (col) {
    if (sheet.getLastRow() < 2) return true;
    const values = sheet
      .getRange(2, col, sheet.getLastRow() - 1, 1)
      .getDisplayValues()
      .flat();

    return values.every(function (value) {
      return String(value || '').trim() === '';
    });
  });

  deletable
    .sort(function (a, b) { return b - a; })
    .forEach(function (col) {
      sheet.deleteColumn(col);
    });
}

function submissionHeaders_(wishlistCount) {
  const headers = SUBMISSION_CORE_HEADERS.slice();
  for (let i = 1; i <= wishlistCount; i++) {
    headers.push('Wishlist ' + i);
  }
  return headers;
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() < 1) return [];
  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(function (x) { return String(x || '').trim(); });
}

function headerMap_(sheet) {
  const map = {};
  getHeaders_(sheet).forEach(function (header, i) {
    if (header) map[header] = i + 1;
  });
  return map;
}

function readObjects_(sheet) {
  return readObjectsWithRow_(sheet).map(function (x) { return x.data; });
}

function readObjectsWithRow_(sheet) {
  if (sheet.getLastRow() < 2) return [];

  const headers = getHeaders_(sheet);
  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getValues();

  return values.map(function (row, index) {
    const obj = {};
    headers.forEach(function (header, i) {
      if (header) obj[header] = row[i];
    });
    return {
      rowNumber: index + 2,
      data: obj
    };
  });
}


/* ========================================================================== */
/* LEDGER READ + REQUEST RESOLUTION                                           */
/* ========================================================================== */

function readCapacityLedger_(sheet) {
  const rows = readObjectsWithRow_(sheet).map(function (item) {
    const r = item.data;
    return {
      rowNumber: item.rowNumber,
      alumniKey: cleanSingleLine_(r['Alumni Key']),
      name: cleanSingleLine_(r.Name),
      company: cleanSingleLine_(r.Company),
      role: cleanSingleLine_(r.Role),
      sourceTimestamp: cleanSingleLine_(r['Source Timestamp']),
      startingCapacity: Number(r['Starting Capacity']) || 0,
      capacityLimit: Number(r['Capacity Limit']) || 0,
      spotsReserved: Number(r['Spots Reserved']) || 0,
      spotsRemaining: Number(r['Spots Remaining']) || 0,
      overbookedBy: Number(r['Overbooked By']) || 0,
      sourceActive: toBoolean_(r['Source Active']),
      lastSynced: r['Last Synced'] || ''
    };
  });

  const byKey = {};
  rows.forEach(function (a) {
    if (a.alumniKey) byKey[a.alumniKey] = a;
  });

  return {
    rows: rows,
    byKey: byKey
  };
}

function resolveRequestedAlum_(item, ledger) {
  item = item || {};

  const requestedKey = cleanSingleLine_(item.alumniKey || item.key);
  if (requestedKey && ledger.byKey[requestedKey]) {
    return ledger.byKey[requestedKey];
  }

  const timestamp = cleanSingleLine_(item.timestamp || item.sourceTimestamp);
  const name = cleanSingleLine_(item.name);
  const company = cleanSingleLine_(item.company);

  if (timestamp) {
    const key = makeAlumniKey_(timestamp, name, company);
    if (ledger.byKey[key]) return ledger.byKey[key];
  }

  const nameNorm = normalizeIdentityText_(name);
  const companyNorm = normalizeIdentityText_(company);

  const candidates = ledger.rows.filter(function (a) {
    if (normalizeIdentityText_(a.name) !== nameNorm) return false;
    if (companyNorm && normalizeIdentityText_(a.company) !== companyNorm) return false;
    return true;
  });

  if (candidates.length === 1) return candidates[0];

  if (!candidates.length) {
    throw new Error('Could not find alum "' + (name || requestedKey || 'unknown') + '" in Alumni Capacity.');
  }

  throw new Error(
    'More than one alum matches "' + name + '". The website must send alumniKey for this selection.'
  );
}


/* ========================================================================== */
/* ON-EDIT CAPACITY CONTROL                                                   */
/* ========================================================================== */

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEETS.CAPACITY) return;
    if (e.range.getRow() < 2) return;

    const headers = headerMap_(sheet);
    const limitCol = headers['Capacity Limit'];
    if (!limitCol) return;

    const editedStart = e.range.getColumn();
    const editedEnd = editedStart + e.range.getNumColumns() - 1;
    if (limitCol < editedStart || limitCol > editedEnd) return;

    // Coerce invalid/negative limits to zero.
    const cell = sheet.getRange(e.range.getRow(), limitCol);
    const raw = Number(cell.getValue());
    if (!Number.isFinite(raw) || raw < 0) cell.setValue(0);
    else cell.setValue(Math.floor(raw));

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return;

    try {
      recalculateSpots_();
      invalidateCapacityCache_();
      rebuildGuaranteedMatches_();
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
  }
}


/* ========================================================================== */
/* FORMATTING                                                                 */
/* ========================================================================== */

function formatSheets_(ss) {
  [
    CONFIG.SHEETS.SUBMISSIONS,
    CONFIG.SHEETS.ARCHIVE,
    CONFIG.SHEETS.FEEDBACK,
    CONFIG.SHEETS.CAPACITY,
    CONFIG.SHEETS.RESERVATIONS,
    CONFIG.SHEETS.MATCHES
  ].forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    sheet.setFrozenRows(1);
    styleHeaderRange_(sheet.getRange(1, 1, 1, sheet.getLastColumn()));
    sheet.autoResizeColumns(1, Math.min(sheet.getLastColumn(), 12));
  });
}

function formatCapacitySheet_(sheet) {
  const headers = headerMap_(sheet);

  ['Starting Capacity', 'Capacity Limit', 'Spots Reserved', 'Spots Remaining', 'Overbooked By']
    .forEach(function (header) {
      if (headers[header] && sheet.getLastRow() >= 2) {
        sheet
          .getRange(2, headers[header], sheet.getLastRow() - 1, 1)
          .setNumberFormat('0');
      }
    });

  if (headers['Last Synced'] && sheet.getLastRow() >= 2) {
    sheet
      .getRange(2, headers['Last Synced'], sheet.getLastRow() - 1, 1)
      .setNumberFormat('m/d/yyyy h:mm am/pm');
  }

  sheet.setFrozenRows(1);
}

function styleHeaderRange_(range) {
  range
    .setFontWeight('bold')
    .setBackground('#4E2A84')
    .setFontColor('#FFFFFF')
    .setWrap(true);
}


/* ========================================================================== */
/* GENERAL HELPERS                                                            */
/* ========================================================================== */

function getDb_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(CONFIG.PROP_SPREADSHEET_ID);

  if (!id) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (!active) {
      throw new Error(
        'Spreadsheet ID is not configured. Run setupSheets() manually once from the bound spreadsheet.'
      );
    }
    id = active.getId();
    props.setProperty(CONFIG.PROP_SPREADSHEET_ID, id);
  }

  return SpreadsheetApp.openById(id);
}

function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function parseRequestBody_(e) {
  if (!e) return {};

  if (e.postData && e.postData.contents) {
    const raw = String(e.postData.contents || '').trim();

    if (raw) {
      try {
        return JSON.parse(raw);
      } catch (jsonErr) {
        // Fall through to form parameters below.
      }
    }
  }

  const p = (e && e.parameter) || {};
  if (p.payload) {
    try {
      return JSON.parse(p.payload);
    } catch (err) {
      throw new Error('Invalid JSON payload.');
    }
  }

  return p;
}

function requestResultCacheKey_(token) {
  const clean = cleanSingleLine_(token);
  return clean ? 'imc-request-result:' + clean : '';
}

function rememberRequestResult_(token, result) {
  const key = requestResultCacheKey_(token);
  if (!key) return;

  try {
    CacheService.getScriptCache().put(
      key,
      JSON.stringify(result),
      600
    );
  } catch (err) {
    console.error('Could not cache request result: ' + err);
  }
}

function readRequestResult_(token) {
  const key = requestResultCacheKey_(token);
  if (!key) return null;

  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    console.error('Could not read request result: ' + err);
    return null;
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function respondGet_(e, obj) {
  const callback = cleanSingleLine_(
    e && e.parameter && (e.parameter.callback || e.parameter.prefix)
  );

  if (!callback) return json_(obj);

  // Only allow a simple JS function identifier. Capacity/status JSONP is
  // read-only; this prevents arbitrary script injection through the callback.
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput('throw new Error("Invalid JSONP callback");')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function respondPost_(e, payload, obj) {
  const params = (e && e.parameter) || {};
  const iframeMode = String(params.transport || '').toLowerCase() === 'iframe';
  const responseToken = cleanSingleLine_(
    params.responseToken || payload.clientToken || ''
  );

  if (!iframeMode) return json_(obj);

  // Cross-origin HTML form POSTs are allowed by browsers even when fetch()
  // would be blocked by CORS. Return a tiny page that reports the result
  // back to the IMC Connects parent window.
  const message = {
    source: 'imc-connects-backend',
    token: responseToken,
    data: obj
  };

  const safeJson = JSON.stringify(message).replace(/</g, '\\u003c');
  const html =
    '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>(function(){var m=' + safeJson + ';' +
    'try{window.parent.postMessage(m,"*");}catch(e){}' +
    'try{if(window.top!==window.parent)window.top.postMessage(m,"*");}catch(e){}' +
    '})();<\\/script></body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonError_(err) {
  console.error(err);
  return json_({
    ok: false,
    error: 'SERVER_ERROR',
    message: String(err && err.message || err)
  });
}

function fail_(code, message, extra) {
  return Object.assign(
    {
      ok: false,
      error: code,
      message: message
    },
    extra || {}
  );
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ''));
}

function cleanSingleLine_(value) {
  return String(value == null ? '' : value)
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeText_(value) {
  return cleanSingleLine_(value).toLowerCase();
}

function normalizeIdentityText_(value) {
  return cleanSingleLine_(value)
    .toLowerCase()
    .replace(/[’']/g, "'")
    .replace(/\s+/g, ' ');
}

function arrayOrEmpty_(value) {
  return Array.isArray(value) ? value : [];
}

function toBoolean_(value) {
  if (value === true) return true;
  const s = String(value == null ? '' : value).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === '1';
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16).toUpperCase();
}

function makeAlumniKey_(timestamp, name, company) {
  const base = cleanSingleLine_(timestamp)
    ? 'timestamp|' + cleanSingleLine_(timestamp)
    : 'identity|' + normalizeIdentityText_(name) + '|' + normalizeIdentityText_(company);

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    base,
    Utilities.Charset.UTF_8
  );

  const hex = digest
    .slice(0, 10)
    .map(function (b) {
      const n = b < 0 ? b + 256 : b;
      return ('0' + n.toString(16)).slice(-2);
    })
    .join('')
    .toUpperCase();

  return 'ALUM-' + hex;
}