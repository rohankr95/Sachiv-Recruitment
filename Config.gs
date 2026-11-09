/**
 * Config.gs
 * ---------------------------------------------------------------------------
 * Central configuration for the Application Correction Portal.
 *
 * This file runs ONLY on the server. Nothing in here is ever sent to the
 * browser except the handful of values explicitly listed in
 * getPublicConfig_() (Code.gs).
 *
 * Minimum you must set before running setupPortal():
 *   1. SPREADSHEET_ID            (or bind the script to the spreadsheet)
 *   2. APPLICATION_SHEET_NAME    (the tab that holds the ~2,500 rows)
 *   3. KEY_COLUMNS               (only if your header names differ from the
 *                                 aliases listed below)
 * Drive folder IDs may be left blank: setupPortal() creates the folders and
 * stores their IDs in Script Properties automatically.
 * ---------------------------------------------------------------------------
 */

var CONFIG = {
  /* ------------------------------------------------------------------ */
  /* Spreadsheet                                                         */
  /* ------------------------------------------------------------------ */

  // ID from the spreadsheet URL: https://docs.google.com/spreadsheets/d/<ID>/edit
  // (pasting the whole URL also works – the ID is extracted automatically).
  // Leave as-is if the script is bound to the spreadsheet (Extensions > Apps Script),
  // or if you ran createSampleData() (it stores the sample sheet ID in Script Properties).
  SPREADSHEET_ID: 'PASTE_SPREADSHEET_ID_HERE',

  // Tab that contains the application records (one row per applicant).
  APPLICATION_SHEET_NAME: 'Form Responses 1',

  // Row number that contains the column headers (data starts on the next row).
  HEADER_ROW: 1,

  // Tabs created by setupPortal() inside the same spreadsheet.
  FIELD_CONFIG_SHEET_NAME: 'Field_Config',
  LOG_SHEET_NAME: 'Correction_Log',

  // Optional: keep Correction_Log in a different spreadsheet. Leave blank to
  // use the application spreadsheet.
  LOG_SPREADSHEET_ID: '',

  /* ------------------------------------------------------------------ */
  /* Key columns – matched case-insensitively against the header row.   */
  /* The first alias that exists in the sheet is used.                  */
  /* ------------------------------------------------------------------ */
  KEY_COLUMNS: {
    APPLICATION_NUMBER: ['Application Number', 'Application No', 'आवेदन क्रमांक', 'Application No.', 'Registration Number', 'Registration No', 'Reg No'],
    EMAIL: ['Email Address', 'ईमेल', 'ई-मेल', 'Email', 'Email ID', 'E-mail', 'E-mail Address', 'Registered Email'],
    MOBILE: ['मोबाइल नंबर', 'मोबाइल नम्बर', 'Mobile Number', 'Mobile No', 'Mobile No.', 'Mobile', 'Registered Mobile Number', 'Contact Number', 'Phone Number'],
    DOB: ['जन्म तिथि', 'जन्मतिथि', 'Date of Birth', 'DOB', 'Birth Date', 'Date Of Birth (DD/MM/YYYY)'],
    NAME: ['आवेदक/आवेदिका का नाम (अंग्रेजी में, कैपिटल लेटर में)', 'आवेदक/आवेदिका का नाम (हिन्दी में)', 'Applicant Name', 'Name of Applicant', 'Candidate Name', 'Full Name', 'Name']
  },

  // Columns maintained by the portal. setupPortal() appends them to the end
  // of the header row if they do not already exist. No other structural
  // change is made to the application sheet.
  SYSTEM_COLUMNS: {
    LAST_MODIFIED: 'Last Modified',
    CORRECTION_COUNT: 'Correction Count',
    CORRECTION_STATUS: 'Correction Status',
    LAST_TRANSACTION_ID: 'Last Correction ID'
  },

  /* ------------------------------------------------------------------ */
  /* Login behaviour                                                     */
  /* ------------------------------------------------------------------ */

  // How to treat an email / mobile number that appears on more than one row.
  //  'MATCH_DOB' (recommended): the applicant sees every application that has
  //      this email/mobile AND the entered date of birth. All of them are
  //      flagged "Duplicate". Rows with the same email but a different date of
  //      birth (i.e. a different person) are never shown.
  //  'SHOW_ALL': once the date of birth matches at least one row, every
  //      application registered with that email/mobile is shown and flagged
  //      "Duplicate". Use only if you are sure shared emails belong to the
  //      same person – otherwise one applicant can edit another's record.
  DUPLICATE_IDENTIFIER_POLICY: 'MATCH_DOB',

  // Rows without a valid date of birth are incomplete applications (e.g. the
  // Google Form ended because the applicant answered "No" to an eligibility
  // question). When true they are excluded: they cannot log in, cannot be
  // opened or corrected, and are counted separately on the admin dashboard.
  EXCLUDE_ROWS_WITHOUT_DOB: true,

  // Mobile number rule (after removing spaces, +91 and leading 0).
  MOBILE_REGEX: '^[6-9][0-9]{9}$',

  /* ------------------------------------------------------------------ */
  /* Sessions and rate limiting                                          */
  /* ------------------------------------------------------------------ */
  SESSION: {
    IDLE_TIMEOUT_MINUTES: 15,      // applicant is logged out after this much inactivity
    ABSOLUTE_TIMEOUT_MINUTES: 90,  // hard limit, regardless of activity
    ADMIN_IDLE_TIMEOUT_MINUTES: 30,
    ADMIN_ABSOLUTE_TIMEOUT_MINUTES: 240
  },

  RATE_LIMIT: {
    MAX_FAILED_ATTEMPTS: 5,        // failed applicant logins per email/mobile ...
    LOCKOUT_MINUTES: 15,           // ... before that email/mobile is locked for this long
    ADMIN_MAX_FAILED_ATTEMPTS: 5,
    ADMIN_LOCKOUT_MINUTES: 30,
    FAILURE_DELAY_MS: 1000         // server-side delay after every failed login (slows brute force)
  },

  /* ------------------------------------------------------------------ */
  /* Google Drive                                                        */
  /* ------------------------------------------------------------------ */
  // Leave blank to let setupPortal() create the folders (IDs are then kept in
  // Script Properties as FOLDER_UPLOAD / FOLDER_ARCHIVE / FOLDER_STAGING).
  DRIVE: {
    UPLOAD_FOLDER_ID: '',   // final location of replacement documents
    ARCHIVE_FOLDER_ID: '',  // copies of documents that were replaced
    STAGING_FOLDER_ID: ''   // temporary uploads awaiting final submission
  },

  // 'COPY' (default): a copy of the old document is placed in the archive
  //                   folder; the original file is left untouched.
  // 'MOVE':           the original file itself is moved to the archive folder.
  // Original documents are never deleted in either mode.
  ARCHIVE_MODE: 'COPY',

  UPLOAD: {
    MAX_FILE_SIZE_MB: 2,
    ALLOWED_EXTENSIONS: ['pdf', 'jpg', 'jpeg'],
    ALLOWED_MIME_TYPES: ['application/pdf', 'image/jpeg', 'image/pjpeg'],
    MAX_UPLOADS_PER_SESSION: 30,
    MAX_VIEW_SIZE_MB: 10,        // largest document the in-portal viewer will stream
    STAGING_RETENTION_HOURS: 24  // abandoned staged uploads older than this are trashed
  },

  /* ------------------------------------------------------------------ */
  /* Dates                                                               */
  /* ------------------------------------------------------------------ */
  DATE: {
    // How ambiguous text dates such as 05/08/2000 are stored in your sheet:
    // 'DMY' = 5 Aug 2000, 'MDY' = 8 May 2000.
    TEXT_DATE_ORDER: 'DMY',
    // Format used when a corrected date is written into a cell that held text.
    STORAGE_FORMAT: 'dd/MM/yyyy',
    DISPLAY_FORMAT: 'dd/MM/yyyy'
  },

  /* ------------------------------------------------------------------ */
  /* Workflow                                                            */
  /* ------------------------------------------------------------------ */
  PORTAL_OPEN: true,                     // false = applicants can log in and view but not submit
  // The portal closes automatically after this moment (ISO date-time with +05:30 for IST).
  // Use '' for no automatic closing. CORRECTION_DEADLINE_TEXT is what applicants see.
  CORRECTION_DEADLINE: '2026-11-10T23:59:59+05:30',
  CORRECTION_DEADLINE_TEXT: '10/11/2026, 11:59 PM',
  PORTAL_CLOSED_MESSAGE: 'The correction window is closed. / सुधार की अवधि समाप्त हो चुकी है।',
  REQUIRE_APPROVAL: false,               // true = corrections are marked "Pending Approval" for admin review
  MAX_CORRECTIONS_PER_APPLICATION: 1,    // 1 = one-time correction; 0 = unlimited

  STATUS: {
    MODIFIED: 'Modified',
    PENDING: 'Pending Approval',
    APPROVED: 'Approved',
    REJECTED: 'Rejected'
  },

  /* ------------------------------------------------------------------ */
  /* Branding                                                            */
  /* ------------------------------------------------------------------ */
  PORTAL_TITLE: 'Gram Panchayat Sachiv Recruitment – Application Correction Portal',
  ORGANISATION_NAME: 'Zila Panchayat, Surajpur (C.G.)',
  SUPPORT_TEXT: 'For assistance, contact the helpdesk during office hours.',

  /* ------------------------------------------------------------------ */
  /* Form sections, in display order. Section names used in the         */
  /* Field_Config sheet must match these (unknown names are appended).  */
  /* ------------------------------------------------------------------ */
  SECTIONS: [
    { name: 'Personal Information', icon: 'fa-user' },
    { name: 'Address Details', icon: 'fa-location-dot' },
    { name: 'Educational Qualifications', icon: 'fa-graduation-cap' },
    { name: 'Category and Reservation Details', icon: 'fa-scale-balanced' },
    { name: 'Contact Information', icon: 'fa-address-book' },
    { name: 'Other Application Details', icon: 'fa-file-lines' },
    { name: 'Uploaded Documents', icon: 'fa-folder-open' }
  ],
  DOCUMENT_SECTION: 'Uploaded Documents',
  DEFAULT_SECTION: 'Other Application Details',

  /* ------------------------------------------------------------------ */
  /* Technical                                                           */
  /* ------------------------------------------------------------------ */
  FIELD_CONFIG_CACHE_SECONDS: 600, // Field_Config is non-sensitive and cached
  LOCK_TIMEOUT_MS: 30000,
  MAX_TEXT_LENGTH: 500,
  MAX_TEXTAREA_LENGTH: 2000
};

/** Field types understood by the portal (used for Field_Config validation). */
var FIELD_TYPES = ['text', 'textarea', 'email', 'mobile', 'date', 'number', 'select', 'pincode', 'document', 'readonly', 'hidden'];

/** Columns of the Field_Config sheet. */
var FIELD_CONFIG_HEADERS = [
  'Column Header', 'Display Label', 'Section', 'Field Type', 'Required (Yes/No)',
  'Editable (Yes/No)', 'Options (comma separated)', 'Validation Regex', 'Max Length',
  'Help Text', 'Display Order'
];

/** Columns of the Correction_Log sheet. */
var LOG_HEADERS = [
  'Correction Transaction ID', 'Application Number', 'Field Name', 'Previous Value',
  'Updated Value', 'Modification Date and Time', 'Document Replacement Details',
  'Correction Status', 'Changed By'
];
