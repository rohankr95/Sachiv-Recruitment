/**
 * SampleData.gs
 * ---------------------------------------------------------------------------
 * Generates a TEST spreadsheet with realistic synthetic records (50 columns,
 * Drive document links, duplicate emails) so the portal can be tested
 * without touching real applicant data.
 *
 * Usage (Apps Script editor):
 *   1. Run createSampleData()  – creates "Correction Portal – Sample Data"
 *      and stores its ID in Script Properties (SPREADSHEET_ID) when
 *      CONFIG.SPREADSHEET_ID has not been set.
 *   2. Run setupPortal().
 *   3. Open View > Logs: sample login credentials are printed.
 * Delete the Script Property SPREADSHEET_ID afterwards to switch back.
 * ---------------------------------------------------------------------------
 */

var SAMPLE_RECORD_COUNT_ = 2500;

var SAMPLE_JPEG_BASE64_ = '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';

function createSampleData() {
  requireEditorContext_();
  var started = Date.now();
  var headers = [
    'Timestamp', 'Application Number', 'Applicant Name', "Father's Name", "Mother's Name", 'Date of Birth', 'Gender',
    'Marital Status', 'Nationality', 'Religion', 'Aadhaar Number', 'Category', 'Sub Caste', 'Are you EWS', 'Are you PwD',
    'Domicile of Chhattisgarh', 'Email Address', 'Mobile Number', 'Alternate Mobile Number', 'Permanent Address',
    'Village/Town', 'Block', 'District', 'State', 'Pin Code', 'Correspondence Address', 'Correspondence District',
    'Correspondence Pin Code', '10th Board', '10th Passing Year', '10th Percentage', '12th Board', '12th Passing Year',
    '12th Stream', '12th Percentage', 'Graduation Degree', 'Graduation University', 'Graduation Passing Year',
    'Graduation Percentage', 'Post Applied For', 'Preferred Exam Centre', 'Work Experience (Years)', 'Identification Mark',
    'Photograph', 'Signature', '10th Marksheet', '12th Marksheet', 'Caste Certificate', 'Domicile Certificate',
    'Aadhaar Card Copy'
  ];

  // Sample documents (private, in the owner's Drive).
  var docFolder = DriveApp.createFolder('Correction Portal – Sample Documents');
  var pdf = docFolder.createFile(Utilities.newBlob(samplePdf_('Sample certificate'), 'application/pdf', 'sample_certificate.pdf'));
  var jpg = docFolder.createFile(Utilities.newBlob(Utilities.base64Decode(SAMPLE_JPEG_BASE64_), 'image/jpeg', 'sample_photo.jpg'));
  var pdfUrl = 'https://drive.google.com/open?id=' + pdf.getId();
  var jpgUrl = 'https://drive.google.com/open?id=' + jpg.getId();

  var first = ['Aarav', 'Vivaan', 'Aditya', 'Priya', 'Ananya', 'Rohit', 'Sunita', 'Kavita', 'Rahul', 'Pooja', 'Deepak', 'Neha', 'Sanjay', 'Meena', 'Ajay', 'Ritu'];
  var last = ['Sharma', 'Verma', 'Sahu', 'Patel', 'Yadav', 'Singh', 'Kurre', 'Markam', 'Netam', 'Dewangan', 'Gupta', 'Tiwari'];
  var blocks = ['Surajpur', 'Bhaiyathan', 'Odgi', 'Pratappur', 'Ramanujnagar', 'Premnagar'];
  var categories = ['General', 'OBC', 'SC', 'ST'];
  var boards = ['CGBSE', 'CBSE', 'ICSE'];
  var streams = ['Science', 'Commerce', 'Arts', 'Agriculture'];
  var posts = ['Assistant Teacher', 'Clerk', 'Data Entry Operator', 'Peon'];
  var seed = 7;
  var rnd = function (n) { seed = (seed * 9301 + 49297) % 233280; return Math.floor(seed / 233280 * n); };
  var pick = function (a) { return a[rnd(a.length)]; };
  var pad = function (n, w) { var s = String(n); while (s.length < w) s = '0' + s; return s; };

  var rows = [];
  var tz = Session.getScriptTimeZone();
  for (var i = 0; i < SAMPLE_RECORD_COUNT_; i++) {
    var fn = pick(first), ln = pick(last);
    var dob = new Date(1985 + rnd(18), rnd(12), 1 + rnd(28));
    var email = (fn + '.' + ln + (i + 11)).toLowerCase() + '@example.com';
    var mobile = String(6000000000 + i * 1373 + 7000000);
    var prev = rows[rows.length - 1];
    if (i > 0 && i % 25 === 0) { email = prev[16]; mobile = prev[17]; dob = prev[5]; }       // same person, 2nd application
    else if (i > 0 && i % 40 === 0) { email = prev[16]; }                                  // shared (cyber-cafe) email, different person
    var block = pick(blocks);
    var y10 = dob.getFullYear() + 16;
    rows.push([
      new Date(2026, 8, 1 + rnd(28), 10, rnd(60)), 'APP2026' + pad(i + 1, 5), fn + ' ' + ln, pick(first) + ' ' + ln, pick(first) + ' ' + ln,
      dob, (['Priya', 'Ananya', 'Sunita', 'Kavita', 'Pooja', 'Neha', 'Meena', 'Ritu'].indexOf(fn) !== -1 ? 'Female' : 'Male'),
      pick(['Unmarried', 'Married']), 'Indian', pick(['Hindu', 'Christian', 'Muslim', 'Sikh']), pad(rnd(999999), 6) + pad(rnd(999999), 6),
      pick(categories), '', pick(['Yes', 'No']), pick(['No', 'No', 'No', 'Yes']), 'Yes', email, mobile, '',
      'House No. ' + (1 + rnd(300)) + ', Ward ' + (1 + rnd(20)) + ', ' + block, block, block, 'Surajpur', 'Chhattisgarh', '497229',
      '', '', '', pick(boards), y10, 45 + rnd(50), pick(boards), y10 + 2, pick(streams), 45 + rnd(50),
      pick(['B.A.', 'B.Sc.', 'B.Com.', '']), pick(['Sant Gahira Guru University', 'Pt. Ravishankar Shukla University', '']),
      y10 + 5, 40 + rnd(50), pick(posts), pick(['Surajpur', 'Ambikapur', 'Bilaspur']), rnd(6), 'Mole on left cheek',
      jpgUrl, jpgUrl, pdfUrl, pdfUrl, pdfUrl, pdfUrl, pdfUrl
    ]);
  }
  // A few realistic imperfections.
  rows[3][20] = ''; rows[3][24] = '';            // missing village and pin code
  rows[7][17] = '+91 ' + rows[7][17];            // mobile with country code
  rows[9][46] = '';                              // missing 12th marksheet

  var ss = SpreadsheetApp.create('Correction Portal – Sample Data');
  var sheet = ss.getSheets()[0].setName(CONFIG.APPLICATION_SHEET_NAME);
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  if (sheet.getMaxRows() < rows.length + 1) sheet.insertRowsAfter(sheet.getMaxRows(), rows.length + 1 - sheet.getMaxRows());
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(2, 6, rows.length, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2, 11, rows.length, 1).setNumberFormat('@');
  sheet.setFrozenRows(1);
  DriveApp.getFileById(ss.getId()).moveTo(docFolder);

  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  }
  RUNTIME_.spreadsheet = null;

  var out = ['Sample spreadsheet created: ' + ss.getUrl(), 'Rows: ' + rows.length + ' in ' + Math.round((Date.now() - started) / 1000) + 's', 'Sample logins (email or mobile + DOB):'];
  [0, 1, 2, 24, 25, 39, 40].forEach(function (i) {
    out.push('  ' + rows[i][1] + ' | ' + rows[i][16] + ' | ' + rows[i][17] + ' | DOB ' + Utilities.formatDate(rows[i][5], tz, 'yyyy-MM-dd'));
  });
  out.push('Rows APP202600025/26 share email + DOB (same person: both applications are listed as duplicates).');
  out.push('Rows APP202600040/41 share email but different DOB (only the matching one is shown under MATCH_DOB).');
  out.push('Next: run setupPortal().');
  console.log(out.join('\n'));
  return out.join('\n');
}

/** Minimal single-page PDF containing one line of text. */
function samplePdf_(text) {
  var content = 'BT /F1 18 Tf 40 80 Td (' + String(text).replace(/[()\\]/g, '') + ') Tj ET';
  var objs = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 200] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    '<< /Length ' + content.length + ' >>\nstream\n' + content + '\nendstream',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
  ];
  var pdf = '%PDF-1.4\n';
  var offsets = [];
  objs.forEach(function (o, i) {
    offsets.push(pdf.length);
    pdf += (i + 1) + ' 0 obj\n' + o + '\nendobj\n';
  });
  var xref = pdf.length;
  pdf += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n';
  offsets.forEach(function (off) { pdf += ('0000000000' + off).slice(-10) + ' 00000 n \n'; });
  pdf += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF';
  return Utilities.newBlob(pdf).getBytes();
}
