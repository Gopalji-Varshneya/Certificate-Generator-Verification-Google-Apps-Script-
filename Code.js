/****************************************************
 * CERTIFICATE GENERATOR  (OPTIMIZED)
 * Google Sheets + Google Slides + Drive + Gmail
 *
 * What changed vs. the previous version (no config,
 * column names, placeholders or function names changed):
 *  - QR codes are fetched in PARALLEL (fetchAll)
 *  - PDF exports are fetched in PARALLEL (fetchAll)
 *  - Sheet writes are BATCHED (one setValues per block
 *    of rows instead of ~5 setValue calls per student)
 *  - Certificate IDs are written in one batch
 *  - Work is done in small chunks with a time guard, so a
 *    big BATCH_SIZE stops cleanly before Apps Script's
 *    execution-time limit instead of dying half-way
 *  - Temp Slides copies are trashed even when a row fails
 *  - Email status is written in batches, and sending
 *    respects the remaining daily mail quota
 ****************************************************/

const CONFIG = {

  // Google Spreadsheet ID
  SPREADSHEET_ID:
    'xxx',

  // Name of the tab inside the spreadsheet
  STUDENT_SHEET:
    'Students',

  // Google Slides certificate template
  SLIDES_TEMPLATE_ID:
    'xxx',

  // Google Drive folder where certificates will be saved
  OUTPUT_FOLDER_ID:
    'xxx',

  // Deployed Web App /exec URL
  VERIFY_URL:
    'https://script.google.com/macros/s/xxx/exec',

  // IMPORTANT:
  // Keep this at 1 while testing.
  BATCH_SIZE: 1,

  // Email is sent separately using sendEmailBatch().
  SEND_EMAIL: true
};


/****************************************************
 * TEST ONE STUDENT
 *
 * This processes only the first pending student.
 ****************************************************/

function testOneStudent() {
  generateBatch();
}


/****************************************************
 * CREATE CUSTOM MENU
 ****************************************************/

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu('Certificates')

    .addItem(
      'Test / Generate next certificate',
      'generateBatch'
    )

    .addItem(
      'Send next email batch',
      'sendEmailBatch'
    )

    .addItem(
      'Generate missing Certificate IDs',
      'generateMissingIDs'
    )

    .addToUi();
}


/****************************************************
 * HELPER: WRITE ONE COLUMN FOR MANY ROWS
 *
 * rowIndexes = ascending 0-based indexes into the
 * `data` array (row 1 of the sheet = index 0).
 * Uses ONE setValues() call per block of consecutive
 * rows instead of one setValue() call per cell, and
 * never touches rows that are not in rowIndexes.
 ****************************************************/

function writeColumnRows_(sheet, col, rowIndexes, getValue) {

  if (!rowIndexes.length) {
    return;
  }

  let start = 0;

  for (let k = 1; k <= rowIndexes.length; k++) {

    if (
      k === rowIndexes.length ||
      rowIndexes[k] !== rowIndexes[k - 1] + 1
    ) {

      const block =
        rowIndexes.slice(start, k);

      sheet
        .getRange(
          block[0] + 1,
          col + 1,
          block.length,
          1
        )
        .setValues(
          block.map(function (r) {
            return [getValue(r)];
          })
        );

      start = k;
    }
  }
}


/****************************************************
 * GENERATE MISSING CERTIFICATE IDs
 ****************************************************/

function generateMissingIDs() {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.STUDENT_SHEET
    );

  if (!sheet) {
    throw new Error(
      'Sheet "' +
      CONFIG.STUDENT_SHEET +
      '" was not found.'
    );
  }

  const data =
    sheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error(
      'No students found.'
    );
  }

  const headers = data[0];

  const nameCol =
    headers.indexOf('Name');

  const idCol =
    headers.indexOf('Certificate_ID');

  if (nameCol === -1 || idCol === -1) {

    throw new Error(
      'Required columns missing. ' +
      'Make sure the sheet contains Name and Certificate_ID.'
    );
  }

  fillMissingIDs_(
    sheet,
    data,
    nameCol,
    idCol
  );

  SpreadsheetApp.flush();

  Logger.log(
    'Certificate IDs generated.'
  );
}


/****************************************************
 * FILL MISSING IDs (shared by generateMissingIDs and
 * generateBatch). Updates `data` in memory AND writes
 * all new IDs to the sheet in batched calls.
 ****************************************************/

function fillMissingIDs_(sheet, data, nameCol, idCol) {

  const existingIDs =
    new Set();

  for (let i = 1; i < data.length; i++) {

    const id =
      String(
        data[i][idCol] || ''
      ).trim();

    if (id) {
      existingIDs.add(id);
    }
  }

  const rows = [];

  for (let i = 1; i < data.length; i++) {

    const name =
      String(
        data[i][nameCol] || ''
      ).trim();

    const currentID =
      String(
        data[i][idCol] || ''
      ).trim();

    if (!name || currentID) {
      continue;
    }

    let newID;

    do {
      newID =
        createCertificateID();
    }
    while (
      existingIDs.has(newID)
    );

    existingIDs.add(newID);

    data[i][idCol] = newID;

    rows.push(i);
  }

  writeColumnRows_(
    sheet,
    idCol,
    rows,
    function (r) {
      return data[r][idCol];
    }
  );

  return rows.length;
}


/****************************************************
 * CREATE CERTIFICATE ID
 ****************************************************/

function createCertificateID() {

  const chars =
    'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

  let result =
    'CERT-';

  for (let i = 0; i < 8; i++) {

    result +=
      chars.charAt(
        Math.floor(
          Math.random() *
          chars.length
        )
      );
  }

  return result;
}


/****************************************************
 * GENERATE CERTIFICATE BATCH
 *
 * BATCH_SIZE = 1 during testing.
 ****************************************************/

function generateBatch() {

  const startTime =
    Date.now();

  // Stop starting new work after this long, so the run
  // ends cleanly before Apps Script's execution limit.
  const MAX_RUNTIME_MS =
    5 * 60 * 1000;

  // Certificates handled together (QR + PDF fetched in parallel).
  const CHUNK_SIZE = 10;

  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.STUDENT_SHEET
    );

  if (!sheet) {
    throw new Error(
      'Sheet "' +
      CONFIG.STUDENT_SHEET +
      '" was not found.'
    );
  }

  const data =
    sheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error(
      'No students found.'
    );
  }

  const headers =
    data[0];

  const nameCol =
    headers.indexOf('Name');

  const emailCol =
    headers.indexOf('Email');

  const streamCol =
    headers.indexOf('Stream');

  const acmCol =
    headers.indexOf('Academic_Year');

  const dateCol =
    headers.indexOf('IssueDate');

  const idCol =
    headers.indexOf('Certificate_ID');

  const pdfCol =
    headers.indexOf('PDF_URL');

  const verifyCol =
    headers.indexOf('Verify_URL');

  const statusCol =
    headers.indexOf('Status');

  if (
    nameCol === -1 ||
    emailCol === -1 ||
    streamCol === -1 ||
    acmCol === -1 ||
    dateCol === -1 ||
    idCol === -1 ||
    pdfCol === -1 ||
    verifyCol === -1 ||
    statusCol === -1
  ) {

    throw new Error(
      'Check your column headings.\n\n' +
      'Required:\n' +
      'Certificate_ID, Team, Role, Name, Email, ' +
      'Stream, Academic_Year, IssueDate, PDF_URL, ' +
      'Verify_URL, Status, Email_Status'
    );
  }

  // Make sure IDs exist first (in memory + one batched write).
  fillMissingIDs_(
    sheet,
    data,
    nameCol,
    idCol
  );

  /*
   * Collect the pending rows (up to BATCH_SIZE).
   */

  const jobs = [];

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      jobs.length >=
      CONFIG.BATCH_SIZE
    ) {
      break;
    }

    const row =
      data[i];

    const name =
      String(
        row[nameCol] || ''
      ).trim();

    const email =
      String(
        row[emailCol] || ''
      ).trim();

    const stream =
      String(
        row[streamCol] || ''
      ).trim();

    const acm =
      String(
        row[acmCol] || ''
      ).trim();

    let issueDate =
      row[dateCol];

    const certificateID =
      String(
        row[idCol] || ''
      ).trim();

    const existingPDF =
      String(
        row[pdfCol] || ''
      ).trim();

    /*
     * Skip incomplete rows.
     */

    if (
      !name ||
      !email ||
      !certificateID
    ) {
      continue;
    }

    /*
     * Skip already generated certificates.
     */

    if (existingPDF) {
      continue;
    }

    /*
     * If IssueDate is blank,
     * automatically use today's date.
     */

    let dateWasBlank =
      false;

    if (!issueDate) {

      issueDate =
        new Date();

      row[dateCol] =
        issueDate;

      dateWasBlank =
        true;
    }

    const formattedDate =
      formatDate(issueDate);

    const verifyURL =
      CONFIG.VERIFY_URL +
      '?id=' +
      encodeURIComponent(
        certificateID
      );

    jobs.push({
      index: i,
      name: name,
      stream: stream,
      acm: acm,
      certificateID: certificateID,
      formattedDate: formattedDate,
      verifyURL: verifyURL,
      dateWasBlank: dateWasBlank
    });
  }

  if (jobs.length === 0) {

    Logger.log(
      'Generated certificates: 0'
    );

    return;
  }

  const templateFile =
    DriveApp.getFileById(
      CONFIG.SLIDES_TEMPLATE_ID
    );

  const outputFolder =
    DriveApp.getFolderById(
      CONFIG.OUTPUT_FOLDER_ID
    );

  const ctx = {
    sheet: sheet,
    data: data,
    dateCol: dateCol,
    pdfCol: pdfCol,
    verifyCol: verifyCol,
    statusCol: statusCol,
    templateFile: templateFile,
    outputFolder: outputFolder
  };

  let processed = 0;

  let next = 0;

  // First guess per certificate; refined after every chunk.
  let perCertMs = 12000;

  while (next < jobs.length) {

    const remainingMs =
      MAX_RUNTIME_MS -
      (Date.now() - startTime);

    const allowed =
      Math.floor(
        remainingMs / perCertMs
      );

    if (allowed < 1) {

      Logger.log(
        'Stopped early to stay inside the Apps Script ' +
        'time limit. ' +
        (jobs.length - next) +
        ' certificate(s) left - run again to continue.'
      );

      break;
    }

    const chunk =
      jobs.slice(
        next,
        next + Math.min(
          CHUNK_SIZE,
          allowed
        )
      );

    next += chunk.length;

    const chunkStart =
      Date.now();

    processed +=
      generateChunk_(
        ctx,
        chunk
      );

    perCertMs =
      Math.max(
        2000,
        (Date.now() - chunkStart) /
        chunk.length
      );
  }

  SpreadsheetApp.flush();

  Logger.log(
    'Generated certificates: ' +
    processed
  );
}


/****************************************************
 * GENERATE ONE CHUNK OF CERTIFICATES
 *
 * 1. mark rows PROCESSING (batched write)
 * 2. fetch all QR codes in parallel
 * 3. copy + fill each Slides template
 * 4. export all PDFs in parallel
 * 5. save PDFs to Drive
 * 6. write all results back (batched)
 *
 * Returns the number of certificates generated.
 ****************************************************/

function generateChunk_(ctx, chunk) {

  const sheet = ctx.sheet;
  const data = ctx.data;
  const dateCol = ctx.dateCol;
  const pdfCol = ctx.pdfCol;
  const verifyCol = ctx.verifyCol;
  const statusCol = ctx.statusCol;
  const templateFile = ctx.templateFile;
  const outputFolder = ctx.outputFolder;

  /*
   * 1. Mark as processing (one write for the whole chunk).
   */

  chunk.forEach(function (job) {
    data[job.index][statusCol] =
      'PROCESSING';
  });

  writeColumnRows_(
    sheet,
    statusCol,
    chunk.map(function (j) {
      return j.index;
    }),
    function (r) {
      return data[r][statusCol];
    }
  );

  writeColumnRows_(
    sheet,
    dateCol,
    chunk
      .filter(function (j) {
        return j.dateWasBlank;
      })
      .map(function (j) {
        return j.index;
      }),
    function (r) {
      return data[r][dateCol];
    }
  );

  /*
   * 2. Fetch all QR codes in parallel.
   *    (Any failure falls back to createQRCode() later.)
   */

  let qrResponses = [];

  try {

    qrResponses =
      UrlFetchApp.fetchAll(
        chunk.map(function (job) {
          return {
            url:
              'https://quickchart.io/qr' +
              '?text=' +
              encodeURIComponent(job.verifyURL) +
              '&size=500' +
              '&margin=2',
            muteHttpExceptions: true
          };
        })
      );

  } catch (e) {

    qrResponses = [];
  }

  chunk.forEach(function (job, k) {

    const resp =
      qrResponses[k];

    job.qrBlob =
      resp &&
      resp.getResponseCode() === 200
        ? resp.getBlob().setName('qr.png')
        : null;
  });

  /*
   * 3. Copy template + fill it in (one row at a time).
   */

  chunk.forEach(function (job) {

    try {

      /*
       * Make temporary copy of template.
       */

      const tempFile =
        templateFile.makeCopy(
          job.certificateID,
          outputFolder
        );

      job.tempFile =
        tempFile;

      job.tempId =
        tempFile.getId();

      const presentation =
        SlidesApp.openById(
          job.tempId
        );

      const slides =
        presentation.getSlides();

      if (slides.length === 0) {

        throw new Error(
          'Template contains no slides.'
        );
      }

      const slide =
        slides[0];

      /*
       * Replace text placeholders.
       *
       * Your Google Slides template should contain:
       *
       * {{NAME}}
       * {{COURSE}}
       * {{Academic_Year}}
       * {{DATE}}
       * {{CERTIFICATE_ID}}
       * {{QR}}
       */

      presentation.replaceAllText(
        '{{NAME}}',
        job.name
      );

      presentation.replaceAllText(
        '{{COURSE}}',
        job.stream
      );

      presentation.replaceAllText(
        '{{Academic_Year}}',
        job.acm
      );

      presentation.replaceAllText(
        '{{DATE}}',
        job.formattedDate
      );

      presentation.replaceAllText(
        '{{CERTIFICATE_ID}}',
        job.certificateID
      );

      /*
       * QR code (already fetched in parallel above;
       * fall back to a single fetch if that one failed).
       */

      const qrBlob =
        job.qrBlob ||
        createQRCode(
          job.verifyURL
        );

      /*
       * Replace {{QR}} with QR image.
       */

      replaceQRPlaceholder(
        slide,
        qrBlob
      );

      presentation.saveAndClose();

    } catch (error) {

      job.error =
        error.message;
    }
  });

  /*
   * 4. Export all finished slides to PDF in parallel.
   */

  const okJobs =
    chunk.filter(function (j) {
      return !j.error;
    });

  const token =
    okJobs.length
      ? ScriptApp.getOAuthToken()
      : '';

  let exportResponses = [];

  if (okJobs.length) {

    try {

      exportResponses =
        UrlFetchApp.fetchAll(
          okJobs.map(function (job) {
            return {
              url:
                'https://docs.google.com/presentation/d/' +
                job.tempId +
                '/export/pdf',
              headers: {
                Authorization:
                  'Bearer ' +
                  token
              },
              muteHttpExceptions: true
            };
          })
        );

    } catch (e) {

      exportResponses = [];
    }
  }

  /*
   * 5. Save PDFs to the output folder.
   */

  okJobs.forEach(function (job, k) {

    try {

      const resp =
        exportResponses[k];

      // If a parallel export failed, retry that one on its own.
      const pdfBlob =
        resp &&
        resp.getResponseCode() === 200
          ? resp
              .getBlob()
              .setContentType('application/pdf')
          : exportPresentationAsPDF(
              job.tempId
            );

      pdfBlob.setName(
        job.certificateID +
        '.pdf'
      );

      const pdfFile =
        outputFolder.createFile(
          pdfBlob
        );

      /*
       * Allow anyone with link to view.
       */

      pdfFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );

      job.pdfURL =
        pdfFile.getUrl();

    } catch (error) {

      job.error =
        error.message;
    }
  });

  /*
   * Remove temporary Slides copies (also for failed rows).
   */

  chunk.forEach(function (job) {

    if (job.tempFile) {

      try {

        job.tempFile.setTrashed(
          true
        );

      } catch (e) {

        // ignore - leftover temp copy is harmless
      }
    }
  });

  /*
   * 6. Save URLs and status (batched).
   *    Email remains pending.
   *    sendEmailBatch() handles email separately.
   */

  let generated = 0;

  chunk.forEach(function (job) {

    if (job.error) {

      data[job.index][statusCol] =
        'ERROR: ' +
        job.error;

      Logger.log(
        'ERROR for row ' +
        (job.index + 1) +
        ': ' +
        job.error
      );

    } else {

      data[job.index][pdfCol] =
        job.pdfURL;

      data[job.index][verifyCol] =
        job.verifyURL;

      data[job.index][statusCol] =
        'GENERATED';

      generated++;

      Logger.log(
        'Generated: ' +
        job.certificateID +
        ' - ' +
        job.name
      );
    }
  });

  const doneRows =
    chunk
      .filter(function (j) {
        return !j.error;
      })
      .map(function (j) {
        return j.index;
      });

  writeColumnRows_(
    sheet,
    pdfCol,
    doneRows,
    function (r) {
      return data[r][pdfCol];
    }
  );

  writeColumnRows_(
    sheet,
    verifyCol,
    doneRows,
    function (r) {
      return data[r][verifyCol];
    }
  );

  writeColumnRows_(
    sheet,
    statusCol,
    chunk.map(function (j) {
      return j.index;
    }),
    function (r) {
      return data[r][statusCol];
    }
  );

  return generated;
}


/****************************************************
 * CREATE QR CODE
 ****************************************************/

function createQRCode(url) {

  const qrURL =
    'https://quickchart.io/qr' +
    '?text=' +
    encodeURIComponent(url) +
    '&size=500' +
    '&margin=2';

  const response =
    UrlFetchApp.fetch(
      qrURL
    );

  if (
    response.getResponseCode() !== 200
  ) {

    throw new Error(
      'QR code generation failed.'
    );
  }

  return response
    .getBlob()
    .setName(
      'qr.png'
    );
}


/****************************************************
 * REPLACE {{QR}} PLACEHOLDER
 ****************************************************/

function replaceQRPlaceholder(
  slide,
  qrBlob
) {

  const shapes =
    slide.getShapes();

  for (
    let i = 0;
    i < shapes.length;
    i++
  ) {

    const shape =
      shapes[i];

    let text = '';

    try {

      text =
        shape
          .getText()
          .asString()
          .trim();

    } catch (e) {

      continue;
    }

    if (
      text === '{{QR}}'
    ) {

      const left =
        shape.getLeft();

      const top =
        shape.getTop();

      const width =
        shape.getWidth();

      const height =
        shape.getHeight();

      shape.remove();

      slide
        .insertImage(qrBlob)
        .setLeft(left)
        .setTop(top)
        .setWidth(width)
        .setHeight(height);

      return;
    }
  }

  throw new Error(
    'QR placeholder {{QR}} was not found.'
  );
}


/****************************************************
 * EXPORT SLIDES TO PDF
 ****************************************************/

function exportPresentationAsPDF(
  presentationId
) {

  const url =
    'https://docs.google.com/presentation/d/' +
    presentationId +
    '/export/pdf';

  const token =
    ScriptApp.getOAuthToken();

  const response =
    UrlFetchApp.fetch(
      url,
      {
        headers: {
          Authorization:
            'Bearer ' +
            token
        },
        muteHttpExceptions:
          true
      }
    );

  if (
    response.getResponseCode() !== 200
  ) {

    throw new Error(
      'PDF export failed: ' +
      response.getContentText()
    );
  }

  return response
    .getBlob()
    .setContentType(
      'application/pdf'
    );
}


/****************************************************
 * FORMAT DATE
 ****************************************************/

function formatDate(value) {

  if (!value) {
    return '';
  }

  if (
    Object.prototype.toString.call(
      value
    ) === '[object Date]'
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'dd MMMM yyyy'
    );
  }

  return String(value);
}


/****************************************************
 * SEND EMAIL BATCH
 ****************************************************/

function sendEmailBatch() {

  const startTime =
    Date.now();

  const MAX_RUNTIME_MS =
    5 * 60 * 1000;

  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.STUDENT_SHEET
    );

  if (!sheet) {
    throw new Error(
      'Students sheet not found.'
    );
  }

  const data =
    sheet
      .getDataRange()
      .getValues();

  if (data.length < 2) {
    throw new Error(
      'No students found.'
    );
  }

  const headers =
    data[0];

  const nameCol =
    headers.indexOf('Name');

  const emailCol =
    headers.indexOf('Email');

  const idCol =
    headers.indexOf(
      'Certificate_ID'
    );

  const pdfCol =
    headers.indexOf('PDF_URL');

  const emailStatusCol =
    headers.indexOf(
      'Email_Status'
    );

  if (
    nameCol === -1 ||
    emailCol === -1 ||
    idCol === -1 ||
    pdfCol === -1 ||
    emailStatusCol === -1
  ) {

    throw new Error(
      'Required email columns are missing.'
    );
  }

  /*
   * Keep this at 20 while testing.
   */

  const EMAIL_BATCH_SIZE = 20;

  /*
   * Never try to send more than today's remaining quota.
   */

  const quotaLeft =
    MailApp.getRemainingDailyQuota();

  const limit =
    Math.min(
      EMAIL_BATCH_SIZE,
      quotaLeft
    );

  if (limit < 1) {

    Logger.log(
      'No email quota left today. Try again tomorrow.'
    );

    return;
  }

  let sent = 0;

  /*
   * Status updates are written in batches
   * (every 10 rows and at the end).
   */

  const updatedRows = [];

  function flushStatus() {

    writeColumnRows_(
      sheet,
      emailStatusCol,
      updatedRows,
      function (r) {
        return data[r][emailStatusCol];
      }
    );

    updatedRows.length = 0;
  }

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      sent >= limit
    ) {
      break;
    }

    if (
      Date.now() - startTime >
      MAX_RUNTIME_MS
    ) {

      Logger.log(
        'Stopped early to stay inside the Apps Script ' +
        'time limit. Run again to continue.'
      );

      break;
    }

    const name =
      String(
        data[i][nameCol] || ''
      ).trim();

    const email =
      String(
        data[i][emailCol] || ''
      ).trim();

    const certificateID =
      String(
        data[i][idCol] || ''
      ).trim();

    const pdfURL =
      String(
        data[i][pdfCol] || ''
      ).trim();

    const emailStatus =
      String(
        data[i][emailStatusCol] || ''
      ).trim();

    if (
      !name ||
      !email ||
      !certificateID ||
      !pdfURL
    ) {
      continue;
    }

    if (
      emailStatus === 'SENT'
    ) {
      continue;
    }

    try {

      const fileId =
        extractDriveFileId(
          pdfURL
        );

      const pdfFile =
        DriveApp.getFileById(
          fileId
        );

      const subject =
        'Your Certificate - ' +
        certificateID;

      const body =
        'Dear ' +
        name +
        ',\n\n' +

        'Congratulations!\n\n' +

        'Please find your certificate attached ' +
        'to this email.\n\n' +

        'Certificate ID: ' +
        certificateID +
        '\n\n' +

        'You can verify the certificate by ' +
        'scanning the QR code on the certificate.\n\n' +

        'Regards';

      MailApp.sendEmail({

        to: email,

        subject: subject,

        body: body,

        attachments: [
          pdfFile.getBlob()
        ]

      });

      data[i][emailStatusCol] =
        'SENT';

      updatedRows.push(i);

      sent++;

    } catch (error) {

      data[i][emailStatusCol] =
        'ERROR: ' +
        error.message;

      updatedRows.push(i);
    }

    if (updatedRows.length >= 10) {
      flushStatus();
    }
  }

  flushStatus();

  SpreadsheetApp.flush();

  Logger.log(
    'Emails sent: ' +
    sent
  );
}


/****************************************************
 * EXTRACT DRIVE FILE ID
 ****************************************************/

function extractDriveFileId(
  url
) {

  const match =
    url.match(
      /[-\w]{25,}/
    );

  if (!match) {

    throw new Error(
      'Could not find Drive file ID.'
    );
  }

  return match[0];
}


/****************************************************
 * WEB APP ENTRY POINT
 ****************************************************/

function doGet(e) {

  const certificateID =
    e &&
    e.parameter &&
    e.parameter.id
      ? String(
          e.parameter.id
        ).trim()
      : '';

  if (certificateID) {

    return verifyCertificate(
      certificateID
    );
  }

  return HtmlService
    .createHtmlOutputFromFile(
      'Index'
    )
    .setTitle(
      'Certificate Verification'
    );
}


/****************************************************
 * VERIFY CERTIFICATE
 ****************************************************/

function verifyCertificate(
  certificateID
) {

  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.STUDENT_SHEET
    );

  if (!sheet) {

    return HtmlService
      .createHtmlOutput(
        '<h2>Configuration Error</h2>' +
        '<p>Students sheet not found.</p>'
      );
  }

  const data =
    sheet
      .getDataRange()
      .getValues();

  const headers =
    data[0];

  const idCol =
    headers.indexOf(
      'Certificate_ID'
    );

  const nameCol =
    headers.indexOf(
      'Name'
    );

  const streamCol =
    headers.indexOf(
      'Stream'
    );

  const acmCol =
    headers.indexOf(
      'Academic_Year'
    );

  const dateCol =
    headers.indexOf(
      'IssueDate'
    );

  const pdfCol =
    headers.indexOf(
      'PDF_URL'
    );

  const statusCol =
    headers.indexOf(
      'Status'
    );

  if (
    idCol === -1 ||
    nameCol === -1 ||
    streamCol === -1 ||
    dateCol === -1
  ) {

    return HtmlService
      .createHtmlOutput(
        '<h2>Configuration Error</h2>'
      );
  }

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const rowID =
      String(
        data[i][idCol] || ''
      ).trim();

    if (
      rowID.toUpperCase() ===
      certificateID.toUpperCase()
    ) {

      const name =
        escapeHTML(
          String(
            data[i][nameCol] || ''
          )
        );

      const stream =
        escapeHTML(
          String(
            data[i][streamCol] || ''
          )
        );

      const date =
        escapeHTML(
          formatDate(
            data[i][dateCol]
          )
        );

      const status =
        statusCol >= 0
          ? String(
              data[i][statusCol] || ''
            ).trim()
          : 'GENERATED';

      const pdf =
        pdfCol >= 0
          ? String(
              data[i][pdfCol] || ''
            ).trim()
          : '';

      if (
        status !== 'GENERATED'
      ) {

        return HtmlService
          .createHtmlOutput(
            invalidCertificatePage(
              certificateID,
              'This certificate is not currently valid.'
            )
          )
          .setTitle(
            'Certificate Verification'
          );
      }

      return HtmlService
        .createHtmlOutput(
          validCertificatePage(
            certificateID,
            name,
            stream,
            date,
            pdf
          )
        )
        .setTitle(
          'Certificate Verified'
        );
    }
  }

  return HtmlService
    .createHtmlOutput(
      invalidCertificatePage(
        certificateID,
        'Certificate not found.'
      )
    )
    .setTitle(
      'Certificate Verification'
    );
}


/****************************************************
 * EMAIL LOOKUP
 ****************************************************/

function findCertificateByEmail(
  email
) {

  email =
    String(
      email || ''
    )
    .trim()
    .toLowerCase();

  if (!email) {

    return {
      found: false,
      message:
        'Please enter your email address.'
    };
  }

  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      CONFIG.STUDENT_SHEET
    );

  const data =
    sheet
      .getDataRange()
      .getValues();

  const headers =
    data[0];

  const nameCol =
    headers.indexOf(
      'Name'
    );

  const emailCol =
    headers.indexOf(
      'Email'
    );

  const streamCol =
    headers.indexOf(
      'Stream'
    );

  const dateCol =
    headers.indexOf(
      'IssueDate'
    );

  const idCol =
    headers.indexOf(
      'Certificate_ID'
    );

  const pdfCol =
    headers.indexOf(
      'PDF_URL'
    );

  if (
    nameCol === -1 ||
    emailCol === -1 ||
    streamCol === -1 ||
    dateCol === -1 ||
    idCol === -1 ||
    pdfCol === -1
  ) {

    return {
      found: false,
      message:
        'System configuration error.'
    };
  }

  const results = [];

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const rowEmail =
      String(
        data[i][emailCol] || ''
      )
      .trim()
      .toLowerCase();

    if (
      rowEmail === email
    ) {

      results.push({

        name:
          String(
            data[i][nameCol] || ''
          ),

        stream:
          String(
            data[i][streamCol] || ''
          ),

        date:
          formatDate(
            data[i][dateCol]
          ),

        certificateID:
          String(
            data[i][idCol] || ''
          ),

        pdfURL:
          String(
            data[i][pdfCol] || ''
          )
      });
    }
  }

  if (
    results.length === 0
  ) {

    return {
      found: false,
      message:
        'No certificate was found for this email address.'
    };
  }

  return {
    found: true,
    certificates: results
  };
}


/****************************************************
 * VALID CERTIFICATE PAGE
 ****************************************************/

function validCertificatePage(
  certificateID,
  name,
  stream,
  date,
  pdf
) {

  let downloadButton = '';

  if (pdf) {

    downloadButton =
      '<p>' +

      '<a class="button" href="' +
      escapeHTML(pdf) +
      '" target="_blank">' +

      'View / Download Certificate' +

      '</a>' +

      '</p>';
  }

  return `
<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<style>

body {
  font-family: Arial, sans-serif;
  background: #f4f6f8;
  padding: 30px;
  margin: 0;
}

.card {
  max-width: 650px;
  margin: 40px auto;
  background: white;
  padding: 35px;
  border-radius: 12px;
  box-shadow: 0 3px 15px rgba(0,0,0,.12);
}

.valid {
  color: #16803c;
  font-size: 24px;
  font-weight: bold;
}

.label {
  color: #666;
  font-size: 13px;
  margin-top: 20px;
}

.value {
  font-size: 18px;
  margin-top: 5px;
}

.button {
  display: inline-block;
  padding: 12px 20px;
  background: #1a73e8;
  color: white;
  text-decoration: none;
  border-radius: 6px;
}

</style>

</head>

<body>

<div class="card">

<div class="valid">
CERTIFICATE VALID
</div>

<div class="label">
Certificate ID
</div>

<div class="value">
${escapeHTML(certificateID)}
</div>

<div class="label">
Student Name
</div>

<div class="value">
${name}
</div>

<div class="label">
Stream
</div>

<div class="value">
${stream}
</div>

<div class="label">
Issue Date
</div>

<div class="value">
${date}
</div>

${downloadButton}

</div>

</body>

</html>
`;
}


/****************************************************
 * INVALID CERTIFICATE PAGE
 ****************************************************/

function invalidCertificatePage(
  certificateID,
  message
) {

  return `
<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<style>

body {
  font-family: Arial, sans-serif;
  background: #f4f6f8;
  padding: 30px;
}

.card {
  max-width: 650px;
  margin: 40px auto;
  background: white;
  padding: 35px;
  border-radius: 12px;
}

.invalid {
  color: #c5221f;
  font-size: 24px;
  font-weight: bold;
}

</style>

</head>

<body>

<div class="card">

<div class="invalid">
CERTIFICATE INVALID
</div>

<p>
${escapeHTML(message)}
</p>

<p>
Certificate ID:
<strong>
${escapeHTML(certificateID)}
</strong>
</p>

</div>

</body>

</html>
`;
}


/****************************************************
 * ESCAPE HTML
 ****************************************************/

function escapeHTML(
  value
) {

  return String(value)

    .replace(
      /&/g,
      '&amp;'
    )

    .replace(
      /</g,
      '&lt;'
    )

    .replace(
      />/g,
      '&gt;'
    )

    .replace(
      /"/g,
      '&quot;'
    )

    .replace(
      /'/g,
      '&#039;'
    );
}
