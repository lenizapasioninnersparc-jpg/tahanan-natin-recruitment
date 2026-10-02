/**
 * TAHANAN NATIN × INNERSPARC
 * RECRUITMENT LANDING PAGE BACKEND
 *
 * FLOW:
 * 1. Applicant submits qualification form
 * 2. Application is saved to Google Sheet
 * 3. Admin receives complete applicant information
 * 4. Applicant receives separate thank-you email
 *
 * NOTE:
 * The 5-day recruitment email automation is intentionally NOT included yet.
 */

// ============================================================
// CONFIGURATION
// ============================================================

const CONFIG = {
  // Admin / owner email
  ADMIN_EMAIL: 'lenizapasion.innersparc@gmail.com',

  // Display name for applicant emails
  SENDER_NAME: 'Leniza Pasion | Tahanan Natin',

  // Google Sheet
  //
  // OPTION A:
  // If this Apps Script project is BOUND to the Google Sheet,
  // leave this blank.
  //
  // OPTION B:
  // If this is a standalone Apps Script project, paste the
  // Spreadsheet ID here.
  //
  // Example:
  // SPREADSHEET_ID: '1AbCdEfGhIjKlMnOpQrStUvWxYz'
  SPREADSHEET_ID: '',

  // Sheet tab name
  SHEET_NAME: 'Recruitment Applicants',

  // Branding images from the existing landing page
  LOGO_TAHANAN_NATIN: 'https://i.imgur.com/1zmJxPa.png',
  LOGO_INNERSPARC: 'https://i.imgur.com/sPeeLaB.png',
  LENIZA_IMAGE: 'https://i.imgur.com/e3jmDE6.png',

  // Additional existing landing-page images
  REWARD_SPA_IMAGE: 'https://i.imgur.com/qcs65Kp.jpeg',
  MILESTONE_13_IMAGE: 'https://i.imgur.com/7C0K1CN.jpeg',
  REWARD_RECOGNITION_IMAGE: 'https://i.imgur.com/JoW4ue5.jpeg',
  REWARD_TEAM_IMAGE: 'https://i.imgur.com/WIPV68J.jpeg',

  // Applicant email pain-point card image
  PAIN_POINT_IMAGE: 'https://i.imgur.com/OPiaiwJ.png',

  // Agent portal link for the button at the bottom of the applicant
  // email. (utm tags let you see clicks coming from this email.)
  AGENT_PORTAL_URL: 'https://innersparcagentportal.vercel.app/?utm_source=applicant_email&utm_medium=email&utm_campaign=recruitment_thank_you',

  // Admin email subject
  ADMIN_SUBJECT_PREFIX: 'NEW RECRUITMENT APPLICANT'
};


// ============================================================
// WEB APP
// ============================================================

function doGet(e) {
  return jsonResponse_({
    success: true,
    service: 'Tahanan Natin recruitment backend',
    message: 'Backend is online. Submit applications through the Vercel website.'
  });
}

/**
 * Receives applications forwarded by the Vercel API function.
 * The shared secret is stored in Apps Script Script Properties and
 * Vercel Environment Variables; it is never placed in browser code.
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const expectedSecret = PropertiesService.getScriptProperties()
      .getProperty('VERCEL_API_SECRET');

    if (!expectedSecret || body.apiSecret !== expectedSecret) {
      return jsonResponse_({
        success: false,
        message: 'Unauthorized request.'
      });
    }

    if (!body.data || typeof body.data !== 'object') {
      return jsonResponse_({
        success: false,
        message: 'Application data is missing.'
      });
    }

    const result = handleRecruitmentQualification(body.data);
    return jsonResponse_(result);

  } catch (err) {
    console.error('Vercel submission error: ' + (err && err.stack ? err.stack : err));
    return jsonResponse_({
      success: false,
      message: (err && err.message) || 'The application could not be processed.'
    });
  }
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// FIRST-TIME SETUP
// ============================================================

/**
 * Run this ONCE manually from Apps Script.
 *
 * It creates the recruitment applicant sheet and headers.
 */
function setupRecruitmentSheet() {
  const spreadsheet = getRecruitmentSpreadsheet_();

  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  }

  const headers = [
    'Timestamp',
    'Applicant Name',
    'Mobile',
    'Email',

    'Area',
    'Near Molino',
    'Barangay / Exact Address',

    'Age',
    'Current Situation',

    'Commission Awareness',
    'Work Preference',
    'Hours Per Week',

    'Why Do You Want To Join?',

    'UTM Source',
    'UTM Medium',
    'UTM Campaign',

    'Consent',
    'Application Status',
    'Applicant Email Status'
  ];

  // If sheet is empty, add headers
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    // Ensure headers are present/replaced with the current structure
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Header formatting
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0F1D3A')
    .setFontColor('#FFFFFF')
    .setWrap(true);

  sheet.setFrozenRows(1);

  // Basic column widths
  const widths = [
    160, // Timestamp
    180, // Name
    130, // Mobile
    220, // Email
    150, // Area
    180, // Near Molino
    260, // Address
    80,  // Age
    190, // Situation
    280, // Commission
    220, // Work preference
    120, // Hours
    400, // Why join
    130, // UTM source
    130, // UTM medium
    160, // UTM campaign
    100, // Consent
    140, // Status
    220  // Applicant email status
  ];

  widths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  return 'Recruitment Sheet is ready: ' + spreadsheet.getUrl();
}


/**
 * ONE-CLICK SETUP
 *
 * Run this function once from the Apps Script editor after pasting
 * this file into code.gs. It prepares the recruitment Sheet and
 * stores a small setup marker so you can verify that setup ran.
 *
 * IMPORTANT:
 * - This does NOT send any email.
 * - This does NOT create the 5-day automation.
 * - This does NOT create any time-based triggers.
 * - If the project is standalone, a recruitment spreadsheet is created automatically.
 */
function setupEverything() {
  const spreadsheet = getRecruitmentSpreadsheet_();

  // 1. Validate the core configuration.
  if (!CONFIG.ADMIN_EMAIL || CONFIG.ADMIN_EMAIL.indexOf('@') === -1) {
    throw new Error('Please check CONFIG.ADMIN_EMAIL.');
  }

  if (!CONFIG.SENDER_NAME) {
    throw new Error('Please check CONFIG.SENDER_NAME.');
  }

  if (!CONFIG.SHEET_NAME) {
    throw new Error('Please check CONFIG.SHEET_NAME.');
  }

  // 2. Create / refresh the recruitment applicant sheet.
  const sheetUrl = setupRecruitmentSheet();

  // 3. Store a simple one-time setup marker.
  // This is only for visibility/debugging; it does not create triggers.
  PropertiesService.getScriptProperties().setProperties({
    RECRUITMENT_SETUP_COMPLETE: 'true',
    RECRUITMENT_SETUP_AT: new Date().toISOString(),
    RECRUITMENT_SHEET_NAME: CONFIG.SHEET_NAME,
    RECRUITMENT_SPREADSHEET_ID: spreadsheet.getId()
  }, true);

  // 4. Return a clear setup summary in the Apps Script execution log.
  const result = {
    setupComplete: true,
    spreadsheetName: spreadsheet.getName(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sheetName: CONFIG.SHEET_NAME,
    adminEmail: CONFIG.ADMIN_EMAIL,
    applicantSenderName: CONFIG.SENDER_NAME,
    fiveDayAutomation: 'NOT CREATED',
    timeBasedTriggers: 'NOT CREATED',
    emailsSent: false,
    message: 'Setup complete. The recruitment Sheet is ready. No automation or test emails were started.'
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}


// ============================================================
// MAIN FORM HANDLER
// ============================================================

/**
 * Called by the landing page:
 *
 * doPost() calls handleRecruitmentQualification(data) after verifying the Vercel API secret.
 */
function handleRecruitmentQualification(data) {

  if (!data || typeof data !== 'object') {
    throw new Error('No application data was received.');
  }

  // ----------------------------------------------------------
  // Normalize incoming fields
  // ----------------------------------------------------------

  const applicant = {
    timestamp: new Date(),

    fullName: clean_(data.fullName),
    mobile: clean_(data.mobile),
    email: clean_(data.email),

    area: clean_(data.area),
    nearMolino: clean_(data.nearMolino),
    exactAddress: clean_(data.exactAddress),

    age: clean_(data.age),
    currentSituation: clean_(data.currentSituation),

    commissionAwareness: clean_(data.commissionAwareness),
    workPreference: clean_(data.workPreference),
    hoursPerWeek: clean_(data.hoursPerWeek),

    whyJoin: clean_(data.whyJoin),

    utmSource: clean_(data.utmSource),
    utmMedium: clean_(data.utmMedium),
    utmCampaign: clean_(data.utmCampaign),

    // The form can only reach this function after the required
    // consent checkbox passes browser validation.
    consent: 'Yes',

    status: 'New'
  };


  // ----------------------------------------------------------
  // Server-side validation
  // ----------------------------------------------------------

  Logger.log('RECRUITMENT FORM RECEIVED: ' + applicant.fullName + ' <' + applicant.email + '>');

  validateApplication_(applicant);


  // ----------------------------------------------------------
  // Save to Google Sheet
  // ----------------------------------------------------------

  const rowNumber = saveApplicationToSheet_(applicant);


  // ----------------------------------------------------------
  // Send applicant thank-you email
  // ----------------------------------------------------------

  // If this fails, the row is marked RETRY and the 1-minute trigger
  // (see installEmailTrigger) sends it again automatically.
  try {
    sendApplicantThankYou_(applicant);
    setEmailStatus_(rowNumber, 'Sent ' + new Date().toISOString());
    Logger.log('Applicant thank-you email sent to ' + applicant.email);
  } catch (applicantError) {
    console.error(
      'APPLICANT EMAIL FAILED for ' + applicant.email + ': ' +
      applicantError.message
    );
    setEmailStatus_(rowNumber, 'RETRY 0: ' + applicantError.message);
  }


  // ----------------------------------------------------------
  // Send admin notification
  // ----------------------------------------------------------

  // Wrapped so that a failed admin email can never stop the
  // applicant email from being sent (and vice versa).
  try {
    sendAdminNotification_(applicant);
  } catch (adminError) {
    console.error('ADMIN EMAIL FAILED: ' + adminError.message);
  }


  // ----------------------------------------------------------
  // Return result to landing page
  // ----------------------------------------------------------

  return {
    success: true,
    message:
      'Thank you, ' +
      applicant.fullName +
      '! We received your application and will review your answers.'
  };
}


// ============================================================
// VALIDATION
// ============================================================

function validateApplication_(applicant) {

  const requiredFields = [
    ['fullName', 'Full name'],
    ['mobile', 'Mobile number'],
    ['email', 'Email'],
    ['area', 'Area'],
    ['nearMolino', 'Near Molino answer'],
    ['exactAddress', 'Exact address'],
    ['age', 'Age'],
    ['currentSituation', 'Current situation'],
    ['commissionAwareness', 'Commission awareness'],
    ['workPreference', 'Work preference'],
    ['hoursPerWeek', 'Hours per week'],
    ['whyJoin', 'Why do you want to join?']
  ];

  requiredFields.forEach(function(item) {
    const key = item[0];
    const label = item[1];

    if (!applicant[key]) {
      throw new Error(label + ' is required.');
    }
  });


  // Age
  const age = Number(applicant.age);

  if (!Number.isFinite(age) || age < 18 || age > 100) {
    throw new Error('Applicant must be at least 18 years old.');
  }


  // Email
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  if (!emailPattern.test(applicant.email)) {
    throw new Error('Please provide a valid email address.');
  }


  // Commission requirement
  if (
    applicant.commissionAwareness ===
    "No, I'm looking for a fixed salary"
  ) {
    throw new Error(
      'This recruitment opportunity is commission-based and does not offer a fixed salary.'
    );
  }
}


// ============================================================
// GOOGLE SHEET
// ============================================================

function saveApplicationToSheet_(applicant) {

  const spreadsheet = getRecruitmentSpreadsheet_();

  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    setupRecruitmentSheet();
    sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  }

  const row = [
    applicant.timestamp,
    applicant.fullName,
    applicant.mobile,
    applicant.email,

    applicant.area,
    applicant.nearMolino,
    applicant.exactAddress,

    applicant.age,
    applicant.currentSituation,

    applicant.commissionAwareness,
    applicant.workPreference,
    applicant.hoursPerWeek,

    applicant.whyJoin,

    applicant.utmSource,
    applicant.utmMedium,
    applicant.utmCampaign,

    applicant.consent,
    applicant.status,

    '' // Applicant Email Status (filled after the email is sent)
  ];

  sheet.appendRow(row);

  // Format timestamp
  const lastRow = sheet.getLastRow();

  sheet
    .getRange(lastRow, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // Wrap longer text
  sheet
    .getRange(lastRow, 1, 1, row.length)
    .setWrap(true);

  return lastRow;
}


// ============================================================
// ADMIN EMAIL
// ============================================================

function sendAdminNotification_(applicant) {

  const subject =
    CONFIG.ADMIN_SUBJECT_PREFIX +
    ' — ' +
    applicant.fullName;


  const htmlBody = `
    <div style="
      margin:0;
      padding:30px 15px;
      background:#f5f2eb;
      font-family:Arial,Helvetica,sans-serif;
      color:#172033;
    ">

      <div style="
        max-width:720px;
        margin:0 auto;
        background:#ffffff;
        border-radius:18px;
        overflow:hidden;
        border:1px solid #e4dfd4;
      ">

        <!-- HEADER -->

        <div style="
          background:#0F1D3A;
          padding:25px 30px;
          text-align:center;
        ">

          <img
            src="${CONFIG.LOGO_TAHANAN_NATIN}"
            alt="Tahanan Natin"
            style="
              width:72px;
              height:72px;
              object-fit:contain;
              margin:0 auto 10px;
            "
          >

          <div style="
            color:#D4A843;
            font-size:12px;
            font-weight:bold;
            letter-spacing:2px;
            text-transform:uppercase;
          ">
            Recruitment Notification
          </div>

          <h1 style="
            color:#ffffff;
            font-size:26px;
            margin:8px 0 0;
          ">
            New Applicant
          </h1>

        </div>


        <!-- INTRO -->

        <div style="padding:28px 30px 10px;">

          <p style="
            font-size:15px;
            line-height:1.7;
            margin:0;
          ">
            A new applicant has completed the Tahanan Natin × InnerSparc
            recruitment qualification form.
          </p>

          <p style="
            font-size:13px;
            color:#68758A;
            margin-top:8px;
          ">
            Submitted:
            ${formatDateTime_(applicant.timestamp)}
          </p>

        </div>


        <!-- APPLICANT INFORMATION -->

        ${adminSection_(
          'Applicant Information',
          [
            ['Full Name', applicant.fullName],
            ['Mobile Number', applicant.mobile],
            ['Email', applicant.email],
            ['Age', applicant.age],
            ['Application Status', applicant.status]
          ]
        )}


        <!-- LOCATION -->

        ${adminSection_(
          'Location',
          [
            ['Area', applicant.area],
            ['Near Molino, Bacoor?', applicant.nearMolino],
            ['Barangay / Exact Address', applicant.exactAddress]
          ]
        )}


        <!-- CURRENT SITUATION -->

        ${adminSection_(
          'Current Situation',
          [
            ['Current Situation', applicant.currentSituation],
            ['Commission Awareness', applicant.commissionAwareness],
            ['How They Want To Start', applicant.workPreference],
            ['Available Hours Per Week', applicant.hoursPerWeek]
          ]
        )}


        <!-- WHY JOIN -->

        <div style="padding:10px 30px 25px;">

          <div style="
            font-size:11px;
            color:#B88A25;
            font-weight:bold;
            letter-spacing:1px;
            text-transform:uppercase;
            margin-bottom:8px;
          ">
            Why Do You Want To Join?
          </div>

          <div style="
            background:#F8F4ED;
            border-left:3px solid #D4A843;
            padding:17px;
            border-radius:8px;
            font-size:14px;
            line-height:1.7;
            white-space:pre-wrap;
          ">
            ${escapeHtml_(applicant.whyJoin)}
          </div>

        </div>


        <!-- MARKETING DATA -->

        ${adminSection_(
          'Marketing / Tracking',
          [
            ['UTM Source', applicant.utmSource || '—'],
            ['UTM Medium', applicant.utmMedium || '—'],
            ['UTM Campaign', applicant.utmCampaign || '—'],
            ['Consent', applicant.consent]
          ]
        )}


        <!-- FOOTER -->

        <div style="
          padding:22px 30px;
          background:#091329;
          color:#8793A8;
          font-size:11px;
          line-height:1.6;
        ">
          TAHANAN NATIN × INNERSPARC<br>
          Recruitment Applicant Notification
        </div>

      </div>
    </div>
  `;


  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: subject,
    htmlBody: htmlBody,
    body: buildAdminPlainText_(applicant),
    name: CONFIG.SENDER_NAME
  });
}


// ============================================================
// APPLICANT THANK-YOU EMAIL
// ============================================================

function sendApplicantThankYou_(applicant) {

  const subject =
    'Natanggap na namin ang application mo, ' +
    firstName_(applicant.fullName) +
    '! — Tahanan Natin';


  if (!CONFIG.AGENT_PORTAL_URL) {
    Logger.log('WARNING: CONFIG.AGENT_PORTAL_URL is empty, so the email button has no link yet.');
  }

  const htmlBody = `
    <div style="
      margin:0;
      padding:25px 12px;
      background:#f5f2eb;
      font-family:Arial,Helvetica,sans-serif;
      color:#172033;
    ">

      <div style="
        max-width:600px;
        margin:0 auto;
        background:#ffffff;
        border-radius:20px;
        overflow:hidden;
        border:1px solid #e5dfd4;
      ">


        <!-- BRAND HEADER -->

        <div style="
          background:#0F1D3A;
          padding:22px 25px;
          text-align:center;
        ">

          <div style="
            display:inline-block;
            vertical-align:middle;
          ">

            <img
              src="${CONFIG.LOGO_TAHANAN_NATIN}"
              alt="Tahanan Natin"
              style="
                width:65px;
                height:65px;
                object-fit:contain;
                vertical-align:middle;
                margin-right:10px;
              "
            >

          </div>

          <div style="
            display:inline-block;
            vertical-align:middle;
            text-align:left;
          ">

            <div style="
              color:#ffffff;
              font-size:17px;
              font-weight:bold;
            ">
              TAHANAN NATIN
            </div>

            <div style="
              color:#D4A843;
              font-size:10px;
              letter-spacing:1.5px;
              margin-top:3px;
            ">
              REAL ESTATE PROJECT SELLING
            </div>

          </div>

        </div>


        <!-- LENIZA IMAGE -->

        <div style="
          background:#0F1D3A;
          padding:0 25px 25px;
          text-align:center;
        ">

          <img
            src="${CONFIG.LENIZA_IMAGE}"
            alt="Leniza Pasion"
            style="
              width:100%;
              max-width:500px;
              height:auto;
              max-height:520px;
              object-fit:cover;
              object-position:center top;
              border-radius:16px;
              border:1px solid rgba(212,168,67,.35);
            "
          >

        </div>


        <!-- MESSAGE -->

        <div style="
          padding:35px 32px 0;
        ">

          <div style="
            color:#B88A25;
            font-size:11px;
            font-weight:bold;
            letter-spacing:1.5px;
            text-transform:uppercase;
            margin-bottom:8px;
          ">
            Thank You For Applying
          </div>

          <h1 style="
            color:#0F1D3A;
            font-family:Georgia,'Times New Roman',serif;
            font-size:31px;
            line-height:1.2;
            margin:0 0 18px;
          ">
            Hello ${escapeHtml_(firstName_(applicant.fullName))},
          </h1>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Leniza from Tahanan Natin here.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Gusto ko lang mag-thank you agad, kasi natanggap na namin ang application mo.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Alam kong hindi laging madali mag-fill out ng form, lalo na kung busy ka sa work, school, o sa ibang priorities.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Kaya thank you sa oras na ibinigay mo para sagutin ang qualification questions namin.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            At thank you rin sa pag-share ng konti about yourself.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Hindi yun maliit na bagay para sa akin.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Kasi alam kong may mga bagay ka ring iniisip bago ka nag-submit.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Baka nagdalawang-isip ka pa nga.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Baka tinanong mo rin ang sarili mo, “Kaya ko ba ’to?”
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Okay lang yun. Normal lang yun.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Ang importante, nag-take ka ng first step.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            At gusto kong i-acknowledge yun.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Natanggap na namin ang application mo para sa <strong>Tahanan Natin × InnerSparc</strong> real estate project-selling opportunity.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Baka iniisip mo ngayon, “Ano na kaya next?”
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Okay lang yan. Normal lang mag-isip ng ganyan.
          </p>

          <div style="
            margin:30px 0 32px;
            padding:19px 20px;
            border-left:3px solid #D4A843;
            background:#F8F4ED;
            border-radius:8px;
          ">

            <strong style="
              display:block;
              color:#0F1D3A;
              font-size:15px;
              margin-bottom:6px;
            ">
              What happens next?
            </strong>

            <span style="
              color:#68758A;
              font-size:13px;
              line-height:1.7;
            ">
              Nire-review na ng team namin ang mga sinagot mo,
              lalo na ang location, availability, current situation,
              at interest mo sa commission-based setup.
            </span>

          </div>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Wala ka pang kailangang gawin for now.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Hindi mo kailangang mag-reply, mag-follow up, o mag-abang sa bawat minuto.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Ang team na mismo ang mag-message sa'yo kapag tapos na naming i-review.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Habang naghihintay, tuloy lang ang normal na araw mo.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Mag-work, mag-aral, o magpahinga.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Kumain ka nang maayos, at huwag mong pabayaan ang sarili mo.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Kasi hindi naman ito something na kailangang madaliin.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Walang deadline na hinahabol ka dito.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Ang gusto lang namin, maging komportable ka sa bawat hakbang.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            At kung may mga tanong ka habang naghihintay, normal lang din yun.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Itabi mo lang muna sa isip mo, para handa ka kapag nag-usap na tayo.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Okay lang kung sideline lang muna ito para sa'yo.
          </p>

        </div>


        <!-- PAIN POINT CARD -->

        <div style="
          margin:0 0 28px;
          background:#0F1D3A;
        ">

          <div style="
            max-height:380px;
            overflow:hidden;
          ">
            <img
              src="${CONFIG.PAIN_POINT_IMAGE}"
              alt=""
              width="600"
              style="
                display:block;
                width:100%;
                max-width:100%;
                height:auto;
                border:0;
              "
            >
          </div>

          <div style="
            padding:20px 28px 24px;
          ">

            <div style="
              color:#D4A843;
              font-size:11px;
              font-weight:bold;
              letter-spacing:1.5px;
              margin-bottom:10px;
            ">
              01
            </div>

            <h2 style="
              color:#ffffff;
              font-family:Georgia,'Times New Roman',serif;
              font-size:20px;
              line-height:1.35;
              margin:0 0 14px;
            ">
              May work ka na, pero pagod na at wala nang time para sa extra income?
            </h2>

            <p style="
              color:#C9D2E3;
              font-size:13px;
              line-height:1.7;
              margin:0 0 12px;
            ">
              Gising ka nang maaga, uwi ka nang gabi, at pagdating sa bahay, ubos na ang lakas mo.
            </p>

            <p style="
              color:#C9D2E3;
              font-size:13px;
              line-height:1.7;
              margin:0 0 12px;
            ">
              Gusto mo sanang may extra income, pero parang wala nang oras o energy na natitira para sa iba.
            </p>

            <p style="
              color:#C9D2E3;
              font-size:13px;
              line-height:1.7;
              margin:0 0 12px;
            ">
              Hindi ka tamad. Sadyang puno lang talaga ang araw mo.
            </p>

            <p style="
              color:#C9D2E3;
              font-size:13px;
              line-height:1.7;
              margin:0 0 12px;
            ">
              Kaya madalas, naiiwan na lang sa isip ang “sana may extra akong kita,” habang lumilipas ang mga buwan.
            </p>

            <p style="
              color:#C9D2E3;
              font-size:14px;
              line-height:1.8;
              margin:0;
            ">
              Kung ganito ang pakiramdam mo, hindi ka nag-iisa. At baka ito na ang dahilan kung bakit sideline muna ang pwedeng simula.
            </p>

          </div>

        </div>


        <!-- MESSAGE (CONTINUED) -->

        <div style="
          padding:0 32px 25px;
        ">

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Okay din kung gusto mo siyang seryosohin in time.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Walang pressure, at walang hard selling dito.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Gusto lang naming makilala ka muna, at makita kung fit ba tayo sa isa't isa.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Kasi hindi lahat ng opportunity ay para sa lahat.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            At okay lang din yun.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Mas gusto naming makasama yung mga taong totoong interesado, at handang magsipag at magtiyaga.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Ang importante, may gana kang matuto.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            At may pagkakataon tayong magkakilala nang maayos.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            So kung ikaw yun, masaya ako na nandito ka.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Masaya ako na nakilala ka, kahit sa pamamagitan lang ng form na ito.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Salamat sa tiwala, kahit konti pa lang ang alam mo tungkol sa amin.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0 0 20px;
          ">
            Hanggang sa susunod nating pag-uusap.
          </p>

          <p style="
            color:#4F5B6E;
            font-size:15px;
            line-height:1.8;
            margin:0;
          ">
            Thank you ulit, ${escapeHtml_(firstName_(applicant.fullName))}. Talk soon!
          </p>

        </div>


        <!-- BRANDING / IMAGES -->

        <div style="
          padding:5px 25px 30px;
        ">

          <div style="
            color:#B88A25;
            font-size:10px;
            font-weight:bold;
            letter-spacing:1.5px;
            text-transform:uppercase;
            text-align:center;
            margin-bottom:12px;
          ">
            Tahanan Natin × InnerSparc
          </div>

          <!-- Row cells share one height: the first picture stretches to
               match the two taller pictures (cover-fit, no distortion). -->
          <table
            role="presentation"
            width="100%"
            cellpadding="0"
            cellspacing="0"
            border="0"
            style="width:100%;"
          >
            <tr>
              <td
                width="33%"
                valign="top"
                bgcolor="#EEF1F6"
                background="${CONFIG.MILESTONE_13_IMAGE}"
                style="
                  background-image:url('${CONFIG.MILESTONE_13_IMAGE}');
                  background-size:cover;
                  background-position:center center;
                  background-repeat:no-repeat;
                  border-radius:10px;
                  font-size:0;
                  line-height:0;
                "
              >
                &nbsp;
              </td>
              <td width="8" style="width:8px;font-size:0;line-height:0;">&nbsp;</td>
              <td width="33%" valign="top">
                <img
                  src="${CONFIG.REWARD_RECOGNITION_IMAGE}"
                  alt="Team recognition"
                  width="170"
                  style="
                    display:block;
                    width:100%;
                    height:auto;
                    border:0;
                    border-radius:10px;
                  "
                >
              </td>
              <td width="8" style="width:8px;font-size:0;line-height:0;">&nbsp;</td>
              <td width="33%" valign="top">
                <img
                  src="${CONFIG.REWARD_TEAM_IMAGE}"
                  alt="Team"
                  width="170"
                  style="
                    display:block;
                    width:100%;
                    height:auto;
                    border:0;
                    border-radius:10px;
                  "
                >
              </td>
            </tr>
          </table>


          <!-- CTA: AGENT PORTAL -->

          <div style="
            margin-top:26px;
            text-align:center;
          ">

            <p style="
              color:#4F5B6E;
              font-size:14px;
              line-height:1.7;
              margin:0 0 16px;
            ">
              Habang naghihintay, pwede mong silipin ang agent portal namin.<br>
              Pwede ka ring mag-inquire doon kung may tanong ka.
            </p>

            <table
              role="presentation"
              cellpadding="0"
              cellspacing="0"
              border="0"
              style="margin:0 auto;"
            >
              <tr>
                <td style="
                  background:#D4A843;
                  border-radius:8px;
                  padding:14px 28px;
                ">
                  <a
                    href="${CONFIG.AGENT_PORTAL_URL ? CONFIG.AGENT_PORTAL_URL.replace(/&/g, '&amp;') : '#'}"
                    style="
                      color:#0F1D3A;
                      font-size:14px;
                      font-weight:bold;
                      letter-spacing:.5px;
                      text-decoration:none;
                    "
                  >
                    VISIT THE AGENT PORTAL
                  </a>
                </td>
              </tr>
            </table>

          </div>

        </div>


        <!-- SIGNATURE -->

        <div style="
          padding:26px 32px 30px;
          background:#F8F4ED;
          border-top:1px solid #e9e2d6;
        ">

          <p style="
            margin:0;
            color:#68758A;
            font-size:13px;
            line-height:1.7;
          ">
            Warm regards,
          </p>

          <p style="
            margin:4px 0 0;
            color:#0F1D3A;
            font-family:Georgia,'Times New Roman',serif;
            font-size:21px;
            font-weight:bold;
          ">
            Leniza Pasion
          </p>

          <p style="
            margin:2px 0 0;
            color:#B88A25;
            font-size:12px;
            font-weight:bold;
          ">
            Tahanan Natin
          </p>

          <div style="
            margin-top:14px;
          ">

            <img
              src="${CONFIG.LOGO_INNERSPARC}"
              alt="InnerSparc"
              style="
                width:55px;
                height:55px;
                object-fit:contain;
              "
            >

          </div>

        </div>


        <!-- DISCLAIMER -->

        <div style="
          padding:18px 32px;
          background:#F8F4ED;
          border-top:1px solid #e9e2d6;
          color:#8A94A6;
          font-size:11px;
          line-height:1.7;
        ">
          Please remember that submitting the qualification form does not
          guarantee recruitment, income, or acceptance into the team.
          Actual role terms, commission structure, requirements, projects,
          and incentives will be discussed with the team.
        </div>


        <!-- FOOTER -->

        <div style="
          background:#091329;
          padding:18px 25px;
          text-align:center;
          color:#7E8CA6;
          font-size:10px;
          line-height:1.6;
        ">
          TAHANAN NATIN × INNERSPARC<br>
          Real Estate Project Selling Recruitment
        </div>

      </div>
    </div>
  `;


  MailApp.sendEmail({
    to: applicant.email,
    subject: subject,
    htmlBody: htmlBody,
    body: buildApplicantPlainText_(applicant),
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.ADMIN_EMAIL
  });

  console.log(
    'Applicant email sent to ' + applicant.email +
    ' | quota left: ' + MailApp.getRemainingDailyQuota()
  );
}


// ============================================================
// EMAIL HELPERS
// ============================================================

function adminSection_(title, rows) {

  let html = `
    <div style="padding:8px 30px 22px;">

      <div style="
        color:#B88A25;
        font-size:11px;
        font-weight:bold;
        letter-spacing:1px;
        text-transform:uppercase;
        margin-bottom:9px;
      ">
        ${escapeHtml_(title)}
      </div>

      <table style="
        width:100%;
        border-collapse:collapse;
        font-size:13px;
      ">
  `;

  rows.forEach(function(row) {

    html += `
      <tr>

        <td style="
          width:38%;
          padding:9px 10px 9px 0;
          border-bottom:1px solid #eee9df;
          color:#68758A;
          vertical-align:top;
          font-weight:bold;
        ">
          ${escapeHtml_(row[0])}
        </td>

        <td style="
          padding:9px 0;
          border-bottom:1px solid #eee9df;
          color:#172033;
          vertical-align:top;
          line-height:1.55;
          white-space:pre-wrap;
        ">
          ${escapeHtml_(row[1] || '—')}
        </td>

      </tr>
    `;
  });

  html += `
      </table>
    </div>
  `;

  return html;
}


function buildAdminPlainText_(applicant) {

  return [
    'NEW RECRUITMENT APPLICANT',
    '',
    'Applicant Information',
    '----------------------',
    'Name: ' + applicant.fullName,
    'Mobile: ' + applicant.mobile,
    'Email: ' + applicant.email,
    'Age: ' + applicant.age,
    'Status: ' + applicant.status,
    '',
    'Location',
    '--------',
    'Area: ' + applicant.area,
    'Near Molino: ' + applicant.nearMolino,
    'Exact Address: ' + applicant.exactAddress,
    '',
    'Current Situation',
    '-----------------',
    'Current Situation: ' + applicant.currentSituation,
    'Commission Awareness: ' + applicant.commissionAwareness,
    'Work Preference: ' + applicant.workPreference,
    'Hours Per Week: ' + applicant.hoursPerWeek,
    '',
    'Why Do You Want To Join?',
    '-------------------------',
    applicant.whyJoin,
    '',
    'Marketing / Tracking',
    '--------------------',
    'UTM Source: ' + (applicant.utmSource || '—'),
    'UTM Medium: ' + (applicant.utmMedium || '—'),
    'UTM Campaign: ' + (applicant.utmCampaign || '—'),
    'Consent: ' + applicant.consent,
    '',
    'Submitted: ' + formatDateTime_(applicant.timestamp)
  ].join('\n');
}


function buildApplicantPlainText_(applicant) {

  const name = firstName_(applicant.fullName);

  return [
    'Hello ' + name + ',',
    '',
    'Leniza from Tahanan Natin here.',
    '',
    'Gusto ko lang mag-thank you agad, kasi natanggap na namin ang application mo.',
    '',
    'Alam kong hindi laging madali mag-fill out ng form, lalo na kung busy ka sa work, school, o sa ibang priorities.',
    '',
    'Kaya thank you sa oras na ibinigay mo para sagutin ang qualification questions namin.',
    '',
    'At thank you rin sa pag-share ng konti about yourself.',
    '',
    'Hindi yun maliit na bagay para sa akin.',
    '',
    'Kasi alam kong may mga bagay ka ring iniisip bago ka nag-submit.',
    '',
    'Baka nagdalawang-isip ka pa nga.',
    '',
    'Baka tinanong mo rin ang sarili mo, “Kaya ko ba ’to?”',
    '',
    'Okay lang yun. Normal lang yun.',
    '',
    'Ang importante, nag-take ka ng first step.',
    '',
    'At gusto kong i-acknowledge yun.',
    '',
    'Natanggap na namin ang application mo para sa Tahanan Natin × InnerSparc real estate project-selling opportunity.',
    '',
    'Baka iniisip mo ngayon, “Ano na kaya next?”',
    '',
    'Okay lang yan. Normal lang mag-isip ng ganyan.',
    '',
    'WHAT HAPPENS NEXT?',
    '',
    'Nire-review na ng team namin ang mga sinagot mo, lalo na ang location, availability, current situation, at interest mo sa commission-based setup.',
    '',
    'Wala ka pang kailangang gawin for now.',
    '',
    'Hindi mo kailangang mag-reply, mag-follow up, o mag-abang sa bawat minuto.',
    '',
    'Ang team na mismo ang mag-message sa\'yo kapag tapos na naming i-review.',
    '',
    'Habang naghihintay, tuloy lang ang normal na araw mo.',
    '',
    'Mag-work, mag-aral, o magpahinga.',
    '',
    'Kumain ka nang maayos, at huwag mong pabayaan ang sarili mo.',
    '',
    'Kasi hindi naman ito something na kailangang madaliin.',
    '',
    'Walang deadline na hinahabol ka dito.',
    '',
    'Ang gusto lang namin, maging komportable ka sa bawat hakbang.',
    '',
    'At kung may mga tanong ka habang naghihintay, normal lang din yun.',
    '',
    'Itabi mo lang muna sa isip mo, para handa ka kapag nag-usap na tayo.',
    '',
    'Okay lang kung sideline lang muna ito para sa\'yo.',
    '',
    'May work ka na, pero pagod na at wala nang time para sa extra income?',
    '',
    'Gising ka nang maaga, uwi ka nang gabi, at pagdating sa bahay, ubos na ang lakas mo.',
    '',
    'Gusto mo sanang may extra income, pero parang wala nang oras o energy na natitira para sa iba.',
    '',
    'Hindi ka tamad. Sadyang puno lang talaga ang araw mo.',
    '',
    'Kaya madalas, naiiwan na lang sa isip ang “sana may extra akong kita,” habang lumilipas ang mga buwan.',
    '',
    'Kung ganito ang pakiramdam mo, hindi ka nag-iisa. At baka ito na ang dahilan kung bakit sideline muna ang pwedeng simula.',
    '',
    'Okay din kung gusto mo siyang seryosohin in time.',
    '',
    'Walang pressure, at walang hard selling dito.',
    '',
    'Gusto lang naming makilala ka muna, at makita kung fit ba tayo sa isa\'t isa.',
    '',
    'Kasi hindi lahat ng opportunity ay para sa lahat.',
    '',
    'At okay lang din yun.',
    '',
    'Mas gusto naming makasama yung mga taong totoong interesado, at handang magsipag at magtiyaga.',
    '',
    'Ang importante, may gana kang matuto.',
    '',
    'At may pagkakataon tayong magkakilala nang maayos.',
    '',
    'So kung ikaw yun, masaya ako na nandito ka.',
    '',
    'Masaya ako na nakilala ka, kahit sa pamamagitan lang ng form na ito.',
    '',
    'Salamat sa tiwala, kahit konti pa lang ang alam mo tungkol sa amin.',
    '',
    'Hanggang sa susunod nating pag-uusap.',
    '',
    'Thank you ulit, ' + name + '. Talk soon!',
    '',
    'Habang naghihintay, pwede mong silipin ang agent portal namin. Pwede ka ring mag-inquire doon kung may tanong ka.',
    CONFIG.AGENT_PORTAL_URL || '',
    '',
    'Warm regards,',
    'Leniza Pasion',
    'Tahanan Natin',
    '',
    '---',
    'Please remember that submitting the qualification form does not guarantee recruitment, income, or acceptance into the team. Actual role terms, commission structure, requirements, projects, and incentives will be discussed with the team.'
  ].join('\n');
}


// ============================================================
// AUTOMATIC EMAIL OUTBOX (backup way to send the thank-you email)
// ============================================================
//
// The landing page tries to send the email immediately. If that
// does not work (for example the web app is not allowed to send
// mail), the row stays marked RETRY and this trigger-based sender
// emails the applicant within about a minute.

const EMAIL_STATUS_COL = 19;
const MAX_EMAIL_RETRIES = 5;


function setEmailStatus_(rowNumber, text) {
  try {
    const sheet = getRecruitmentSpreadsheet_()
      .getSheetByName(CONFIG.SHEET_NAME);

    sheet.getRange(rowNumber, EMAIL_STATUS_COL).setValue(text);
  } catch (error) {
    console.error('Could not write email status: ' + error.message);
  }
}


/**
 * Runs every minute (installed by installEmailTrigger).
 * Sends the thank-you email for every row that has not been
 * emailed yet.
 */
function sendPendingApplicantEmails() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(20000)) {
    return;
  }

  try {
    const sheet = getRecruitmentSpreadsheet_()
      .getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    const rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, EMAIL_STATUS_COL)
      .getValues();

    rows.forEach(function(row, index) {

      const rowNumber = index + 2;
      const status = String(row[EMAIL_STATUS_COL - 1] || '').trim();
      const email = clean_(row[3]);
      const fullName = clean_(row[1]);

      if (!email) {
        return;
      }

      let attempts = 0;

      if (status === '') {
        attempts = 0;
      } else if (/^RETRY\s+\d+/.test(status)) {
        attempts = Number(status.match(/^RETRY\s+(\d+)/)[1]);
      } else {
        return; // already Sent / Skipped / Failed
      }

      if (attempts >= MAX_EMAIL_RETRIES) {
        setEmailStatus_(
          rowNumber,
          'FAILED after ' + MAX_EMAIL_RETRIES + ' tries. ' + status
        );
        return;
      }

      try {
        sendApplicantThankYou_({
          fullName: fullName,
          email: email
        });

        setEmailStatus_(rowNumber, 'Sent ' + new Date().toISOString());

      } catch (error) {
        setEmailStatus_(
          rowNumber,
          'RETRY ' + (attempts + 1) + ': ' + error.message
        );
      }
    });

  } finally {
    lock.releaseLock();
  }
}


/**
 * RUN THIS ONCE from the Apps Script editor.
 *
 * - Asks you to approve email permission (this is what is
 *   usually missing when emails do not arrive).
 * - Marks existing rows as "Not sent (before auto-email)" so old
 *   test applicants do not get emailed.
 * - Installs the every-minute sender.
 */
function installEmailTrigger() {

  setupRecruitmentSheet();

  const sheet = getRecruitmentSpreadsheet_()
    .getSheetByName(CONFIG.SHEET_NAME);

  if (sheet.getLastRow() >= 2) {
    const range = sheet.getRange(2, EMAIL_STATUS_COL, sheet.getLastRow() - 1, 1);
    const values = range.getValues().map(function(r) {
      return [String(r[0]).trim() === '' ? 'Not sent (before auto-email)' : r[0]];
    });
    range.setValues(values);
  }

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendPendingApplicantEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sendPendingApplicantEmails')
    .timeBased()
    .everyMinutes(1)
    .create();

  // Forces the email permission prompt to appear now.
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
  Logger.log('Auto-email trigger installed.');
}


// ============================================================
// EMAIL DIAGNOSTIC
// ============================================================

/**
 * Run this from the Apps Script editor when emails are not arriving.
 * Change TEST_TO below to an email you can check, press Run, then
 * read the result in View > Execution log.
 */
function diagnoseEmail() {

  const TEST_TO = 'PUT-A-REAL-EMAIL-HERE@gmail.com';

  Logger.log('Script runs as: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());

  MailApp.sendEmail({
    to: TEST_TO,
    subject: 'Tahanan Natin email test',
    body: 'If you can read this, MailApp is working.',
    name: CONFIG.SENDER_NAME
  });

  Logger.log('Plain test email sent to ' + TEST_TO + '. Check Inbox and Spam.');

  testApplicantEmail_(TEST_TO);
}


function testApplicantEmail_(to) {

  sendApplicantThankYou_({
    timestamp: new Date(),
    fullName: 'Test Applicant',
    email: to,
    status: 'TEST'
  });

  Logger.log('Full applicant email sent to ' + to);
}


// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function getRecruitmentSpreadsheet_() {

  // 1. If a Spreadsheet ID is explicitly configured, use it.
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  // 2. If this project is bound to a Google Sheet, use that Sheet.
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  // 3. If this is a standalone Apps Script project, automatically
  // create the recruitment spreadsheet on first setup.
  const props = PropertiesService.getScriptProperties();
  const savedSpreadsheetId = props.getProperty('RECRUITMENT_SPREADSHEET_ID');

  if (savedSpreadsheetId) {
    try {
      return SpreadsheetApp.openById(savedSpreadsheetId);
    } catch (err) {
      // Saved ID is no longer accessible. Create a fresh spreadsheet below.
      props.deleteProperty('RECRUITMENT_SPREADSHEET_ID');
    }
  }

  const spreadsheet = SpreadsheetApp.create('TAHANAN NATIN × InnerSparc — Recruitment Applicants');

  // Remember the new Spreadsheet ID so future submissions use the same file.
  props.setProperty('RECRUITMENT_SPREADSHEET_ID', spreadsheet.getId());

  return spreadsheet;
}


function clean_(value) {

  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}


function escapeHtml_(value) {

  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


function firstName_(fullName) {

  const name = clean_(fullName);

  if (!name) {
    return 'there';
  }

  return name.split(/\s+/)[0];
}


function formatDateTime_(date) {

  return Utilities.formatDate(
    new Date(date),
    Session.getScriptTimeZone(),
    'MMMM d, yyyy — h:mm a'
  );
}


// ============================================================
// TEST FUNCTIONS
// ============================================================

/**
 * Run this to verify that the Sheet connection works.
 */
function testRecruitmentSheet() {

  const spreadsheet = getRecruitmentSpreadsheet_();

  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sheetName: CONFIG.SHEET_NAME
  };
}


/**
 * Run this to send a TEST applicant email.
 *
 * IMPORTANT:
 * Change TEST_APPLICANT_EMAIL to your own email before running.
 */
function testApplicantEmail() {

  const testApplicant = {
    timestamp: new Date(),

    fullName: 'Test Applicant',
    mobile: '0917 123 4567',
    email: 'YOUR-EMAIL-HERE@example.com',

    area: 'Bacoor',
    nearMolino: 'Yes — malapit sa Molino',
    exactAddress: 'Test Barangay, Bacoor, Cavite',

    age: '25',
    currentSituation: 'Employee',

    commissionAwareness:
      "Yes, I understand and I'm open to commission-based earning",

    workPreference: 'Sideline / part-time',
    hoursPerWeek: '6–10',

    whyJoin:
      'I want to explore real estate as a possible sideline and learn more about project selling.',

    utmSource: 'test',
    utmMedium: 'test',
    utmCampaign: 'test',

    consent: 'Yes',
    status: 'TEST'
  };

  sendApplicantThankYou_(testApplicant);

  return 'Test applicant email sent.';
}


/**
 * Run this to test the ADMIN email.
 *
 * IMPORTANT:
 * This sends the admin notification to CONFIG.ADMIN_EMAIL.
 */
function testAdminEmail() {

  const testApplicant = {
    timestamp: new Date(),

    fullName: 'Test Applicant',
    mobile: '0917 123 4567',
    email: 'test@example.com',

    area: 'Bacoor',
    nearMolino: 'Yes — malapit sa Molino',
    exactAddress: 'Test Barangay, Bacoor, Cavite',

    age: '25',
    currentSituation: 'Employee',

    commissionAwareness:
      "Yes, I understand and I'm open to commission-based earning",

    workPreference: 'Sideline / part-time',
    hoursPerWeek: '6–10',

    whyJoin:
      'I want to explore real estate as a possible sideline and learn more about project selling.',

    utmSource: 'test',
    utmMedium: 'test',
    utmCampaign: 'test',

    consent: 'Yes',
    status: 'TEST'
  };

  sendAdminNotification_(testApplicant);

  return 'Test admin email sent to ' + CONFIG.ADMIN_EMAIL;
}


// Simple health check
function testWebApp() {
  return 'OK — recruitment landing page backend is ready.';
}