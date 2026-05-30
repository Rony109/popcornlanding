function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) throw new Error("Empty request body");

    const data  = JSON.parse(raw);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    sheet.appendRow([
      new Date().toLocaleString("en-CA", { timeZone: "America/Toronto" }),
      data.name  || "",
      data.city  || "",
      data.email || "",
    ]);

    if (data.email) {
      sendConfirmationEmail(data.name, data.email);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log("doPost error: " + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendConfirmationEmail(name, email) {
  const firstName = (name && name.trim()) ? name.trim().split(" ")[0] : "there";
  const subject   = "You are on the list, " + firstName + " \uD83C\uDF7F";
  const unsubUrl  = "https://popcornuniverse.ca/unsubscribe?email=" + encodeURIComponent(email);

  const html = getEmailTemplate(firstName, unsubUrl);

  MailApp.sendEmail(email, subject, " ", {
    htmlBody: html,
    from: "marketing.popcornuniverse@gmail.com",
    name: "Popcorn Universe"
  });
}

function getEmailTemplate(firstName, unsubUrl) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>You're in, ${firstName}. Welcome to Popcorn Universe. 🍿</title>
  <style>
    /* Base reset */
    body { margin:0; padding:0; background-color:#0C0A08 !important; font-family:'DM Sans',-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif; }
    a { text-decoration:none; }
 
    /* Force dark bg on the Gmail iOS wrapper divs it injects */
    u + #body-wrap { background-color:#0C0A08 !important; }
    #body-wrap { background-color:#0C0A08 !important; }
 
    /* iOS Gmail targets: override any white injection */
    div[style*="background"] { background-color:#0C0A08 !important; }
 
    /* Prevent iOS from auto-styling phone numbers / dates blue */
    a[x-apple-data-detectors] {
      color: inherit !important;
      text-decoration: none !important;
    }
 
    @media only screen and (max-width:600px) {
      .container { width:100% !important; }
      .h1 { font-size:28px !important; }
      .pad { padding:24px 20px !important; }
      .btn { width:100% !important; display:block !important; }
    }
 
    /* Dark mode media query — reinforces colours if OS is in dark mode */
    @media (prefers-color-scheme: dark) {
      body { background-color:#0C0A08 !important; }
      .email-bg { background-color:#0C0A08 !important; }
    }
  </style>
</head>
<!--
  The id="body-wrap" + u + #body-wrap selector is the standard hack
  to target Gmail iOS's injected wrapper and force our background.
-->
<body id="body-wrap" style="background:#0C0A08 !important; background-color:#0C0A08 !important; margin:0; padding:0;">
 
  <!-- Preheader -->
  <div style="display:none;max-height:0;overflow:hidden;font-size:1px;color:#0C0A08;">
    Your seat is reserved. Here's what comes next. 🍿&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;
  </div>
 
  <!--
    Outer div wrapper: Gmail iOS strips body styles but respects
    inline styles on a direct child div. This is the most reliable fix.
  -->
  <div class="email-bg" style="background-color:#0C0A08 !important; margin:0; padding:0;">
 
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#0C0A08" style="background-color:#0C0A08 !important;">
    <tr>
      <td align="center" style="padding:20px 12px; background-color:#0C0A08 !important;" bgcolor="#0C0A08">
        <table class="container" role="presentation" cellpadding="0" cellspacing="0" border="0" width="580" style="max-width:580px;">
 
          <!-- HEADER -->
          <tr>
            <td align="center" bgcolor="#0C0A08" style="padding:22px 28px 18px; border-radius:10px 10px 0 0;">
              <img src="https://popcornuniverse.ca/Secondary%20Logo/Logo-H_Light.png" alt="Popcorn Universe" height="40" style="display:block;border:0;" />
            </td>
          </tr>
 
          <!-- HERO -->
          <tr>
            <td class="pad" align="center" bgcolor="#131109" style="padding:36px 36px 30px; background:linear-gradient(180deg,rgba(245,200,66,0.08) 0%,#131109 60%);">
              <h1 class="h1" style="font-size:36px;font-weight:700;color:#F5F0E8;line-height:1.15;letter-spacing:-0.5px;margin:0 0 12px;">
                You're in, ${firstName}. 🍿
              </h1>
              <p style="font-size:15px;color:#7A7260;line-height:1.6;margin:0;max-width:420px;">
                Welcome to the Popcorn Universe waitlist — the best decision for your next night out.
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:24px auto 0;">
                <tr>
                  <td style="height:3px;width:64px;background:#F5C842;border-radius:2px;font-size:0;">&nbsp;</td>
                </tr>
              </table>
            </td>
          </tr>
 
          <!-- FOUNDERS NOTE -->
          <tr>
            <td class="pad" bgcolor="#1A1710" style="padding:28px 36px;">
              <p style="font-size:14px;color:#F5F0E8;line-height:1.75;margin:0 0 12px;">
                We're two builders who got tired of juggling five apps just to book a movie. So we built the one app Canada's entertainment lovers actually deserve. You joining this waitlist means the world to us.
              </p>
              <p style="font-size:13px;font-weight:600;color:#C47A2B;margin:0;">— Nishant &amp; Rohan, Co-Founders</p>
            </td>
          </tr>
 
          <!-- DIVIDER -->
          <tr>
            <td bgcolor="#1A1710" style="padding:0 36px;"><div style="height:1px;background:#2C2820;font-size:0;line-height:0;">&nbsp;</div></td>
          </tr>
 
          <!-- WHAT'S NEXT -->
          <tr>
            <td class="pad" bgcolor="#1A1710" style="padding:28px 36px;">
              <p style="font-size:10px;font-weight:700;color:#F5C842;letter-spacing:2px;text-transform:uppercase;margin:0 0 8px;">What's Next</p>
              <h2 style="font-size:22px;font-weight:700;color:#F5F0E8;line-height:1.25;margin:0 0 10px;">Beta is coming. You'll be first.</h2>
              <p style="font-size:14px;color:#F5F0E8;line-height:1.7;margin:0 0 22px;">
                Waitlist members get beta access before anyone else — plus exclusive early-member perks. Keep an eye on your inbox.
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" class="btn">
                <tr>
                  <td align="center" bgcolor="#F5C842" style="border-radius:5px;">
                    <a href="https://popcornuniverse.ca" style="display:inline-block;font-size:14px;font-weight:700;color:#0C0A08;padding:12px 24px;letter-spacing:0.2px;">Stay Tuned &rarr;</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
 
          <!-- DIVIDER -->
          <tr>
            <td bgcolor="#1A1710" style="padding:0 36px;"><div style="height:1px;background:#2C2820;font-size:0;line-height:0;">&nbsp;</div></td>
          </tr>
 
          <!-- FEEDBACK -->
          <tr>
            <td class="pad" bgcolor="#1A1710" style="padding:28px 36px;">
              <p style="font-size:10px;font-weight:700;color:#F5C842;letter-spacing:2px;text-transform:uppercase;margin:0 0 8px;">Your Voice Matters</p>
              <h2 style="font-size:22px;font-weight:700;color:#F5F0E8;line-height:1.25;margin:0 0 10px;">Tell us what you want to see.</h2>
              <p style="font-size:14px;color:#F5F0E8;line-height:1.7;margin:0 0 22px;">What would make Popcorn Universe a must-have for you? We read every reply.</p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" class="btn">
                <tr>
                  <td align="center" bgcolor="#1A1710" style="border:2px solid #F5C842;border-radius:5px;">
                    <a href="mailto:feedback@popcornuniverse.ca" style="display:inline-block;font-size:14px;font-weight:700;color:#F5C842;padding:10px 24px;letter-spacing:0.2px;">Share Your Feedback &rarr;</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
 
          <!-- SOCIAL -->
          <tr>
            <td class="pad" align="center" bgcolor="#131109" style="padding:28px 36px 32px;">
              <p style="font-size:10px;font-weight:700;color:#F5C842;letter-spacing:2px;text-transform:uppercase;margin:0 0 6px;">Join the Community</p>
              <p style="font-size:20px;font-weight:700;color:#F5F0E8;margin:0 0 6px;">Follow along as we build.</p>
              <p style="font-size:13px;color:#7A7260;margin:0 0 22px;">Updates, behind-the-scenes, and sneak peeks.</p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <!-- Instagram -->
                  <td align="center" valign="top" style="padding:0 5px;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td align="center" bgcolor="#201D14" style="border:1px solid #F5C842;border-radius:7px;">
                          <a href="#" style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#F5C842;text-decoration:none;display:block;padding:10px 10px;">IG</a>
                        </td>
                      </tr>
                      <tr><td align="center" style="padding-top:5px;font-family:Arial,sans-serif;font-size:9px;color:#7A7260;">Instagram</td></tr>
                    </table>
                  </td>
                  <!-- LinkedIn -->
                  <td align="center" valign="top" style="padding:0 5px;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td align="center" bgcolor="#201D14" style="border:1px solid #F5C842;border-radius:7px;">
                          <a href="#" style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#F5C842;text-decoration:none;display:block;padding:10px 10px;">in</a>
                        </td>
                      </tr>
                      <tr><td align="center" style="padding-top:5px;font-family:Arial,sans-serif;font-size:9px;color:#7A7260;">LinkedIn</td></tr>
                    </table>
                  </td>
                  <!-- X -->
                  <td align="center" valign="top" style="padding:0 5px;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td align="center" bgcolor="#201D14" style="border:1px solid #F5C842;border-radius:7px;">
                          <a href="#" style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#F5C842;text-decoration:none;display:block;padding:10px 13px;">X</a>
                        </td>
                      </tr>
                      <tr><td align="center" style="padding-top:5px;font-family:Arial,sans-serif;font-size:9px;color:#7A7260;">X / Twitter</td></tr>
                    </table>
                  </td>
                  <!-- TikTok -->
                  <td align="center" valign="top" style="padding:0 5px;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td align="center" bgcolor="#201D14" style="border:1px solid #F5C842;border-radius:7px;">
                          <a href="#" style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#F5C842;text-decoration:none;display:block;padding:10px 9px;">TT</a>
                        </td>
                      </tr>
                      <tr><td align="center" style="padding-top:5px;font-family:Arial,sans-serif;font-size:9px;color:#7A7260;">TikTok</td></tr>
                    </table>
                  </td>
                  <!-- Snapchat -->
                  <td align="center" valign="top" style="padding:0 5px;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td align="center" bgcolor="#201D14" style="border:1px solid #F5C842;border-radius:7px;">
                          <a href="#" style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#F5C842;text-decoration:none;display:block;padding:10px 8px;">SC</a>
                        </td>
                      </tr>
                      <tr><td align="center" style="padding-top:5px;font-family:Arial,sans-serif;font-size:9px;color:#7A7260;">Snapchat</td></tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
 
          <!-- FOOTER -->
          <tr>
            <td align="center" bgcolor="#0C0A08" style="padding:24px 36px 32px;border-radius:0 0 10px 10px;border-top:1px solid #1E1C14;">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td align="center" style="padding-bottom:14px;">
                    <img src="https://popcornuniverse.ca/LogoMark/LogoMark-Light.png" alt="Popcorn Universe" width="28" height="28" style="display:inline-block;border:0;opacity:0.6;" />
                  </td>
                </tr>
                <tr>
                  <td align="center" style="padding-bottom:6px;">
                    <p style="font-size:11px;color:#7A7260;margin:0;">&copy; 2026 Popcorn Universe Inc. &middot; Waterloo, Ontario, Canada</p>
                  </td>
                </tr>
                <tr>
                  <td align="center" style="padding-bottom:12px;">
                    <p style="font-size:11px;color:#7A7260;margin:0;">You're receiving this because you joined the Popcorn Universe waitlist.</p>
                  </td>
                </tr>
                <tr>
                  <td align="center" style="padding-bottom:12px;">
                    <div style="height:1px;background:#1E1C14;width:40%;margin:0 auto;font-size:0;">&nbsp;</div>
                  </td>
                </tr>
                <tr>
                  <td align="center" style="padding-bottom:16px;">
                    <p style="font-size:11px;color:#7A7260;margin:0;line-height:1.7;">
                      Don't want to hear from us?&nbsp;<a href="${unsubUrl}" style="color:#C47A2B;text-decoration:underline;font-weight:500;">Unsubscribe</a>
                      &nbsp;&middot;&nbsp;
                      <a href="https://popcornuniverse.ca/privacy" style="color:#7A7260;text-decoration:underline;">Privacy Policy</a>
                    </p>
                  </td>
                </tr>
                <tr>
                  <td align="center">
                    <p style="font-size:10px;color:#3A3628;letter-spacing:1.5px;text-transform:uppercase;margin:0;">Entertainment Unleashed</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
 
        </table>
      </td>
    </tr>
  </table>
 
  </div><!-- /email-bg -->
 
</body>
</html>`;
}

// Run this once manually to add headers
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow(["Timestamp", "Name", "City", "Email"]);
  sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#F5C842").setFontColor("#0C0A08");
  sheet.setFrozenRows(1);
}