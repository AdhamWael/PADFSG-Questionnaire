# PADFSG Website Discovery Questionnaire

Open `index.html` in a browser to test the questionnaire.

This adaptive client-discovery form uses simple, nontechnical choices to reveal the right PADFSG website structure, visitor experience, content, services and delivery approach. D-Foot International is used as a structural benchmark while preserving a distinct Pan-African identity and service model.

## What was updated

- The card between the hero and the questions was removed.
- The questionnaire now begins directly with Question 1.
- Next, Previous, and question-dot navigation scroll to the active question, not to the top of the website.
- The hero button scrolls directly to Question 1.
- Formspree and Google Sheets can run together. Formspree provides email notifications and a submission inbox. Google Sheets stores structured responses and creates Google Docs and PDF reports.
- Indirect questions about trust, priorities, information comfort and visitor expectations reveal the appropriate design direction without asking the client to choose a visual style.
- The submission calculates an internal scope profile and pricing factors from the selected pages, services, languages, content support, approvals and publishing needs.
- Follow-up sections appear only when the client selects the matching website area.

## Connect Formspree

1. Create an account at Formspree and verify your email.
2. In the Formspree dashboard, select `New Form`.
3. Name it `PADFSG Website Questionnaire`.
4. Set the notification email address.
5. Open the form's Integration area and copy the endpoint. It looks like:

   `https://formspree.io/f/abcde123`

6. Open `config.js` and replace the Formspree URL if you created a new form:

   `https://formspree.io/f/mgogyqza`

   with your real endpoint.

7. Upload the files to a web server and send one test response.
8. Open the Formspree dashboard and confirm the response appears in Submissions and arrives by email.

Formspree is useful as a reliable notification and backup inbox. Keep Google Sheets connected as the main structured reporting database.

## Connect Google Sheets and automatic reports

1. Create a Google Sheet named `PADFSG Website Questionnaire Responses`.
2. Open `Extensions > Apps Script`.
3. Paste the contents of `google-apps-script.gs`.
4. Select `Deploy > New deployment > Web app`.
5. Set `Execute as` to `Me`.
6. Set access to `Anyone`.
7. Authorise Sheets, Drive, and Docs access.
8. Copy the deployed web-app URL.
9. Open `config.js` and replace:

   `https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec`

   with the deployed URL.

Each response will create:

- One row in the Responses sheet with one clearly labelled column per answered question
- A `Complete Answer Summary` cell containing a second full copy of every question and answer
- One formatted Google Docs report with the exact questions and readable answers
- One PDF report containing an executive summary, inferred direction, scope analysis and recommended next steps
- Links to both reports inside the sheet
- An email to the Apps Script owner with the PDF attached

The browser creates one canonical readable answer list and sends the same list to all destinations. Formspree receives individually numbered question fields plus a `complete_answers` field. Google Sheets records the expected answer count and rejects the submission if the readable-answer list is incomplete. The PDF is generated from that same verified list.

To send reports to a different address, set `REPORT_RECIPIENT_EMAIL` near the top of `google-apps-script.gs`. Set `SEND_COPY_TO_RESPONDENT` to `true` only if respondents should also receive the analysis.

After updating `google-apps-script.gs`, create a new Apps Script deployment version. Existing web-app URLs continue to work, but they do not use changed code until a new version is deployed.

The receiver also uses a submission ID to prevent accidental duplicate rows and protects the sheet from formula injection.

## Test the connections

Serve the folder over HTTPS (or a local HTTP server), complete the respondent name and email, and submit once. Confirm that:

1. The response appears in Formspree Submissions.
2. A new row appears on the Google Sheet's `Responses` tab.
3. The row contains working Google Document and PDF Report links.

If Google Sheets does not receive the response, redeploy the Apps Script as a **Web app**, confirm access is **Anyone**, and paste the URL ending in `/exec` into `config.js`. After editing the Apps Script later, create a new deployment version so the live `/exec` endpoint receives the changes.

## Deployment note

The questionnaire must be hosted through HTTPS. Do not test final submission by opening the HTML through a local `file://` URL. Use a local web server, GitHub Pages, Netlify, Cloudflare Pages, or your website hosting.


## Formspree integration status

This package uses Vanilla JavaScript with AJAX because the questionnaire has adaptive steps, a custom success screen, and optional Google Sheets submission. The endpoint `https://formspree.io/f/mgogyqza` is already configured in `config.js`. The submission also maps `respEmail` to Formspree's standard `email` field and `respName` to `name`, which improves email reply handling and submission display in the Formspree dashboard.
