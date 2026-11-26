# Invite Mail Sender (Chrome extension)

Sends one personal email to each person in a CSV/Excel list from your own Gmail account. The email goes out in the background through Google's official Gmail API: no Gmail tab opens, nothing gets clicked, and your inbox tab isn't touched. Emails show up in your Gmail **Sent** folder like normal ones.

It's free. The Gmail API has no paid tier, and you never add a card, so nothing can ever bill you.

## 1. Install the extension

1. Open `chrome://extensions`.
2. Turn on **Developer mode** (top right).
3. Click **Load unpacked** and pick this folder (the one with `manifest.json`).
   If an older version is loaded, **Remove** it first, then load this folder again (the extension ID changed).
4. Check that the ID on the card is `dlibgmnmglameeebdkhpelifjmkaimml`.
5. Pin it (puzzle icon), then click it. The sender opens in the right-hand side panel.

## 2. Gmail API setup (one time, about 10 minutes)

Google needs to know this extension is allowed to send mail as you. You set that up once in Google Cloud. No billing account or card is needed.

1. Go to <https://console.cloud.google.com/> and sign in with the Gmail account you'll send from.
   Click the project picker at the top → **New project** → name it e.g. `Invite Sender` → **Create**, and make sure it's selected.
2. **Turn on the Gmail API**: open <https://console.cloud.google.com/apis/library/gmail.googleapis.com> and click **Enable**.
3. **Consent screen**: open <https://console.cloud.google.com/auth/overview> and click **Get started**.
   - App name: `Invite Sender`. Support email: your Gmail. **Next**.
   - Audience: **External**. **Next**.
   - Contact email: your Gmail. **Next**, accept the policy, **Create**.
4. **Add yourself as a test user**: go to **Audience** (left menu) → under *Test users*, **Add users** → your Gmail address → **Save**.
   Skip this and Google will block the sign-in with "access denied".
5. **Create the Client ID**: go to **Clients** (left menu) → **Create client**.
   - Application type: **Web application**. Name: anything.
   - Under **Authorized redirect URIs** click **Add URI** and paste:
     `https://dlibgmnmglameeebdkhpelifjmkaimml.chromiumapp.org/`
     (the side panel also shows it, with a Copy button).
   - **Create**, then copy the **Client ID** (it ends with `.apps.googleusercontent.com`).
6. In the extension's side panel, paste the Client ID → **Save** → **Connect Gmail account**.
   Pick your account. Google will warn *"Google hasn't verified this app"*. That's expected, because it's your own app: click **Continue**, then allow **Send email on your behalf**.

The top of the panel now says **Sending from you@gmail.com**. You only do this once. Later, if Google asks you to sign in again, it's a single click.

### What the pieces are

- **Gmail API**: the official "send an email" service. Turning it on in the project is what lets programs send mail.
- **OAuth**: the permission step. It's the Google popup where you say "yes, this extension may send email as me". The **Client ID** is the name tag that tells Google which app is asking.
- **Test users**: the project is in "testing" mode, so only Gmail addresses on this list can say yes. The limit is 100 addresses.

### Using your other Gmail accounts

You don't need a new project or Client ID. The same one works for all your Gmails:

1. Open <https://console.cloud.google.com/auth/audience>, signed in as the account that owns the project, with the same project selected.
2. Under **Test users**, click **Add users**, add the other Gmail address, and click **Save**.
3. In the side panel, click **Switch account**, pick that Gmail, and click **Continue** on the "hasn't verified" warning.

The panel now says **Sending from** that address. Switch back the same way.

### On another computer or Chrome profile

Load the same folder (the ID stays `dlibgmnmglameeebdkhpelifjmkaimml`), paste the same Client ID, and connect. You can find the Client ID again under **Clients** at <https://console.cloud.google.com/auth/clients>.

## 3. Use

1. **Recipients**: choose your `.csv` / `.xlsx` file. It needs an email column (header like `email`, `e-mail`, `邮箱`). A name column (`name`, `姓名`…) is used for `{{name}}`. If there's no header row, the columns are detected automatically. Invalid and duplicate addresses are flagged and skipped. See `examples/recipients-example.csv`.
2. **Templates**: write 5–10 subject/body pairs. Use `{{name}}`, `{{email}}`, or any column header from your file (e.g. `{{company}}`). One template is chosen at random per person. Templates are saved automatically, and you can export/import them as JSON (`examples/templates-example.json`).
3. **Settings**:
   - *Min / max delay*: the random wait between emails, in seconds (e.g. 2–6).
   - *Dry run*: goes through the list and the delays without sending anything.
   - *Send test*: type the address the test should go to (your own is fine) and click it. A random template is sent there.
4. **Send**: press **Start sending**. You'll see a progress bar, a countdown to the next email, a status per person and a log. **Pause** and **Stop** take effect between emails.

Keep the side panel open while it sends. You can use Gmail and other tabs normally.

## Safety notes

- Progress is saved. If you close the side panel or Chrome crashes, rows already sent stay marked **Sent**, and pressing Start again only sends to the **Pending** ones.
- If an email was interrupted mid-send, it's marked **Check Sent folder**. Look in Gmail's Sent folder, then click **Queue** or **Skip** for that row.
- Failed rows aren't retried automatically. Use **Queue failed again**.
- After 3 failures in a row the run stops and the log says why.
- Use **Skip** on a row to leave someone out.
- Gmail's own limit is about 500 emails/day for a regular account, so 50–100 is fine.

## Limitations

- Emails are plain text.
- While the Google project is in "testing" mode, only the test users you added (yourself) can connect. That's all you need.

## Files

| File | What it is |
| --- | --- |
| `manifest.json` | Extension manifest (Manifest V3). `key` pins the extension ID so the redirect URI never changes |
| `background.js` | Makes the toolbar icon open the side panel |
| `app.html` / `app.css` / `app.js` | The side panel: file upload, templates, settings, Google sign-in, sending loop |
| `timer-worker.js` | Accurate timers for the random delays |
| `lib/xlsx.full.min.js` | [SheetJS](https://sheetjs.com) 0.20.3, for reading Excel/CSV files |
