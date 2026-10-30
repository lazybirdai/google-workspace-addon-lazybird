# Privacy policy

Last updated: October 30, 2026

This policy covers **Lazybird for Google Docs, Sheets and Slides** (the "add-on"), made by LAZYBIRD COMPANY LIMITED. It explains which Google user data the add-on reads, what it does with it, and where it goes.

## What the add-on can access

The add-on asks Google for these permissions:

| Permission | Why |
| --- | --- |
| View and manage documents that this application has been installed in (`documents.currentonly`) | To read the text of the document you open the add-on in. |
| View and manage Google Slides presentations that this application has been installed in (`presentations.currentonly`) | To read slide text and speaker notes in the presentation you open the add-on in. |
| View and manage spreadsheets that this application has been installed in (`spreadsheets.currentonly`) | To read cell values in the spreadsheet you open the add-on in. |
| See, edit, create and delete only the specific Google Drive files you use with this app (`drive.file`) | For slide voiceovers: to save the voiceover clips in your Drive and add them to the presentation you select in Google's file picker. |
| Display and run third-party web content in prompts and sidebars inside Google applications (`script.container.ui`) | To show the Lazybird sidebar, the file picker and the help dialog. |

The add-on only reads the file you open it in, and only when you press **Read aloud** or **Add voiceover**. It reads the part you chose: the selection, the text from the cursor on, the whole document, one or more slides, the selected cells or the whole sheet. Reading never changes your files, and the add-on can't see your other files, your email or your contacts.

**Add voiceover** is the only feature that changes anything. It saves one video clip per slide in a **Lazybird voiceovers** folder in your Drive and adds it to that slide, in place of the slide's earlier Lazybird voiceover. With `drive.file`, the add-on can open only the files it created and the presentations you selected for it.

## What the add-on sends, and to whom

The add-on talks to one service: the Lazybird API at `api.lazybird.app`, over HTTPS.

- **The text you chose to hear**, with the voice and speed you picked, to generate the audio.
- **Your email address and the 6-digit code**, when you sign in with email. Signing in creates a free Lazybird account if you don't have one, and returns an API key for the add-on.
- **Your Lazybird API key**, to sign each request.
- **Voice searches and filters**, to list voices in the voice picker.

The add-on does not send Google user data anywhere else. It has no analytics, no ads and no tracking.

## What the add-on stores

- **Your Lazybird API key and your settings** (voice, speed and what to read) are stored in your Google account with Apps Script's user properties. Only you and the add-on can read them. **Disconnect** in the add-on's settings deletes the key.
- **Audio** stays in your browser's memory while it plays and is discarded when you stop or close the sidebar. It is saved to your computer only when you press **MP3**.
- **Voiceover clips** are saved in your own Google Drive when you press **Add voiceover**. They stay there until you delete them. Lazybird keeps no copy.

## How Lazybird handles the data

Lazybird processes the text to generate speech and keeps usage records for billing, as described in the [Lazybird privacy policy](https://lazybird.app/privacy). Lazybird does not sell your data and does not use your text to train AI models.

Lazybird's use and transfer of information received from Google APIs adheres to the [Google API Services User Data Policy](https://developers.google.com/terms/api-services-user-data-policy), including the Limited Use requirements. Google user data is used only to provide the add-on's reading features, is not transferred to anyone except as needed to provide those features, is not used for advertising, and is not read by people unless you ask us to for support or the law requires it.

## Removing the add-on

You can remove the add-on at any time from **Extensions → Add-ons → Manage add-ons** in Docs, Sheets or Slides, and revoke its access at [myaccount.google.com/permissions](https://myaccount.google.com/permissions). To delete your Lazybird account and its data, use **Settings → Account → Delete account** in [Lazybird Studio](https://lazybird.app/studio/settings).

## Contact

Questions about privacy: support@lazybird.app
