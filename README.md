# Lazybird for Google Docs, Sheets and Slides

Read your documents, slides and spreadsheets aloud with natural AI voices, and save them as MP3 files. 500+ voices in 150+ languages and accents, plus the voices you clone in [Lazybird](https://lazybird.app).

## Features

- **Google Docs.** Read the selection, everything from the cursor on, or the whole document. Lists and tables are read too.
- **Google Slides.** Read this slide, every slide from here on, or the whole presentation. Choose the slide text, the speaker notes or both. Text on a slide is read top to bottom, left to right, and skipped slides are left out.
- **Google Sheets.** Read the selected cells or the whole sheet, row by row.
- **Starts in seconds.** Playback starts as soon as the first lines are ready, even for long documents, and keeps going while the rest is generated.
- **Player controls.** Pause, resume, stop and replay. The sidebar shows which slide or row is playing.
- **Save as MP3.** Save the whole reading as one MP3 file, named after your file.
- **500+ voices.** Search by name, style or accent, filter by language and gender, and preview a voice before you pick it. Voices you cloned in Lazybird show up under **My voices**.
- **Speed** from 0.5× to 2×.
- **Remembers your choices.** Your voice, speed and what to read are saved in your Google account.

## Setup

1. Install **Lazybird** from the Google Workspace Marketplace.
2. In Docs, Sheets or Slides, open **Extensions → Lazybird → Read aloud**.
3. Enter your email and select **Email me a code**. Type the 6-digit code from the email and select **Sign in**. If you're new to Lazybird, this creates a free account.

   Or, if you already have an API key, select **Have an API key? Use it instead** and paste it. You can create one in [Lazybird Studio settings](https://lazybird.app/studio/settings).
4. Pick a voice and a speed, choose what to read, and select **Read aloud**.

## Account and payments

- **A Lazybird account is required.** The add-on turns text into speech on Lazybird's servers.
- **The free plan includes 10,000 characters a month**, no card needed. Each character read uses one credit.
- **Paid plans are optional.** If you need more characters, see [lazybird.app/pricing](https://lazybird.app/pricing).

## Permissions and privacy

The add-on uses only the "current file" permissions: it can read the document, presentation or spreadsheet you open it in, and nothing else in your Drive. It reads only the part you choose, only when you select **Read aloud**, and never changes your files.

The text you choose is sent to the Lazybird API at `api.lazybird.app` to generate the audio. Your API key and settings are stored in your Google account. The add-on has no analytics. See [PRIVACY.md](PRIVACY.md) for the details, and Lazybird's [privacy policy](https://lazybird.app/privacy) and [terms](https://lazybird.app/terms).

## Development

The add-on is a standalone [Apps Script](https://developers.google.com/apps-script) project. The code lives in `src/`:

- `Code.js` runs on Apps Script. It adds the menu, reads the text of the open file and keeps the user's settings.
- `Sidebar.html`, `Styles.html` and `Client.html` make up the sidebar. The sidebar calls the Lazybird API from the browser, so the script itself needs no network permission.
- `Help.html` is the help dialog.

To work on it with [clasp](https://github.com/google/clasp):

```bash
npx @google/clasp login
npx @google/clasp create-script --type standalone --title "Lazybird" --rootDir src
git checkout src/appsscript.json   # create-script overwrites the manifest
npx @google/clasp push
```

Then open the script, choose **Deploy → Test deployments**, add an **Editor Add-on** test with **Latest code**, pick a document and select **Execute**.

## Support

Open an issue at https://github.com/lazybirdai/google-workspace-addon-lazybird/issues or write to support@lazybird.app.

## License

[MIT](LICENSE)
