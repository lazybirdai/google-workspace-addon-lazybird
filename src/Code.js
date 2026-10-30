/**
 * Lazybird for Google Docs, Sheets and Slides.
 *
 * The server side only reads the text of the open file and keeps the user's
 * settings. The sidebar calls the Lazybird API itself, so the audio goes
 * straight to the browser.
 */

var SETTINGS_KEY = 'lazybird.settings';
var API_KEY_KEY = 'lazybird.apiKey';

function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  var ui = hostUi_();
  if (!ui) return;
  ui.createAddonMenu()
    .addItem('Read aloud', 'showSidebar')
    .addItem('Help', 'showHelp')
    .addToUi();
}

function showSidebar() {
  var template = HtmlService.createTemplateFromFile('Sidebar');
  template.host = host_();
  var html = template.evaluate().setTitle('Lazybird');
  hostUi_().showSidebar(html);
}

function showHelp() {
  var html = HtmlService.createHtmlOutputFromFile('Help').setWidth(420).setHeight(440);
  hostUi_().showModalDialog(html, 'Lazybird help');
}

/** Lets the sidebar template pull in the style and script files. */
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ---------------------------------------------------------------------------
// Settings

function getState() {
  var props = PropertiesService.getUserProperties();
  var settings = {};
  try {
    settings = JSON.parse(props.getProperty(SETTINGS_KEY) || '{}');
  } catch (err) {
    settings = {};
  }
  return { host: host_(), apiKey: props.getProperty(API_KEY_KEY) || '', settings: settings };
}

function saveApiKey(apiKey) {
  var props = PropertiesService.getUserProperties();
  if (apiKey) props.setProperty(API_KEY_KEY, String(apiKey).trim());
  else props.deleteProperty(API_KEY_KEY);
}

function saveSettings(settings) {
  PropertiesService.getUserProperties().setProperty(SETTINGS_KEY, JSON.stringify(settings || {}));
}

// ---------------------------------------------------------------------------
// Reading the open file

/**
 * Returns the file's title and the text to read as a list of parts, each
 * with an optional label shown while it plays (for example "Slide 3").
 *
 * @param {Object} options
 *   Docs:   { scope: 'selection' | 'cursor' | 'document' }
 *   Slides: { scope: 'slide' | 'from-slide' | 'presentation', source: 'text' | 'notes' | 'both' }
 *   Sheets: { scope: 'selection' | 'sheet' }
 */
function getText(options) {
  options = options || {};
  var host = host_();
  var parts;
  var title;
  if (host === 'docs') {
    parts = docsText_(options.scope);
    title = DocumentApp.getActiveDocument().getName();
  } else if (host === 'slides') {
    parts = slidesText_(options.scope, options.source);
    title = SlidesApp.getActivePresentation().getName();
  } else if (host === 'sheets') {
    parts = sheetsText_(options.scope);
    title = SpreadsheetApp.getActiveSpreadsheet().getName();
  } else {
    throw new Error('Open Lazybird from Google Docs, Sheets or Slides.');
  }
  parts = parts.filter(function (part) {
    return /\S/.test(part.text);
  });
  return { title: title, parts: parts };
}

// Docs

function docsText_(scope) {
  var doc = DocumentApp.getActiveDocument();
  if (scope === 'selection') {
    var selection = doc.getSelection();
    if (!selection) throw new Error('Select some text first.');
    var lines = selection.getRangeElements().map(function (range) {
      var el = range.getElement();
      if (range.isPartial()) {
        return el.asText().getText().substring(range.getStartOffset(), range.getEndOffsetInclusive() + 1);
      }
      return elementText_(el);
    });
    return [{ label: '', text: lines.join('\n') }];
  }
  var body = doc.getBody();
  if (scope === 'cursor') {
    var cursor = doc.getCursor();
    if (!cursor) throw new Error('Click in the document where reading should start.');
    var block = cursor.getElement();
    while (block.getParent() && block.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
      block = block.getParent();
    }
    if (!block.getParent()) throw new Error('Click in the main text of the document.');
    var first = cursor.getSurroundingText().getText().substring(cursor.getSurroundingTextOffset());
    var rest = [];
    for (var i = body.getChildIndex(block) + 1; i < body.getNumChildren(); i++) {
      rest.push(elementText_(body.getChild(i)));
    }
    return [{ label: '', text: [first].concat(rest).join('\n') }];
  }
  return [{ label: '', text: body.getText() }];
}

function elementText_(el) {
  var Type = DocumentApp.ElementType;
  switch (el.getType()) {
    case Type.PARAGRAPH:
      return el.asParagraph().getText();
    case Type.LIST_ITEM:
      return el.asListItem().getText();
    case Type.TABLE:
      return el.asTable().getText();
    case Type.TABLE_ROW:
      return el.asTableRow().getText();
    case Type.TABLE_CELL:
      return el.asTableCell().getText();
    case Type.TEXT:
      return el.asText().getText();
    default:
      return '';
  }
}

// Slides

function slidesText_(scope, source) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var current = presentation.getSelection().getCurrentPage();
  var currentIndex = -1;
  if (current) {
    var id = current.getObjectId();
    for (var i = 0; i < slides.length; i++) {
      if (slides[i].getObjectId() === id) currentIndex = i;
    }
  }
  var picked;
  if (scope === 'presentation') {
    picked = slides.map(function (slide, index) {
      return { slide: slide, index: index };
    });
  } else {
    if (currentIndex < 0) throw new Error('Select a slide first.');
    var end = scope === 'from-slide' ? slides.length : currentIndex + 1;
    picked = [];
    for (var j = currentIndex; j < end; j++) picked.push({ slide: slides[j], index: j });
  }
  // Skipped slides don't show in the slideshow, so they aren't read either, unless picked on their own.
  if (picked.length > 1) {
    picked = picked.filter(function (p) {
      return !p.slide.isSkipped();
    });
  }
  return picked.map(function (p) {
    var texts = [];
    if (source !== 'notes') texts.push(slideText_(p.slide));
    if (source !== 'text') texts.push(notesText_(p.slide));
    return { label: 'Slide ' + (p.index + 1), text: texts.filter(Boolean).join('\n'), slideId: p.slide.getObjectId() };
  });
}

/** Text on a slide in reading order: top to bottom, then left to right. */
function slideText_(slide) {
  var items = [];
  collectText_(slide.getPageElements(), items);
  items.sort(function (a, b) {
    return Math.abs(a.top - b.top) > 10 ? a.top - b.top : a.left - b.left;
  });
  return items
    .map(function (item) {
      return item.text;
    })
    .join('\n');
}

function collectText_(elements, items) {
  var Type = SlidesApp.PageElementType;
  elements.forEach(function (el) {
    var type = el.getPageElementType();
    var text = '';
    if (type === Type.SHAPE) {
      text = el.asShape().getText().asString();
    } else if (type === Type.TABLE) {
      var table = el.asTable();
      var rows = [];
      for (var r = 0; r < table.getNumRows(); r++) {
        var cells = [];
        for (var c = 0; c < table.getNumColumns(); c++) {
          var cell = table.getCell(r, c);
          if (cell.getMergeState() !== SlidesApp.CellMergeState.MERGED) cells.push(cell.getText().asString().trim());
        }
        rows.push(cells.filter(Boolean).join(', '));
      }
      text = rows.filter(Boolean).join('\n');
    } else if (type === Type.GROUP) {
      collectText_(el.asGroup().getChildren(), items);
      return;
    }
    text = text.trim();
    if (text) items.push({ top: el.getTop(), left: el.getLeft(), text: text });
  });
}

function notesText_(slide) {
  var shape = slide.getNotesPage().getSpeakerNotesShape();
  return shape ? shape.getText().asString().trim() : '';
}

// Sheets

function sheetsText_(scope) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = scope === 'sheet' ? sheet.getDataRange() : sheet.getActiveRange();
  if (!range) throw new Error('Select some cells first.');
  var firstRow = range.getRow();
  // One part per row: the sidebar shows which row is playing, and rows of
  // look-alike values stay in separate requests, where the voice can't skip
  // back to an earlier row.
  return range.getDisplayValues().map(function (row, i) {
    var text = row
      .map(function (value) {
        return String(value).trim();
      })
      .filter(Boolean)
      .join(', ');
    return { label: 'Row ' + (firstRow + i), text: text };
  });
}

// ---------------------------------------------------------------------------
// Host detection

function host_() {
  try {
    DocumentApp.getUi();
    return 'docs';
  } catch (err) {}
  try {
    SlidesApp.getUi();
    return 'slides';
  } catch (err) {}
  try {
    SpreadsheetApp.getUi();
    return 'sheets';
  } catch (err) {}
  return '';
}

function hostUi_() {
  switch (host_()) {
    case 'docs':
      return DocumentApp.getUi();
    case 'slides':
      return SlidesApp.getUi();
    case 'sheets':
      return SpreadsheetApp.getUi();
    default:
      return null;
  }
}

// ---------------------------------------------------------------------------
// Slide voiceovers
//
// Each voiceover is a small video clip saved in the user's Drive and placed on
// its slide, where it starts by itself in the slideshow. Slides only plays
// videos, so the sidebar wraps the reading in one. With the drive.file
// scope Lazybird can only open files it created and decks the user picked for
// it in Google's file picker, so the picker shows once per presentation.

var GCP_PROJECT_NUMBER = '1008306091565';
var VOICEOVER_PREFIX = 'lazybird_vo_';
var VOICEOVER_FOLDER = 'Lazybird voiceovers';
var VOICEOVER_FOLDER_KEY = 'lazybird.voiceoverFolder';
var VOICEOVER_SIZE = 40; // points
var PICKER_RESULT_KEY = 'lazybird.picker';

/** Whether Lazybird may edit the open presentation through the Slides API. */
function hasDeckAccess() {
  try {
    Slides.Presentations.get(SlidesApp.getActivePresentation().getId(), { fields: 'presentationId' });
    return true;
  } catch (err) {
    if (/not found|permission/i.test(err.message)) return false;
    throw err;
  }
}

/** Opens Google's file picker, showing only the open presentation. */
function showDeckPicker() {
  CacheService.getUserCache().remove(PICKER_RESULT_KEY);
  var html = HtmlService.createHtmlOutputFromFile('Picker').setWidth(640).setHeight(440);
  SlidesApp.getUi().showModalDialog(html, 'Select this presentation');
}

function getPickerConfig() {
  return { token: ScriptApp.getOAuthToken(), fileId: SlidesApp.getActivePresentation().getId(), appId: GCP_PROJECT_NUMBER };
}

/** The picker reports "picked" or "cancel" here; the sidebar waits for it. */
function setPickerResult(result) {
  CacheService.getUserCache().put(PICKER_RESULT_KEY, result === 'picked' ? 'picked' : 'cancel', 600);
}

/** "ok" once Lazybird may edit the deck, "cancel" if the picker was closed, else "waiting". */
function deckAccessStatus() {
  if (hasDeckAccess()) return 'ok';
  return CacheService.getUserCache().get(PICKER_RESULT_KEY) === 'cancel' ? 'cancel' : 'waiting';
}

/**
 * Saves the clip to Drive and puts it on the slide in place of the slide's
 * earlier Lazybird voiceover. The old file stays in Drive, since a copy of the
 * deck may still use it.
 *
 * @param {string} slideId a slide of the open presentation
 * @param {string} data the WebM clip, base64-encoded
 */
function addVoiceover(slideId, data) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var index = -1;
  for (var i = 0; i < slides.length; i++) {
    if (slides[i].getObjectId() === slideId) index = i;
  }
  if (index < 0) throw new Error('That slide is no longer in the presentation.');
  var slide = slides[index];
  var name = presentation.getName() + ' – Slide ' + (index + 1) + '.webm';
  var blob = Utilities.newBlob(Utilities.base64Decode(data), 'video/webm', name);
  var file;
  try {
    file = Drive.Files.create({ name: name, mimeType: 'video/webm', parents: [voiceoverFolder_()] }, blob);
  } catch (err) {
    if (/quota/i.test(err.message)) throw new Error('Your Google Drive is full, so the voiceover could not be saved.');
    throw err;
  }

  var presentationId = presentation.getId();
  var page = Slides.Presentations.Pages.get(presentationId, slide.getObjectId(), { fields: 'pageElements(objectId)' });
  var requests = (page.pageElements || [])
    .filter(function (el) {
      return el.objectId.indexOf(VOICEOVER_PREFIX) === 0;
    })
    .map(function (el) {
      return { deleteObject: { objectId: el.objectId } };
    });
  var objectId = VOICEOVER_PREFIX + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  var margin = 12;
  requests.push(
    {
      createVideo: {
        objectId: objectId,
        source: 'DRIVE',
        id: file.id,
        elementProperties: {
          pageObjectId: slide.getObjectId(),
          size: { width: { magnitude: VOICEOVER_SIZE, unit: 'PT' }, height: { magnitude: VOICEOVER_SIZE, unit: 'PT' } },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: presentation.getPageWidth() - VOICEOVER_SIZE - margin,
            translateY: presentation.getPageHeight() - VOICEOVER_SIZE - margin,
            unit: 'PT',
          },
        },
      },
    },
    { updateVideoProperties: { objectId: objectId, videoProperties: { autoPlay: true }, fields: 'autoPlay' } },
  );
  try {
    Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
  } catch (err) {
    Drive.Files.update({ trashed: true }, file.id);
    throw err;
  }
}

/** The "Lazybird voiceovers" folder in the user's Drive, created on first use. */
function voiceoverFolder_() {
  var props = PropertiesService.getUserProperties();
  var id = props.getProperty(VOICEOVER_FOLDER_KEY);
  if (id) {
    try {
      if (!Drive.Files.get(id, { fields: 'trashed' }).trashed) return id;
    } catch (err) {}
  }
  id = Drive.Files.create({ name: VOICEOVER_FOLDER, mimeType: 'application/vnd.google-apps.folder' }).id;
  props.setProperty(VOICEOVER_FOLDER_KEY, id);
  return id;
}
