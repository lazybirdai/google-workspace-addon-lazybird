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
    return { label: 'Slide ' + (p.index + 1), text: texts.filter(Boolean).join('\n') };
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
  var rows = range.getDisplayValues().map(function (row) {
    return row
      .map(function (value) {
        return String(value).trim();
      })
      .filter(Boolean)
      .join(', ');
  });
  return [{ label: '', text: rows.filter(Boolean).join('\n') }];
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
