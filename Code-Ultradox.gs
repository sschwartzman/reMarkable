// Global variable to track number of document updates
// To return to Ultradox in results (to verify it's working)
var numChanged = 0;

function formatMarkdownOfDoc(doc) {
  processSourceCode(doc);
  var docbody = doc.getBody();
  processBackquotes(docbody);
  processBold(docbody);
  processLinks(docbody);
  processItalics(docbody);
  processHeadings(docbody);
  processLists(docbody);
}

/*
* Utility function to remove paragraphs.
* Deals with the "can't delete last paragraph in a doc" issue.
* Returns the line spacing of that paragraph to be applied to the new object.
*/
function deleteParagraph(docbody, paragraph) {
  var linespacing = paragraph.asParagraph().getLineSpacing();
  try {
    docbody.removeChild(paragraph);
  } catch(e) {
    if (e.message.indexOf("remove the last paragraph") !== -1) {
      docbody.appendParagraph(' ');
      docbody.removeChild(paragraph);
    } else {
      throw e;
    }
  }
  return linespacing;
}

/**
 * Search for two lines starting with ``` in the doc,
 * and add all the lines appearing between them into a single-cell
 * table, set the font to a monospace font.
 */
function processSourceCode(doc) {
  var body = doc.getBody();

  var startingTripleTick = body.findText('```');
  if (!startingTripleTick) return;
  var endTripleTick = body.findText('```', startingTripleTick);
  if (!endTripleTick) return;

  var firstLine = startingTripleTick.getElement();
  var lastLine = endTripleTick.getElement();

  var rangeBuilder = doc.newRange();
  rangeBuilder.addElementsBetween(firstLine, lastLine);
  var range = rangeBuilder.build();
  var lineRanges = range.getRangeElements();
  var lines = [];

  var firstLineIndex = body.getChildIndex(lineRanges[0].getElement());
  var code = "";

  // Don't iterate over 0th and last line because they are the tripleticks
  lineRanges[0].getElement().removeFromParent();
  for (var i = 1; i < lineRanges.length - 1; ++i) {
    code += lineRanges[i].getElement().asText().getText() + '\n';
    lineRanges[i].getElement().removeFromParent();
  }
  lineRanges[lineRanges.length-1].getElement().removeFromParent();

  var cell = body.insertTable(firstLineIndex)
                 .setBorderWidth(0)
                 .appendTableRow()
                 .appendTableCell();

  var params = {
    'code': code.trim(),
    'lexer': /```(.*)/.exec(firstLine.asText().getText())[1],
    'style': 'monokai'
  };
  var response = UrlFetchApp.fetch(
    "http://hilite.me/api",
    {
      'method': 'post',
      'payload': params
    }
  );

  var xmlDoc = XmlService.parse(response.getContentText());
  // The XML document is structured as
  // - comment
  // - div
  //   - pre
  //     - spans
  var divTag = xmlDoc.getAllContent()[1];
  var preTag = divTag.getAllContent()[0];
  var spans_or_texts = preTag.getAllContent();
  var span_ranges = [];

  var startCharIdx = 0;
  for (var i = 0; i < spans_or_texts.length; ++i) {
    var span_or_text = spans_or_texts[i];
    if (span_or_text.getType() == XmlService.ContentTypes.ELEMENT) {
      // We are seeing a span (spans are styled while texts are not)
      var span_range = {
        start: startCharIdx,
        endInclusive: startCharIdx + span_or_text.getValue().length - 1,
        span: span_or_text
      };
      span_ranges.push(span_range);
    }
    startCharIdx += span_or_text.getValue().length;
  }

  var getTagColor = function (tag) {
    return tag.getAttribute('style').getValue().match(/#[0-9 a-f A-F]{6}/);
  };

  cell.setText(preTag.getValue().trim());

  var cellText = cell.editAsText();
  for (var i = 0; i < span_ranges.length; ++i) {
    var span_range = span_ranges[i];
    cellText.setForegroundColor(
      span_range.start,
      span_range.endInclusive,
      getTagColor(span_range.span)
    );
  }
  cell.setBackgroundColor(getTagColor(divTag));
  cell.setFontFamily('Consolas');
  numChanged++;
  processSourceCode(doc);
}


/**
 * Search for `some text` and replace it with
 * its backtick-free version with a monospace font
 * (uses slack color theme presently)
 */
function processBackquotes(docbody) {
  var backquote = docbody.findText('`[^`]+?`');
  if (backquote) {
    var start = backquote.getStartOffset();
    var end = backquote.getEndOffsetInclusive();
    var text = backquote.getElement().asText();
    text.setBackgroundColor(start, end, '#f0f0f2');
    text.setFontFamily(start, end, 'Consolas');
    text.setForegroundColor(start, end, '#cc2255');
    text.deleteText(end, end);
    text.deleteText(start, start);
    numChanged++;
    processBackquotes(docbody);
  }
}


/**
 * Search for **some text** and replace it with its
 * asterisk-free version with a bold face.
 */
function processBold(docbody) {
  var bold = docbody.findText('\\*\\*[^*]+?\\*\\*');
  if (bold) {
    var start = bold.getStartOffset();
    var end = bold.getEndOffsetInclusive();
    var text = bold.getElement().asText();
    text.setBold(start, end, true);
    text.deleteText(end-1, end);
    text.deleteText(start, start+1);
    numChanged++;
    processBold(docbody);
  }
}


/**
 * Search for _some text_ and replace it with its
 * underscore-free, italicized version.
 */
function processItalics(docbody) {
  var italics = docbody.findText('_[^_]+?_');
  if (italics) {
    var start = italics.getStartOffset();
    var end = italics.getEndOffsetInclusive();
    var text = italics.getElement().asText();
    text.setItalic(start, end, true);
    text.deleteText(end-1, end-1);
    text.deleteText(start+1, start+1);
    numChanged++;
    processItalics(docbody);
  }
}


/**
 * Convert patterns of the form [Link Name](http://example.com/address)
 * to hyperlinks where the link text is "Link Name" and
 * the link url is "http://example.com/address"
 */
function processLinks(docbody) {
  // Links are of the form "[Link Name](http://example.com/page/address)"
  var link = docbody.findText('\\[.*?\\]\\(https?:\\/\\/.*?\\)');
  if (link) {
    var start = link.getStartOffset();
    var end = link.getEndOffsetInclusive();
    var text = link.getElement().asText();
    var linkName = text.getText().split('[')[1].split(']')[0];
    var url = text.getText().split(']')[1].split('(')[1].split(')')[0];
    text.deleteText(start, end);
    text.insertText(start, linkName);
    text.setLinkUrl(start, start + linkName.length - 1, url);
    numChanged++;
    processLinks(docbody);
  }
}


/**
 * Do the following conversions:
 * # my heading   -> "my heading" styled as Heading1
 * ## another one -> "another one" styled as Heading2
 * ### third      -> "third" styled as Heading3
 */
function processHeadings(docbody) {
  var headingStarts = ['# ', '## ', '### '];
  var headingFormats = [
    DocumentApp.ParagraphHeading.HEADING1,
    DocumentApp.ParagraphHeading.HEADING2,
    DocumentApp.ParagraphHeading.HEADING3
  ];
  for (var i = 0; i < headingStarts.length; ++i) {
    var headingStart = headingStarts[i];
    var heading = docbody.findText('^\\s*'+ headingStart + '.*');
    while (heading) {
      if (heading.getStartOffset() == 0) {
        var elem = heading.getElement();
        elem.asText().deleteText(0, i+1);
        while (elem.getType() != DocumentApp.ElementType.PARAGRAPH) {
          elem = elem.getParent();
        }
        elem.setHeading(headingFormats[i]);
        numChanged++;
      }
      heading = docbody.findText(headingStart + '.*', heading);
    }
  }
}


/**
 * Convert ordered bulleted items (starting with a number and a '.')
 * to ordered list items and
 * unordered bulleted items (starting with '*', '-' or '+')
 * to unordered list items.
 */
function processLists(docbody) {
  var docchildren = docbody.getNumChildren();
  var listregex = [
    {
      regex : '^[*\\-+]\\s+',
      glyph : DocumentApp.GlyphType.BULLET
    },
    {
      regex : '^\\d+\\.\\s+',
      glyph : DocumentApp.GlyphType.NUMBER
    }
  ];

  for(var i = 0; i < docchildren; i++ ) {
    var paragraph = docbody.getChild(i);
    if(paragraph.getType() == DocumentApp.ElementType.PARAGRAPH) {
      listloop: for(var j = 0; j < listregex.length; j++) {
        var list = paragraph.findText(listregex[j].regex + '.*');
        if (list) {
          var text = list.getElement().asText().replaceText(listregex[j].regex, '').getText();
          var linespacing = deleteParagraph(docbody, paragraph);
          docbody.insertListItem(i, text);
          docbody.getChild(i).asListItem().setGlyphType(listregex[j].glyph);
          docbody.getChild(i).asListItem().setLineSpacing(linespacing);
          numChanged++;
          break listloop;
        }
      }
    }
  }
}


// FUNCTIONS FOR ULTRADOX

/*
* Main function for formatting markdown from Ultradox
* docid : Google Doc ID passed as parameter
* (Where to find in URL: https://docs.google.com/document/d/<docid>/edit#)
*/
function formatMarkdownUltradox(docid) {
  var doc = DocumentApp.openById(docid);
  splitParagraphs(doc);
  formatMarkdownOfDoc(doc);
}


/**
* Variables inserted into a template by Ultradox as "${variable;string(wrap)}"
* use the "vertical tab" (\v) character to wrap lines. This replaces them
* with standard newline (\n) characters so that the wrapped lines are treated as
* separate paragraphs, which is necessary for some of the functions to work.
*/
function splitParagraphs(doc) {
  var docbody = doc.getBody();
  var docchildren = docbody.getNumChildren();
  for(var i = 0; i < docchildren; i++) {
    var paragraph = docbody.getChild(i);
    if(paragraph.getType() == DocumentApp.ElementType.PARAGRAPH) {
      var softreturn = paragraph.asText().replaceText("\\v", "\\n").getText().split("\\n");
      if (softreturn.length > 1) {
        var linespacing = deleteParagraph(docbody, paragraph);
        for(var j = 0; j < softreturn.length; j++) {
          docbody.insertParagraph(i + j, softreturn[j]);
          docbody.getChild(i + j).asParagraph().setLineSpacing(linespacing);
        }
      docchildren += softreturn.length-1;
      i += softreturn.length-1;
     }
   }
 }
}


// The following method will be invoked whenever Ultradox executes the script
function execute(model) {
  formatMarkdownUltradox(model.docID);
  return {
    executionReport : "Updated doc ID " + model.docID + " - " + numChanged + " elements updated",
    elementsChanged : numChanged
  }
}


// Return a description of your script
function getDescription(items) {
  return {
    input : [
          { name : 'docID',
            prompt : 'Enter the ID of document to be updated:',
            description : 'ID of document to be updated',
            defaultValue: '',
            format : 'STRING',
            required : true
          }
    ],
    output : [
          { name : 'elementsChanged',
            description : 'Contains the number of elements updated by the script',
            format : 'NUMBER'
          }
    ],
    icon : undefined, // Optional: URL to 64x64 .png icon
    modelRequired : false, // Only set to true if you need access to the full data model
    inputPrefix : '', // Assign a default input prefix
    outputPrefix : '' // Assign a default output prefix
  }
}


// DO NOT CHANGE OR REMOVE THE FOLLOWING CODE
function doGet(request) {
  try {
    var ultradocId = request.parameters.ultradocId;
    var itemId = request.parameters.itemId;
    var html = "<html><head><script>function init() { gapi.client.load('ultradox', 'v1', function() { gapi.client.ultradox.scriptAuthorized({ 'ultradocId': '"+ultradocId+"', 'itemId' : '"+itemId+"'}).then(function(resp) { console.log('Script authorized'); }, function(reason) { console.log('Error: ' + reason.result.error.message);});}, 'https://oauth2-dot-floreysoftudx.appspot.com/_ah/api');}</script><script src='https://apis.google.com/js/client.js?onload=init'></script></head><body style='height:500px;background-image:url(https://www.ultradox.com/ultradoxBg.png);background-repeat: no-repeat;background-position: right top;'><table cellspacing='25px' width='450px'><tr><td rowspan='2'><img src='https://www.ultradox.com/ultradoxOk.png'></td><td style='font:22px Ubuntu'>Access granted</td></tr><tr valign='top'><td style='color:#999;font:16px Ubuntu'>You can now close this window and use the script.</td></tr></table></body></html>";
    return HtmlService.createHtmlOutput(html).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  } catch ( err ) {
    var msg = err;
    if ( typeof err == Error ) {
      msg = err.message;
    }
    return ContentService.createTextOutput("{ __error : '"+msg+"' }").setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(request) {
  try {
    var command = request.parameters.cmd;
    var json = request.postData.getDataAsString();
    var model = JSON.parse(json);
    var result;
    if (command == 'description') {
      result = getDescription(model);
    } else {
      result = execute(model);
    }
    var json = JSON.stringify(result);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  } catch ( err ) {
    var msg = err;
    if ( typeof err == Error ) {
      msg = err.message;
    }
    return ContentService.createTextOutput("{ __error : '"+msg+"' }").setMimeType(ContentService.MimeType.JSON);
  }
}
