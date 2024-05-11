// This function is triggered when the Google Doc is opened.
// It adds a custom menu item to the Google Docs UI.
function onOpen() {
  DocumentApp.getUi().createMenu('HTML')
    .addItem('Conversion', 'applyStyles') // Adds a menu item that triggers `applyStyles()` when clicked.
    .addToUi(); // Finalizes the UI update.
}

// Main function that coordinates the style application process.
function applyStyles() {
  applyParagraphHeadings(); // Apply heading styles based on HTML tags.
  applyAttributes(); // Apply text attributes like bold, italic, etc.
  removeHTMLTags(); // Clean up HTML tags from the text.
}

// Apply text attributes such as bold, italic based on HTML-like tags found in the document.
function applyAttributes() {
  let body = DocumentApp.getActiveDocument().getBody();
  let text = body.editAsText();

  // Define mapping of HTML-like tags to Google Docs text attributes.
  let AttributePatterns = {
    'b': [DocumentApp.Attribute.BOLD, true],
    'i': [DocumentApp.Attribute.ITALIC, true],
    'u': [DocumentApp.Attribute.UNDERLINE, true],
    's': [DocumentApp.Attribute.STRIKETHROUGH, true],
    'a': [DocumentApp.Attribute.LINK_URL, null]
  };

  // Loop through each tag to apply corresponding attributes.
  Object.keys(AttributePatterns).forEach(tag => {
    let attribute = AttributePatterns[tag][0];
    let value = AttributePatterns[tag][1];
    let openTag = `\\<${tag}\\>`;
    let closeTag = `\\</${tag}\\>`;
    let searchPattern = tag === 'a' ? `\\<a href="([^"]+)"\\>(.*?)\\</a\\>` : openTag + "(.*?)" + closeTag;
    let searchResult = text.findText(searchPattern);

    while (searchResult) {
      let foundText = searchResult.getElement().asText();
      let startOffset = searchResult.getStartOffset() + openTag.length - 2;
      let endOffset = searchResult.getEndOffsetInclusive() - closeTag.length + 3;

      // Special handling for links.
      if (tag === 'a') {
        let url = searchResult.getText().match(/href="([^"]+)"/)[1];
        let linkTextStart = searchResult.getText().indexOf('>') + 1;
        let linkTextEnd = searchResult.getText().lastIndexOf('<') - 1;
        foundText.setLinkUrl(startOffset + linkTextStart, startOffset + linkTextEnd, url);
      } else {
        if (startOffset < endOffset) {
          foundText.setAttributes(startOffset, endOffset, { [attribute]: value });
        }
      }

      searchResult = text.findText(searchPattern, searchResult);
    }
  });
}

// Function to apply paragraph headings based on HTML-like tags.
function applyParagraphHeadings() {
  let body = DocumentApp.getActiveDocument().getBody();

  // Map of HTML-like tags to Google Docs heading styles.
  let HeadingPatterns = {
    'h1': DocumentApp.ParagraphHeading.TITLE,
    'h2': DocumentApp.ParagraphHeading.HEADING1,
    'h3': DocumentApp.ParagraphHeading.HEADING2,
    'h4': DocumentApp.ParagraphHeading.HEADING3,
    'h5': DocumentApp.ParagraphHeading.HEADING4,
    'h6': DocumentApp.ParagraphHeading.HEADING5,
    'p' : DocumentApp.ParagraphHeading.NORMAL, 
  }

  Object.keys(HeadingPatterns).forEach(tag => {
    let heading = HeadingPatterns[tag];
    let openTag = `\\<${tag}\\>`;
    let closeTag = `\\</${tag}\\>`;
    let searchPattern = openTag + "(.*?)" + closeTag;
    let searchResult = body.findText(searchPattern);

    while (searchResult) {
      let foundTextElement = searchResult.getElement();
      let foundTextParent = foundTextElement.getParent();

      // Apply the heading style if the found text is within a paragraph.
      if (foundTextParent.getType() == DocumentApp.ElementType.PARAGRAPH) {
        let paragraph = foundTextParent.asParagraph();
        paragraph.setHeading(heading);
      }
      
      searchResult = body.findText(searchPattern, searchResult);
    }
  });
}

// Removes all HTML-like tags from the document.
function removeHTMLTags() {
  let body = DocumentApp.getActiveDocument().getBody();
  let content = body.editAsText();
  
  let text = content.getText();
  let regex = /<[^>]*>/g; // Regular expression to match all tags.
  let matches = [];

  // Collect all matches for removal.
  let result;
  while ((result = regex.exec(text)) !== null) {
    matches.push({start: result.index, end: regex.lastIndex});
  }

  // Remove tags in reverse order to avoid disrupting the indexes.
  for (var i = matches.length - 1; i >= 0; i--) {
    content.deleteText(matches[i].start, matches[i].end - 1);
  }
}
