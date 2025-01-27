/**
 * Replaces one or more placeholders in the doc with inline charts/images,
 * preserving their exact location (even if in a table, header, etc.).
 *
 * @param {Document} doc - The Google Doc object where placeholders are replaced.
 * @param {Array} chartMappings - An array of objects describing which chart to place:
 *   [
 *     {
 *       placeholder: '{{CHART_EXAMPLE}}', // string to find in doc
 *       blob: <Blob>,                    // e.g. chart.getAs('image/png')
 *       maxWidthPx: 400                  // optional max width in px
 *     },
 *     {
 *       placeholder: '{{CHART_AUTOCONSUMO}}',
 *       blob: <Blob>,
 *       maxWidthPx: 500
 *     }
 *   ]
 *
 * Example usage:
 *   insertChartsAtExactLocation(doc, [
 *     { placeholder: '{{CHART_1}}', blob: chartBlob, maxWidthPx: 400 },
 *     { placeholder: '{{CHART_2}}', blob: chartBlob2 }
 *   ]);
 */


  

function insertCharts(doc, chartMappings) {
  const body = doc.getBody();

  // Helps ensure placeholders are in one piece of text
  body.editAsText();

  chartMappings.forEach(mapping => {
    const placeholder = mapping.placeholder;
    const chartBlob = mapping.blob;
    const maxWidth = mapping.maxWidthPx || 360; // default max width if not specified

    if (!placeholder || !chartBlob) {
      Logger.log('Skipping invalid mapping: ' + JSON.stringify(mapping));
      return;
    }

    // Repeatedly find this placeholder in the doc
    let foundRange = body.findText(placeholder);
    while (foundRange) {
      // Force-get the text containing the match
      const text = foundRange.getElement().asText();
      const startOffset = foundRange.getStartOffset();
      const endOffset = foundRange.getEndOffsetInclusive();

      // 1) Remove the placeholder text from the doc
      text.deleteText(startOffset, endOffset);

      // 2) Insert the image into the parent element (paragraph, cell, list item, etc.)
      const parentElem = text.getParent();
      let inlineImage;

      // If the parent is a Paragraph, we might center it, etc.
      // Otherwise, we still just append the image inline.
      if (parentElem.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const paragraph = parentElem.asParagraph();
        inlineImage = paragraph.appendInlineImage(chartBlob);
        // Optionally center the entire paragraph
        paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      } else {
        // e.g. table cell, header footer, list item
        inlineImage = parentElem.appendInlineImage(chartBlob);
      }

      // 3) Enforce a max width (if the original is larger)
      const currentWidth = inlineImage.getWidth();
      const currentHeight = inlineImage.getHeight();
      if (currentWidth > maxWidth) {
        const ratio = maxWidth / currentWidth;
        inlineImage.setWidth(maxWidth);
        inlineImage.setHeight(Math.floor(currentHeight * ratio));
      }

      Logger.log('Inserted chart for placeholder: ' + placeholder);

      // 4) Search for the next occurrence of the same placeholder
      foundRange = body.findText(placeholder, foundRange);
    }
  });
}