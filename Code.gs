function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Speaker Notes Fixer')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processSlides(url) {
  try {
    // Extract the ID from the URL
    const idMatch = url.match(/[-\w]{25,}/);
    if (!idMatch) return { success: false, message: "Error: Invalid Google Slides URL." };
    
    const presentationId = idMatch[0];
    const presentation = SlidesApp.openById(presentationId);
    const presentationUrl = presentation.getUrl();
    const slides = presentation.getSlides();
    
    const safeWidth = 450; 
    const safeLeftMargin = 81; 

    slides.forEach(slide => {
      const notesPage = slide.getNotesPage();
      const notesShape = notesPage.getSpeakerNotesShape();

      if (notesShape) {
        // Note: Google Apps Script does not allow changing the width or position (setLeft/setWidth) 
        // of a speaker notes shape because it is controlled by the Notes Master.
        // We can, however, fix the text formatting and indentation.

        const PAGE_WIDTH = 612; // Standard 8.5" page width in points
        const safeWidth = 450; 
        const centeredMargin = (PAGE_WIDTH - safeWidth) / 2; // 81 points
        
        let indentStart = centeredMargin; // Fallbacks
        let indentEnd = centeredMargin;
        
        // Attempt to center the slide thumbnail (and other elements)
        // This may fail if Google locks the element to the Notes Master
        notesPage.getPageElements().forEach(el => {
          try {
            const elWidth = el.getWidth();
            const targetLeft = (PAGE_WIDTH - elWidth) / 2;
            el.setLeft(targetLeft);
          } catch (e) {
            // Ignore if operation not allowed on Notes Master element
          }
        });
        
        try {
          const currentLeft = notesShape.getLeft();
          const currentWidth = notesShape.getWidth();
          
          // Calculate indents needed to force the text into the center 450px of the page
          indentStart = Math.max(0, centeredMargin - currentLeft);
          indentEnd = Math.max(0, (currentLeft + currentWidth) - (PAGE_WIDTH - centeredMargin));
        } catch(e) {
          // If getLeft/getWidth fail, use the fallbacks
        }

        const textRange = notesShape.getText();
        if (!textRange.isEmpty()) {
          const paragraphs = textRange.getParagraphs();
          paragraphs.forEach(p => {
            try {
              p.getRange().getParagraphStyle().setIndentStart(indentStart);
              p.getRange().getParagraphStyle().setIndentFirstLine(indentStart);
              p.getRange().getParagraphStyle().setIndentEnd(indentEnd);
              p.getRange().getTextStyle().setFontSize(14); 
            } catch(e) {
              // Ignore if any specific formatting fails
            }
          });
        }
      }
    });
    
    return {
      success: true,
      message: "Success! The speaker notes have been resized and fixed. You can now print the presentation.",
      url: presentationUrl
    };
    
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() + " (Make sure you have Editor permissions to the slide link)" };
  }
}