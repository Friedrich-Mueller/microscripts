/**
 * @OnlyCurrentDoc
 */
function colorTagsGreen() {
  var body = DocumentApp.getActiveDocument().getBody();
  var pattern = "\\[[^\\]]+\\]"; // Matches anything inside [ square brackets ]
  var found = body.findText(pattern);
  
  while (found) {
    var element = found.getElement().asText();
    var start = found.getStartOffset();
    var end = found.getEndOffsetInclusive();
    
    element.setForegroundColor(start, end, "#008000"); // Dark Green
    element.setBold(start, end, true);                // Bold
    element.setFontFamily(start, end, "Roboto Mono"); // Monospace / Robotic font
    
    found = body.findText(pattern, found);
  }
}