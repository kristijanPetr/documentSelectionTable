function onOpen() {
  DocumentApp.getUi().createMenu('Selection')
    .addItem("Add a Task", 'insertTableAtSelection' )
    .addToUi();
}


function insertTableAtSelection() {
  // insert table as selection
  var theDoc = DocumentApp.getActiveDocument();
  var selection = theDoc.getSelection();
  Logger.log('selection: ' + selection);
  
  if (selection) {
    var elements = selection.getRangeElements();
    var element;
    var arr = [];
    for (var i = 0; i < elements.length; i++) {
      element = elements[i];
      var theElmt = element;
      var selectedElmt = theElmt.getElement();
  
      var parent = selectedElmt.getParent();
      
      var insertPoint = theDoc.getBody().getChildIndex(parent);
      
      if(theElmt.isPartial()){
        arr.push(selectedElmt.asText().getText().substring(theElmt.getStartOffset(), theElmt.getEndOffsetInclusive()+1))
        selectedElmt.asText().deleteText(theElmt.getStartOffset(), theElmt.getEndOffsetInclusive());
      }else{
        arr.push(selectedElmt.asText().getText())
        selectedElmt.asText().removeFromParent();
      }
    }
 
    Logger.log('insertPoint: ' + insertPoint);    
    if(arr.length < 1){
      arr = 'Task Body \n ... \n ... \n <task break> \n ... \n ... \n ... ';
    }
    else{
      arr.unshift('Task Body');
      arr = arr.join('\n');
    }
    var body = theDoc.getBody();
    var style = {};
    style[DocumentApp.Attribute.BOLD] = true;
    var par = body.insertParagraph(insertPoint+1, '<task>');
    par.setAttributes(style);
    var table = body.insertTable(insertPoint + 2, [['Enter Task Title '],
                                 [arr],
                                 ['Copyrights information']]);
    
  }
}

function selectTable(){
   var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
   var rangeElem = selection.getRangeElements();
  for(var i=0;i < rangeElem.length;i++){
    var element = rangeElem[i].getElement();
 
    if(element.getType() == 'TABLE_CELL'){
      var table = element.asTableCell().getParentTable();
    // var width = table.getCell(1, 1).getWidth();
      var rows = table.getNumRows();
      
      Logger.log(table.getChild(1).asTableRow().getWidth() + " ch" )
    }
    else if(element.getType() == 'TABLE_ROW'){
      var table = element.asTableRow().getParentTable();
      Logger.log(table.getNumRows() + " " );
    }
  }
}


function selectionMakeSameWidth () {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();

  var rangeElem = selection.getRangeElements();
  Logger.log(rangeElem)
  for(var i=0;i< rangeElem.length;i++){
      var element = rangeElem[i].getElement();
      var startOffset = rangeElem[i].getStartOffset();      // -1 if whole element
      var endOffset = rangeElem[i].getEndOffsetInclusive(); // -1 if whole element
      var selectedText = element.asText().getText();       // All text from element
    Logger.log(doc.getBody().getChildIndex(element.getParent()));
    
      // Is only part of the element selected?
    if (rangeElem[i].isPartial()){
      Logger.log(selectedText.substring(startOffset,endOffset+1))
//      selectedText = selectedText.substring(startOffset,endOffset+1);
    }
//    Logger.log(rangeElem[i].getElement().asText().getText())
  }
  var ui = DocumentApp.getUi();

  var report = "Your Selection: ";

  if (!selection) {
    report += " No current selection ";
  }
  else {
    var elements = selection.getRangeElements();
    // Report # elements. For simplicity, assume elements are paragraphs
    report += " Paragraphs selected: " + elements.length + ". ";
    if (elements.length > 1) {
    }
    else {
      var element = elements[0].getElement();
      var startOffset = elements[0].getStartOffset();      // -1 if whole element
      var endOffset = elements[0].getEndOffsetInclusive(); // -1 if whole element
      var selectedText = element.asText().getText();       // All text from element
      // Is only part of the element selected?
      if (elements[0].isPartial())
        selectedText = selectedText.substring(startOffset,endOffset+1);

      // Google Doc UI "word selection" (double click)
      // selects trailing spaces - trim them
      selectedText = selectedText.trim();
      endOffset = startOffset + selectedText.length - 1;

      // Now ready to hand off to format, setLinkUrl, etc.

      report += " Selected text is: '" + selectedText + "', ";
      report += " and is " + (elements[0].isPartial() ? "part" : "all") + " of the paragraph."
    }
  }
  Logger.log(report)
 
}
