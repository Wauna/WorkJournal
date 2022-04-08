const monthNames = ["January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];
const doc = DocumentApp.getActiveDocument();


function createMonthlyJournal(){
  try
    {
      var dt = new Date();
          
      /** SET THE HEADER **/
      var header= doc.getHeader();
      header.clear();
       // GET THE PARAGRAPH
        for (let i = 0; i < header.getNumChildren(); i += 1) {
          var p1 = header.getChild(i);
          p1.setText(monthNames[dt.getMonth()] + " " + dt.getFullYear() + " Work Journal");
          p1.setHeading(DocumentApp.ParagraphHeading.TITLE);
          p1.setSpacingBefore(0);
          p1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);          
        }
    
      /** BUILD THE BODY **/
      var body = doc.getBody();
      body.clear();

      /** LOOP THROUGH DATES */
      var currMonth = dt.getMonth();
      var loopDate = dt;
      while (loopDate.getMonth()== currMonth) {

        if(loopDate.getDay() > 0 && loopDate.getDay() < 6){
            addDay(loopDate,body);
        } 
        loopDate.setDate(dt.getDate() + 1);
      }


    }
    catch(err){
      Logger.log('ERROR: ' + err.message);

    }
}

function addDay(dt, body){

      body.appendPageBreak();
      
      /** ADD DAYS TO THE DOC **/

      var pageHeader = body.appendParagraph(dt.toDateString());      
      pageHeader.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      pageHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      pageHeader.setSpacingBefore(0);
      var pH = pageHeader.editAsText();
      pH.setBold(true);
      
      /** ADD GOALS */
      //addSection("Daily Goals",body);
      addGoalEffortSection(body);

      /** ADD NOTES */
      addSection("Notes",body);
      
      /** ADD ACTIONS */
      addSection("Actions",body);
      
      /** ADD MEETINGS */
      addSection("Meetings",body);

}

function addGoalEffortSection(body){
      var tbl = body.appendTable();
      var headerRow = tbl.appendTableRow();
      tbl.setBorderWidth(1);
      tbl.setBorderColor('#f2f2f2');

      var headingCell = headerRow.appendTableCell();
      headingCell.clear();
      
      headingCell.setBackgroundColor("#f2f2f2");
   
      // GET THE PARAGRAPH
      for (let i = 0; i < headingCell.getNumChildren(); i += 1) {
        var pG = headingCell.getChild(i);
        pG.setText("DAILY GOALS")
        pG.setSpacingAfter(0);
        pG.setSpacingBefore(0);
        pG.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        
        var t = pG.editAsText();
        t.setForegroundColor("#000000");
        t.setBold(true);
        t.setFontSize(12);      
      }


      headingCell = headerRow.appendTableCell();
      headingCell.clear();
      headingCell.setBackgroundColor("#f2f2f2");
      
      // GET THE PARAGRAPH
      for (let i = 0; i < headingCell.getNumChildren(); i += 1) {
      
        var pG = headingCell.getChild(i);
        pG.setText("EFFORTS & TASKS")
        pG.setSpacingAfter(0);
        pG.setSpacingBefore(0);
        pG.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        

        var t = pG.editAsText();
        t.setForegroundColor("#000000");
        t.setBold(true);
        t.setFontSize(12);      
      }

    var contentRow = tbl.appendTableRow();
    var cell1 = contentRow.appendTableCell();
    cell1.setBold(false);
    cell1.setFontSize(10);
    var li = cell1.appendListItem("Goal 1");
    cell1.getChild(0).removeFromParent();


    var cell2 = contentRow.appendTableCell();
    cell2.setBold(false);
    cell2.setFontSize(10);
    cell2.appendListItem("Item1");
    cell2.getChild(0).removeFromParent();
    
    // Add a little space
    body.appendParagraph("");
    
}

function addSection(heading, body){
      var tbl = body.appendTable();
      var row = tbl.appendTableRow();
      var cell = row.appendTableCell();
      cell.clear();
      tbl.setBorderWidth(0);

      cell.setBackgroundColor("#f2f2f2")
      
      // GET THE PARAGRAPH
      for (let i = 0; i < cell.getNumChildren(); i += 1) {
      
        var pG = cell.getChild(i);
        pG.setText(heading)
        pG.setSpacingAfter(0);
        pG.setSpacingBefore(0);
        
        var t = pG.editAsText();
        t.setForegroundColor("#000000");
        t.setBold(true);
        t.setFontSize(14);      
      }

   
    
      body.appendParagraph("");
      body.appendParagraph("");
      body.appendParagraph("");
}




function menuItem4(){
    var sel = doc.getSelection();
    var elements = sel.getSelectedElements();
    Logger.log(elements.length);

    for (let i = 0; i < elements.length; i += 1) {
              var range   = elements[i];
              var element = range.getElement();
              Logger.log(element.getNumChildren());
              var e2 = element.getChild[0];
              Logger.log(e2);

              Logger.log(element);
             
     }
    
    Logger.log("Hook Called");
}


function getToday(){
  const content = doc.getBody().getParagraphs()
  var dt = new Date();
  const searchTerm = dt.toDateString();

  for (var item in content){
      Logger.log(item);
  }
  
  Logger.log(searchTerm);
  const position = content.indexOf(searchTerm, content.indexOf(searchTerm)+1);
  Logger.log(position.toString());
    

}


function onOpen() {
  var ui = DocumentApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Work Journal')
      .addItem('Create Monthly Journal', 'menuItem1')
      .addItem('Add Today\'s Meetings', 'menuItem2')
      .addItem('Settings', 'menuItem3')
      .addItem('HOOK', 'menuItem4')
      
      
      //.addSeparator()
      //.addSubMenu(ui.createMenu('Sub-menu')
      //    .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function menuItem1() {
    createMonthlyJournal();
     
}

function menuItem2() {
  addUpcomingEvents();
}

function menuItem3() {
 var ui = DocumentApp.getUi();
 var html = HtmlService.createHtmlOutputFromFile('Settings.html');
 ui.showModalDialog(html,"Work Journal Settings");

}


function buildNotes(id){
  DocumentApp.getUi() // Or DocumentApp or FormApp.
  .alert('You clicked the item -' + id);
}


function addUpcomingEvents() {
  try {

  const calendarId = 'primary';
  const dt = (new Date());
 
  const dtMin = new Date(dt.getFullYear(), dt.getMonth(),dt.getDate(),0,0,0,0);
  dt.setDate(dt.getDate() + 1);
  const dtMax = new Date(dt.getFullYear(), dt.getMonth(),dt.getDate(),0,0,0,0);

  // Add query parameters in optionalArgs
  const optionalArgs = {
    timeMin: (dtMin.toISOString())
    ,showDeleted: false
    ,singleEvents: true
    ,maxResults: 10
    ,orderBy: 'startTime'
    ,timeMax:(dtMax.toISOString())
    // use other optional query parameter here as needed.
  };

    // Define a custom paragraph style.
    const eventStyle = {};
        eventStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
        eventStyle[DocumentApp.Attribute.BOLD] = false;

    const timeStyle = {};
        timeStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
        timeStyle[DocumentApp.Attribute.BOLD] = false;

    const idStyle = {};
        idStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
        idStyle[DocumentApp.Attribute.BOLD] = false;
        idStyle[DocumentApp.Attribute.FOREGROUND_COLOR="#ffffff"];



    // call Events.list method to list the calendar events using calendarId optional query parameter
    const response = Calendar.Events.list(calendarId, optionalArgs);
    const events = response.items;
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const pos = doc.getCursor();
    const element = pos.getElement();
    const parent = element.getParent();
    const tz = response.timeZone;
  
    if (events.length === 0) {
      Logger.log('No upcoming events found');
      return;
    }
  
    var tbl = body.insertTable(parent.getChildIndex(element)+1);
    tbl.setBorderColor("#f2f2f2");
  
    // Print the calendar events
    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      let when = event.start.dateTime;
      if (!when) {
        when = event.start.dateTime;
      }
        
      if(event.visibility != 'private'){
        var startDate = new Date(event.start.dateTime);
        var endDate = new Date(event.end.dateTime);
       
      var row = tbl.appendTableRow();
    
      // ADD THE EVENT COLUMN
      var tcEventSummary = row.appendTableCell();

      // -- GET THE FIRST(DEFAULT) PARARGRAPH
      tcEventSummary.setWidth(400);
      // GET THE PARAGRAPH
      for (let i = 0; i < tcEventSummary.getNumChildren(); i += 1) {
        var pG = tcEventSummary.getChild(i);
        pG.setLinkUrl("https://www.google.com");
        pG.setText(event.summary);
        pG.setAttributes(eventStyle);
        Logger.log(pG.getLinkUrl());
        
        }

      //ADD THE TIME
      var tcEventTime = row.appendTableCell();
      
      // GET THE PARAGRAPH
      for (let i = 0; i < tcEventTime.getNumChildren(); i += 1) {
        var pg2 = tcEventTime.getChild(i);
        pg2.setText(Utilities.formatDate(startDate,tz,"hh:mm") + ' - ' + Utilities.formatDate(endDate,tz,"hh:mm"));
        pg2.setAttributes(timeStyle);
        }


      //row.appendTableCell(event.id);

      }
    }

  } catch (err) {
    // TODO (developer) - Handle exception from Calendar API
    Logger.log('Failed with error %s', err.message);
  }
}
