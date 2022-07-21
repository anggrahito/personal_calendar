function createCalenderEvent() {

  var personalCalender = CalendarApp.getCalendarById("calendarId");
  var sheet = SpreadsheetApp.getActiveSheet();
  var schedule = sheet.getDataRange().getValues();
  var index = 3;
  var lastRow = sheet.getLastRow();
  var fromDate = new Date('1/1/2022');
  var toDate = new Date('12/31/2022');

/* data cleaning process */

    var clear = personalCalender.getEvents( fromDate, toDate); /* Data is setup for 1 year, from January 1st 2022 - December 31st 2022 */
    clear_calendar(clear);

/* Input process from Google sheet to Google Calendar */

  for (;index <=lastRow; index++) {

      var number = sheet.getRange(index, 1, 1, 1).getValue();
      var title = sheet.getRange(index, 2, 1, 1).getValue();
      var startDate = sheet.getRange(index, 3, 1, 1).getValue();
      var endDate = sheet.getRange(index, 4, 1, 1).getValue();
      var partnerName = sheet.getRange(index, 5, 1, 1).getValue();
      var partnerEmail = sheet.getRange(index, 6, 1, 1).getValue();
      var description = sheet.getRange(index, 7, 1, 1).getValue();
      var status = sheet.getRange(index, 8, 1, 1).getValue();


  if (startDate && endDate && status !='Finish' )
  {
    var events = personalCalender.getEvents( startDate, endDate);
    delete_events(events);
    var calendar = personalCalender.createEvent(title, startDate, endDate, {description: description});
  }
}

/* sync update button in Google Sheet */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
      .addItem('Sync Now',
              'createCalenderEvent')
      .addToUi();
}

/* Remove all the events */
function delete_events(events) {

  for(var i=0; i<events.length;i++){
    var ev = events[i];
    ev.deleteEvent();
  }
}

/* Cleaning the data function */
function clear_calendar(clear) {

  for(var i=0; i<clear.length;i++){
    var cl = clear[i];
    cl.deleteEvent();
  }
}

}
