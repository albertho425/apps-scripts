//***************************************************************************
  
/**
 *  Display the menu in the spreadsheet
 *  */  

function initMenu() {
  
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Custom Menu");
  
  menu.addItem("Get Calendar Events","getCalendarEvents");
  menu.addItem("Instructions","info");
  menu.addItem("Clear the spreadsheet" , "clearSheet");
  menu.addItem("Delete Events From Calendar", "delete_events");
  menu.addSeparator();


  // render the menus and sub menus
  menu.addToUi();

}

//***************************************************************************

/**
 * When the spreadsheet opens, run the functions
 */

function onOpen()
{
 initMenu(); 
 info();
}

//***************************************************************************

/**
 * Get calendar events and put them into the spreadsheet
 */

function getCalendarEvents(){


    // this is the testing calendar
    var mycal = "c_ljfdbnmht0ca7hr6mj3ncvk440@group.calendar.google.com";

    var cal = CalendarApp.getCalendarById(mycal);

    var events = cal.getEvents(new Date("April 1, 2021 00:00:00 PST"), new Date("December 31, 2021 23:59:59 PST"), {search: '-Huddle'});


    var sheet = SpreadsheetApp.getActiveSheet();
    // sheet.clearContents();  

    var header = [["Event ID", 
                  "Event Name",  
                  "Event Details", 
                  "Rooms and Resources", 
                    "Event Start", 
                    "Event End", 
                    "Date Created", 
                    "Date Updated"]]
    // var range = sheet.getRange(1,1,1,8);
    var range = sheet.getRange(1,1,1,8);
    range.setValues(header);

    var addedEvents = 0;  
    // Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
    for (var i=0;i<events.length;i++) {
    var row=i+2;
    var myformula_placeholder = '';

    var details=[[events[i].getId(), 
                events[i].getTitle(), 
                events[i].getDescription(), 
                events[i].getLocation(), 
                events[i].getStartTime(), 
                events[i].getEndTime(), 
                events[i].getDateCreated(), 
                events[i].getLastUpdated()]];
    var range=sheet.getRange(row,1,1,8);
    range.setValues(details);


    addedEvents++;

    }
    Logger.log("Total events added: " + addedEvents);


}
//***************************************************************************

/**
 *  Clear the spreadsheet
 */

function clearSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    sheet.clear({ formatOnly: false, contentsOnly: true })
}

//***************************************************************************

/**
 *  Info box
 */

function info() {
  Browser.msgBox('Instructions', '1) This script will run from the script editor\\n2) It can also run from the spreadsheet.\\n3) Have a good day', Browser.Buttons.OK);
}

/**
 *  Delete events from the calendar
 */
//***************************************************************************

function deleteEvents(){

    var fromDate = new Date("April 1, 2021 00:00:00 PST");
    var endDate = new Date("December 31, 2021 23:59:59 PST");
    var calendarName = "c_ljfdbnmht0ca7hr6mj3ncvk440@group.calendar.google.com";
    var mycal = "c_ljfdbnmht0ca7hr6mj3ncvk440@group.calendar.google.com";
    var cal = CalendarApp.getCalendarById(mycal);

    var now = new Date();
    var twoHoursFromNow = new Date(now.getTime() + (2 * 60 * 60 * 1000));
    // var events = cal.getEvents(new Date("January 1, 2020 00:00:00 PST"), new Date("December 31, 2020 23:59:59 PST"), {search: '-project123'});


    // var calendar = CalendarApp.getCalendarsByName(mycal)[0];
    var cal = CalendarApp.getCalendarsByName('testing calendar');
    Logger.log(cal);
    var events = CalendarApp.getDefaultCalendar().getEvents(now, twoHoursFromNow);
    // var events = cal.getEvents(new Date("January 1, 2020 00:00:00 PST"), new Date("December 31, 2020 23:59:59 PST"));
    Logger.log(events);

    Logger.log(events.length);

    for (var i=0; i<events.length;i++) {
      var ev = events[i];
      Logger.log(ev.getTitle());
      Logger.log("inside for loop");
      ev.deleteEvent();
    }
      
}
//***************************************************************************

/**
 *  Delete events from the calendar
 */

function delete_events()
{
  //Please note: Months are represented from 0-11 (January=0, February=1). Ensure dates are correct below before running the script.
  var fromDate = new Date(2020,0,1,0,0,0); //This represents June 1st 2020
  var toDate = new Date(2020,11,31,0,0,0); //This represents June 30th 2020
  var calendarID = 'c_ljfdbnmht0ca7hr6mj3ncvk440@group.calendar.google.com'; //Enter your calendar ID here

  var calendar = CalendarApp.getCalendarById(calendarID);
  
  //Search for events between fromdate and todate with given search criteria
  var events = calendar.getEvents(fromDate, toDate);
  // var events = calendar.getEvents(fromDate, toDate,{search: 'Sim Man'});
  var counter = 0;
  for(var i=0; i<events.length;i++) //loop through all events
  {
    var ev = events[i];
    Logger.log('Event: '+ev.getTitle()+' found on '+ev.getStartTime()); // Log event name and title
      ev.deleteEvent(); // delete event
      counter++;
  }
  Logger.log("Total events deleted: " + counter);

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Total Events Deleted in Calendar: ' + counter, ui.ButtonSet.OK);


}
//***************************************************************************

function delete_parts_of_sheet()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange("A2:O7");
  range.clear();
}



