
/**
 * Copyright Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
// Main function 
function main() 
{ 
  var filename = "Schedule.ics";
  // Change this to the email you want When2Work ICS file emailed to
  var email = ""; 
  
  var label = GmailApp.getUserLabelByName("Schedule");
  var threads = label.getThreads();

  //console.log(targetEmails); 
  //console.log(workShifts); 

  if(threads.length >= 1) {
    createICS(filename, email); 
    threads[0].removeLabel(label);  
  }
  else 
  {
    console.log("No email from W2W"); 
  } 
}
// Create ICS by searching through inbox for emails from W2W 
function createICS(filename, email) { 
  try {
    
   // Emails recieved from W2W 
    var targetEmails = searchEmail("W2W Schedule");
   // Shifts extracted from W2W email
    var regex = /\w\w\w\s([0-9]|[0-9][0-9])[,]\s[0-9][0-9][0-9][0-9]\s[-]\s([0-9]\w\w|[0-9][0-9]\w\w|[0-9][:][0-9][0-9]\w\w|[0-9][0-9][:][0-9][0-9]\w\w)\s\w\w\s([0-9]\w\w|[0-9][:][0-9][0-9]\w\w)/g
    var workShifts = getShifts(targetEmails);  
    var workShiftsString = workShifts.toString();
    var workShiftsArray = workShiftsString.match(regex);
    var years = workShiftsString.match(/\s[0-9][0-9][0-9][0-9]\s/g);  

     
   var startFile = writeStart() +"\n";
   var string = ``; 
   string += startFile;
   for (var index = 0; index < workShiftsArray.length; index++) {
     var block = workShiftsArray[index].split('-');
     var year = years[index].toString(); 
     var timeBlock = block[1]; 
     var monthDateBlock = block[0].toString(); 
     var month = convertMonth(monthDateBlock.substring(0,3)); 
     var splitTime = timeBlock.split('to'); 
     var dayRegexResult = monthDateBlock.match(/([0-9]|[0-9][0-9])[,]/g)
     var day = dayRegexResult.toString().padStart(3,'0').substring(0,2); 
     
     var startShift = splitTime[0];  
     var endShift = splitTime[1]; 
  
     var start = makeMilitary(startShift);  
     var end = makeMilitary(endShift); 
    
     string += writeMiddle('VCU',year.trim(),month, day, start, end);
     string += "\n"; 
   }
   var fileInfo = {
     title:`${filename}`,
     mimeType: 'text/calendar'
    };
   var endFile = writeEnding(); 
   string += endFile;  
   var blob = Utilities.newBlob(string); 
   Drive.Files.insert(fileInfo, blob);
   sendICS(filename, email);
   console.log("ICS file sent!"); 
  }
  catch(err) {
    Logger.log("Email does not exist"); 
    Logger.log(err);
    // Put your email below no W2W email is received 
    MailApp.sendEmail(email, 'Alert', 'No W2W email'); 
    return ; 
  }
}

// Send the ICS File via email
function sendICS(filename, email) { 
  var ICSFileID = getFileId(filename);
  var ICSFile = DriveApp.getFileById(ICSFileID);
  
  MailApp.sendEmail(`${email}`,'VCU W2W Schedule','Schedule',{
    attachments: [ICSFile],
    name: 'Attachment'
  }); 
  Drive.Files.remove(ICSFileID); 
}
// Get files in google drive
function getFileId(name) { 
  var files = DriveApp.getFiles(); 
  while(files.hasNext()) { 
     var file = files.next();
     var filename = file.getName();
     var fileID = file.getId(); 
    if(name == filename) { 
       return fileID; 
    }    
  }
}
// Convert Times into military time 
function makeMilitary(timeBlock) { 
  // Variables
  var regex = null
  var hourText = ''
  var hour = ''; 
  var minutes = '00'; 
  var totalTime = ''; 
 
  // Conditionals 
  var isPM = timeBlock.includes("p");
  var notNoon = null 
  var minutesIndex = timeBlock.indexOf(":"); 
  var noMinutes = timeBlock.indexOf(":") == -1; 
  
  
  if(noMinutes) {
    regex = /([0-9][0-9]|[0-9])/g
    hourText = timeBlock.match(regex); 
    hour = Number(hourText); 
  }
  else {
  	hourText = timeBlock.substring(0,minutesIndex); 
    hour = Number(hourText); 
    minutes = timeBlock.substring(minutesIndex+1,minutesIndex+3);
  }
  notNoon = hour != 12; 
  if(isPM && notNoon) {
  	hour += 12; 
  }
  totalTime = `${hour}${minutes}`; 
  return totalTime.padStart(4,'0'); 
}
// Return Beginning boilerplate for ICS file 
function writeStart() { 
  var start = `BEGIN:VCALENDAR
PRODID:-//Google Inc//Google Calendar 70.9054//EN
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:PUBLISH
X-WR-CALNAME:Work Schedule
X-WR-TIMEZONE:America/New_York
BEGIN:VTIMEZONE
TZID:America/New_York
X-LIC-LOCATION:America/New_York
BEGIN:DAYLIGHT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE`
  return start;  
}
// Boilerplate for actual ICS content containing work shifts 
function writeMiddle(summary, year, month, day, start, end) { 
  var middle = `BEGIN:VEVENT
DTSTART;TZID=America/New_York:${year}${month}${day}T${start}00
DTEND;TZID=America/New_York:${year}${month}${day}T${end}00
RRULE:FREQ=DAILY;COUNT=1
DSTAMP:20201108T014109Z
CREATED:20201108T014107Z
LAST-MODIFIED:20201108T014108Z
LOCATION:
SEQUENCE:0
STATUS:CONFIRMED
SUMMARY: ${summary}
TRANSP:OPAQUE
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:This is an event reminder
TRIGGER:-P0DT4H0M0S
DESCRIPTION:
END:VALARM
END:VEVENT`
 return middle;  
}
// Convert month from text to number 
function convertMonth(monthText) { 
   var months = {
    'Jan' : '01',
    'Feb' : '02',
    'Mar' : '03',
    'Apr' : '04',
    'May' : '05',
    'Jun' : '06',
    'Jul' : '07',
    'Aug' : '08',
    'Sep' : '09',
    'Oct' : '10',
    'Nov' : '11',
    'Dec' : '12'
    }
   return months[monthText]; 
}
// Return ending boilerplate for ICS file 
function writeEnding() { 
   var end = 'END:VCALENDAR';
   return end; 
}
// Function for searching through email 
function searchEmail(search)
{
  var threads = GmailApp.getInboxThreads(0, 1);
  for (var i = 0; i < threads.length; i++) 
  {
    if(threads[i].getFirstMessageSubject() == search)
    {
      var email = threads[i].getMessages();
    } 
  }
  return email; 
}
// Extract shifts from targeted emails 
function getShifts(emails) {
   var shifts_text = []
   for (var index = 0; index < emails.length; index++)
    {
        var str = emails[index].getRawContent(); 
        var matchIndex = str.search("day"); 
        var shifts = str.substring(matchIndex, matchIndex+435); 
        var regex = /\w\w\w\s([0-9]|[0-9][0-9])[,]\s[0-9][0-9][0-9][0-9]\s[-]\s([0-9]\w\w|[0-9][0-9]\w\w|[0-9][:][0-9][0-9]\w\w|[0-9][0-9][:][0-9][0-9]\w\w)\s\w\w\s([0-9]\w\w|[0-9][:][0-9][0-9]\w\w)/g
        var matches = shifts.match(regex);
        shifts_text.push(matches); 
    }
  return shifts_text; 
}
function getYear(emails) { 
  for(var index = 0; index < emails.length; index++) { 
    var str = emails[index].getRawContent(); 
    return str; 
  } 
}
