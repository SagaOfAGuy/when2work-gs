// Main function 
function main() 
{ 
  var filename = "Schedule.ics";
  var email = ""; 
  createICS(filename, email); 
}

// Create ICS file by searching through email inbox and extracting shift times from W2W email
function createICS(filename, email) {
  try { 
    var targetEmails = searchEmail("W2W Schedule");

   // Shifts extracted from W2W email
    var workShifts = getShifts(targetEmails);  
    var workShiftsString = workShifts.toString();
    var workShiftsArray = workShiftsString.split(/[m][,]/g);
    var years = workShiftsString.match(/\s[0-9][0-9][0-9][0-9]\s/g); 
    var fileInfo = {
     title:`${filename}`,
    mimeType: 'text/calendar'
    }; 
   var startFile = writeStart() +"\n";
   var string = ``; 
   string += startFile;
   for (var index = 0; index < workShiftsArray.length; index++) {
     var block = workShiftsArray[index].split('-');
     var year = years[index].toString(); 
     var timeBlock = block[1].split('to');
     var monthDateBlock = block[0]; 
     var month = convertMonth(monthDateBlock.substring(0,3)); 
     var dayRegexResult = monthDateBlock.match(/([0-9]|[0-9][0-9])[,]/g)
     var day = dayRegexResult.toString().padStart(3,'0').substring(0,2); 
     var start = makeMilitary(timeBlock[0]).padStart(4,'0'); 
     var end = makeMilitary(timeBlock[1]).padStart(4,'0');
     string += writeMiddle('VCU',year.trim(),month, day, start, end);
     string += "\n"
   }
   var endFile = writeEnding(); 
   string += endFile; 
   var blob = Utilities.newBlob(string); 
   Drive.Files.insert(fileInfo, blob);
   var filename = "Schedule.ics"; 
   sendICS(filename, email); 
  }
 catch(err) { 
  Logger.log("Email does not exist"); 
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

// Get file ID in Google Drive based on file ID
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
  var hour = Number(timeBlock.substring(0,2));
  var minutes = timeBlock.substring(3,5); 
  if(timeBlock.includes("p")) { 
     hour += 12; 
  }
  return hour.toString() + minutes;   
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
  var threads = GmailApp.getInboxThreads(0, 10);
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

// Get the year of work shift
function getYear(emails) { 
  for(var index = 0; index < emails.length; index++) { 
    var str = emails[index].getRawContent(); 
    return str; 
  }
}
