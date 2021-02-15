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

/** 
 * SERVICES
 * --------
 *  - Drive API V2
 *  
 * 
*/

/**
 * DEPENDENCIES
 * ------------
 *  - Create a 'Schedule' Label
 *  - Create a Label filter for emails that contain the subject line 'A Change To Your Schedule' and 'W2W Schedule'
 * 
 */
// Entry point 
function main() {
    let label = GmailApp.getUserLabelByName("Schedule");
    let threads = label.getThreads();
    if(threads.length >= 1) {
          let workShiftsObject = getWorkShifts(); 
          let shifts = workShiftsObject.shifts; 
          let shiftsString = writeShifts(workShiftsObject);
          if(shifts == undefined || shifts == '' || shifts == null || shiftsString == undefined) {
            Logger.log("No new Shits(s) from When2Work.com"); 
          }
          else {
            createICS('Schedule.ics',shiftsString); 
            // Email goes below
            sendICS('Schedule.ics', shiftsString.subjectString, ''); 
          }
          threads[0].removeLabel(label); 
    }
    else {
        Logger.log("No emails in Schedule Label"); 
    }
}
function getWorkShifts(){
  let messageContent = createMessageString(0,1); 
  let regex; 
  let updatedWorkShifts = (messageContent.includes("shifts has changed"));    
  let cancelledWorkShifts =  (messageContent.includes("no longer")); 
  let newWorkShifts = (messageContent.includes("new")); 
  let emailType = '';
  if(updatedWorkShifts) {regex = /\w\w\w\s([0-9][0-9]|[0-9])\w[,]\s[0-9][0-9][0-9][0-9]\s\s\s\s\s\s\w\w\w\s\w\w\w\w\s\w\w\w\w\w\s\s\s\s\s\s([0-9][0-9]|[0-9]|[0-9][0-9][:][0-9][0-9]|[0-9][:][0-9][0-9])(\w\w)\s\w\w\s([0-9][0-9][:][0-9][0-9]|[0-9][:][0-9][0-9]|[0-9][0-9]|[0-9])(\w\w|\w)/g; emailType='update'}; 
  if(newWorkShifts) {regex = /\w\w\w\s([0-9]|[0-9][0-9])[,]\s[0-9][0-9][0-9][0-9]\s[-]\s([0-9]\w\w|[0-9][0-9]\w\w|[0-9][:][0-9][0-9]\w\w|[0-9][0-9][:][0-9][0-9]\w\w)\s\w\w\s([0-9]\w\w|[0-9][:][0-9][0-9]\w\w)/g; emailType='new'}; 
  if(cancelledWorkShifts) {regex = /\w\w\w\s([0-9]|[0-9][0-9])[,]\s[0-9][0-9][0-9][0-9]\s\s\s\s\s\s\w\w\w\s\w\w\w\w\s\w\w\w\w\w\s\s\s\s\s\s([0-9][0-9]|[0-9])[:][0-9][0-9]\w\s\w\w\s([0-9][0-9]|[0-9])[:][0-9][0-9]\w/g; emailType='cancelled'};
  let shiftMatches = messageContent.match(regex); 
  let shiftMatchesObject = {shifts:shiftMatches,type:emailType}
  return shiftMatchesObject; 
}
function createMessageString(start,end) { 
  let firstThread = GmailApp.getInboxThreads(start,end)[0];
  let messages = firstThread.getMessages();
  let messageContent;
  for(let index = 0; index < messages.length; index++) {
    messageContent += messages[index].getPlainBody();
    messageContent += " "; 
  } return messageContent; 
}
function writeShifts(workShiftsObject) {
  try {
  let shifts = workShiftsObject.shifts;
  let string = ``; 
  let subject = 'Schedule | New Shift(s)'; 
  let fileHeading = writeStart() + "\n"; 
  let fileEnding = writeEnding(); 
  let indexMap = {monthIndex:0,dayIndex:1,yearIndex:2,startIndex:4,endIndex:6};
  let stringObject; 
  if(workShiftsObject.type != 'new') {
      indexMap.startIndex = 12;
      indexMap.endIndex =  14; 
      if(workShiftsObject.type == 'cancelled') {
        subject = 'Schedule | Shifts Cancelled';  
        // Email goes below
        MailApp.sendEmail('',subject, `These Work shifts been Cancelled:\n\n${shifts}`);  
        Logger.log("Shift Cancellation Email Sent"); 
        return ; 
      }
      else {
        subject = 'Schedule | Shift Change'; 
      }
    }  
    string += fileHeading; 
    for (index = 0; index < shifts.length; index++) {
      let shiftPieces = shifts[index].toString().split(' '); 
      let month = convertMonth(shiftPieces[indexMap.monthIndex]); 
      let day = shiftPieces[indexMap.dayIndex].substring(0,shiftPieces[indexMap.dayIndex].indexOf(',')); 
      let year = shiftPieces[indexMap.yearIndex]; 
      let startShift = makeMilitary(shiftPieces[indexMap.startIndex]); 
      let endShift = makeMilitary(shiftPieces[indexMap.endIndex]); 
      string += writeMiddle('Work',year.trim(),month,day,startShift,endShift); 
      string += "\n"; 
    }
    string += fileEnding; 
    stringObject = {shiftString:string, subjectString:subject}; 
    return stringObject; 
  }
  catch(TypeError) {
    return ; 
  }
  }
  
function createICS(filename,obj) {
  if(obj == undefined) {
    return ; 
  } 
  let fileInfo = {
     title:`${filename}`,
     mimeType: 'text/calendar'
    };
  let blob = Utilities.newBlob(obj.shiftString); 
  Drive.Files.insert(fileInfo, blob);
  Logger.log("ICS File Created");  
}
function sendICS(filename, subject, email) {
  let ICSFileID = getFileId(filename);
  let ICSFile = DriveApp.getFileById(ICSFileID);
  MailApp.sendEmail(`${email}`,subject,'Work Schedule',{
    attachments: [ICSFile],
    name: 'Attachment'
  }); 
  Logger.log('ICS File Sent'); 
  Drive.Files.remove(ICSFileID); 
} 
function makeMilitary(time) { 
    let isAfternoon = time.includes('p'); 
    let notNoon = !(time.includes('12')); 
    let hasMinutes = time.includes(':'); 
    let times; 
    let milTime; 
    let minutes = '00';
    let totalTime;  
 	if(hasMinutes) { 
       times = time.split(':'); 
       milTime = times[0]; 
       minutes = times[1].match(/([0-9][0-9]|[0-9])/g);
    }else {
    	times = time.match(/([0-9][0-9]|[0-9])/g)
        milTime = times; 
    }
  if(isAfternoon && notNoon) { 
    	milTime = Number(milTime); 
        milTime += 12; 
  }
  totalTime = `${milTime}${minutes}`; 
	return totalTime.padStart(4,0);
}
function convertMonth(monthText) { 
   let months = {
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
function getFileId(name) { 
  let files = DriveApp.getFiles(); 
  while(files.hasNext()) { 
     let file = files.next();
     let filename = file.getName();
     let fileID = file.getId(); 
    if(name == filename) { 
       return fileID; 
    }    
  }
}
function writeStart() { 
  let start = `BEGIN:VCALENDAR
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
function writeMiddle(summary, year, month, day, start, end) { 
  let middle = `BEGIN:VEVENT
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
function writeEnding() { 
   let end = 'END:VCALENDAR';
   return end; 
}
