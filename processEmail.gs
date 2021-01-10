// Listen for email containing shift times from when2work.com
function listen() {
  var label = GmailApp.getUserLabelByName("Schedule");
  var threads = label.getThreads();  
  if(threads.length > 0) { 
    When2WorkAlert.main(); 
  }
  else { 
    Logger.log("W2W Email has not been recieved");  
  }
}

