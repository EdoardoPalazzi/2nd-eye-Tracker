// author: Edoardo Palazzi

var WebWhooklink = //insert link of gchat webhook

// button to trigger automatic assignment

function onOpen() {
  //creating the button
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Assign');
  menu.addItem('Assign Reviewer', 'assignByEqualDistribution');
  menu.addToUi();
}


// function to assign reviewers

function assignByEqualDistribution(e) {
  //importGoogleCalendar();
  oooStatus();

  // list of reviewers
  var reviewers = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('Assignees').getValues(); 
  reviewers = reviewers.filter(row => 
    row != '' && row != 'Reviewers' //filter out any blank rows or rows named 'Assignee' --- deleted row, might be able to delete this
  )

  // all people that reviewed a task
  var reviewees = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(3, 9, 1000).getValues();

  // create dictionary
  var dict = {};
  var numOfReviews = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange(3, 1, 50, 2).getValues();

  numOfReviews.forEach((v) => {
    let key = v[0];
    let value = v[1];

    dict[key] = value;
  })
  for (key in dict){
    if (key == ""){
    delete dict[key];
        }
  }  

  // order dict
  var items = Object.keys(dict).map(
  (key) => { return [key, dict[key]] });
  items.sort(
    (first, second) => { return first[1] - second[1] }
  );

  var dictReviewees = {};

  items.forEach((v) => {
    let key = v[0];
    let value = v[1];

    dictReviewees[key] = value;
  })
  console.log(dictReviewees);

  //count the # of projects without a reviewer assigned
  var count = 3;
  for(i=0; i < reviewees.length; i++) {
    if (reviewees[i] != ""){
      count += 1;
    }
  }
  // get the number of tasks that don't have a reviewer assigned
  var unassigned = 0;
  var idAdded = 0;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tracker");
  var data = sheet.getRange(count, 1, 100, 6).getValues();
  data.forEach(function (row) {
    if (row[0] != ""){
      unassigned += 1;
      //console.log(row);
    }
    if (row[5] != ""){
      idAdded += 1;
    }
  });
  //console.log(unassigned);
  //console.log(idAdded);

  // assign reviewer(s)
  /*
  Conditions to add:
  - if someone reviewed the patterns then it should assign the same person for the Project 4 (if that person is in office)
  - check if reviewer is not OOO
  - check if reviewer was not assigned a review in the last 5 requested projects 
  - check if reviewer is not the requester
  */
  var i = 0;
  var numOfReviewers = Object.keys(dictReviewees).length;
  //console.log(Object.keys(dictReviewees)[1]);
  console.log(count);
  if (unassigned > 0 && idAdded == unassigned){
    var data1 = sheet.getRange(count, 2, unassigned, 5).getValues(); // data of unassigned tasks
    var data2 = sheet.getRange(3, 2, count-unassigned-2, 8).getValues(); // data of all tasks assigned
    var assigned = 0; // number of reviewers assigned from the dictReviewees dict
    while (i < unassigned){
      if (data1[i][0] == "Project 4"){
        var taskId = data1[i][4];
        var reviewer = "";
        data2.forEach(function (row) {
          if (row[4] == taskId){
            reviewer = row[7];
          }
        });
        console.log(reviewer);
        var reviewers1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange(3, 1, 50, 3).getValues();
        reviewers1 = reviewers1.filter(row => 
          row[0] != '' && row[0] != 'Assignee' //filter out any blank rows or rows named 'Assignee' --- deleted row, might be able to delete this
        )
        var dict2 = {};
        reviewers1.forEach((v) => {
          let key = v[0];
          let value = v[2];

          dict2[key] = value;
        })

        if (dict2[reviewer] == "No" ){
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tracker").getRange(count+i, 9).setValue(reviewer);
          var link = "https://docs.google.com/spreadsheets/d/15ryKlYv8d_8cep2jkmqzpzzrbHPUnxgj08TmAPIZitU/edit#gid=0&range=I";
          var lastRow = count + i;
          link += lastRow;
          var bug = "https://thisIsTheLink.com/issues/" + data1[i][4];
          var message = { text: "*<@" + newReviewer + ">*" + ", you have been assigned a new review!" + "\n*Please take a look here:* " + link + "\n\n Assign yourself as a reviewer to this bug: " + bug + "\n\n _P.S. If you are too busy and can't review this task please ping your TL and the QL in this chat to let them know :)"};
          var payload = JSON.stringify(message);
          var options = {
                  method: 'POST',
                  contentType: 'application/json',
                  payload: payload
          };
          var response =  UrlFetchApp.fetch(WebWhooklink, options ).getContentText();
          var email = newReviewer + "@google.com";
          MailApp.sendEmail({to: email, subject: "New Assigned Review", htmlBody: "<b>" + newReviewer + "</b>" + ", you have been assigned a new review!" + "<hr><b>Please take a look here:</b> " + link + "<hr><hr> Assign yourself as a reviewer to this bug: " + bug + "<hr><hr> <i>P.S. If you are too busy and can't review this task please ping your TL and the QL in this chat to let them know :)</i>"});
        }
        else{
          if (count < 8){  // change here if looking at different amount of reviewers
            var lastFive = [];
          }
          else{
            var lastF = sheet.getRange(count - 5 + i, 9, 5).getValues();
            var lastFive = [];
            //console.log(lastFive);
            for (row of lastF) for (e of row) lastFive.push(e);
          }
          var list = Object.keys(dictReviewees);
          
          for (j = 0; j < list.length; j++){
            if (lastFive.indexOf(list[j]) == -1 && dict2[list[j]] == "No" && data1[i][2] != list[j]){
              var newReviewer = list[j];
              //console.log(newReviewer);
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tracker").getRange(count+i, 9).setValue(newReviewer);
              var link = "https://docs.google.com/spreadsheets/d/15ryKlYv8d_8cep2jkmqzpzzrbHPUnxgj08TmAPIZitU/edit#gid=0&range=I";
              var lastRow = count + i;
              link += lastRow;
              var bug = "https://thisIsTheLink.com/issues/" + data1[i][4];
              var message = { text: "*<@" + newReviewer + ">*" + ", you have been assigned a new review!" + "\n*Please take a look here:* " + link + "\n\n Assign yourself as a reviewer to this bug: " + bug + "\n\n _P.S. If you are too busy and can't review this task please ping your TL and the QL in this chat to let them know :)_"};
              var payload = JSON.stringify(message);
              var options = {
                      method: 'POST',
                      contentType: 'application/json',
                      payload: payload
              };
              var response =  UrlFetchApp.fetch(WebWhooklink, options ).getContentText();
              var email = newReviewer + "@google.com";
            MailApp.sendEmail({to: email, subject: "New Assigned Review", htmlBody: "<b>" + newReviewer + "</b>" + ", you have been assigned a new review!" + "<hr><b>Please take a look here:</b> " + link + "<hr><hr> Assign yourself as a reviewer to this bug: " + bug + "<hr><hr> <i>P.S. If you are too busy and can't review this task please ping your TL and the QL in this chat to let them know :)</i>"});
              break;        
            }      
          }
        }
      }
      else{
        // get the ldaps of the last 5 reviewers
        if (count < 8){  // change here if looking at different amount of reviewers
          var lastFive = [];
        }
        else{
          var lastF = sheet.getRange(count - 5 + i, 9, 5).getValues();
          var lastFive = [];
          //console.log(lastFive);
          for (row of lastF) for (e of row) lastFive.push(e);
        }
        var list = Object.keys(dictReviewees);
        // dict2 is to get the OOO status of the reviewers
        var reviewers1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange(3, 1, 50, 3).getValues();
        reviewers1 = reviewers1.filter(row => 
          row[0] != '' && row[0] != 'Assignee' 
        )
        var dict2 = {};
        reviewers1.forEach((v) => {
          let key = v[0];
          let value = v[2];

          dict2[key] = value;
        })
        
        for (j = 0; j < list.length; j++){
          if (lastFive.indexOf(list[j]) == -1 && dict2[list[j]] == "No" && data1[i][2] != list[j]){
            var newReviewer = list[j];
            //console.log(newReviewer);
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tracker").getRange(count+i, 9).setValue(newReviewer);

            var link = "https://docs.google.com/spreadsheets/d/15ryKlYv8d_8cep2jkmqzpzzrbHPUnxgj08TmAPIZitU/edit#gid=0&range=I";
            var lastRow = count + i;
            link += lastRow;
            var bug = "https://thisIsTheLink.com/issues/" + data1[i][4];
            var message = { text: "*" + newReviewer + "*" + ", you have been assigned a new review!" + "\n*Please take a look here:* " + link + "\n\n Assign yourself as a reviewer to this bug: " + bug + "\n\n _P.S. If you are too busy and can't review this task please ping your TL and the QL in this chat to let them know :)_"};
            var payload = JSON.stringify(message);
            var options = {
                    method: 'POST',
                    contentType: 'application/json',
                    payload: payload
            };
            var response =  UrlFetchApp.fetch(WebWhooklink, options ).getContentText();
            var email = newReviewer + "@google.com";
            MailApp.sendEmail({to: email, subject: "New Assigned Review", htmlBody: "<b>" + newReviewer + "</b>" + ", you have been assigned a new review!" + "<hr><b>Please take a look here:</b> " + link + "<hr><hr> Assign yourself as a reviewer to this bug: " + bug + "<hr><hr> <i>P.S. If you are too busy and can't review this task please ping your TL and the QL in this chat to let them know :)</i>"});

            break;        
          }      
        }       
      }
      
      i +=1;
    }
}
else{
  //SpreadsheetApp.getUi().alert("Make sure that at least one task is missing an assigned reviewer or that all tasks have an ID in column F");
  SpreadsheetApp.getUi().alert("ATTENTION", "Make sure that at least one task is missing an assigned reviewer or that all tasks have an ID in column F", SpreadsheetApp.getUi().ButtonSet.OK);
}
}




// google calendar data

function importGoogleCalendar() { 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar Data");
  var calendarId = sheet.getRange('B1').getValue().toString(); 
  var calendar = CalendarApp.getCalendarById(calendarId);

  var range = SpreadsheetApp
               .getActive()
               .getSheetByName("Calendar Data")
               .getRange(7,1, 1000, 4);
  range.clearContent();
 
  // Set filters
  var startDate = sheet.getRange('B2').getValue();
  var endDate = sheet.getRange('B3').getValue();
  var searchText = sheet.getRange('B4').getValue();
  if (searchText == "All"){
    searchText = "";
  }
 
  // Print header
  var header = [["Creator", "Title", "Start", "End"]];
  var range = sheet.getRange("A6:D6");
  range.setValues(header);
  range.setFontWeight("bold");
 
  // Get events based on filters
  var events = (searchText == '') ? calendar.getEvents(startDate, endDate) : calendar.getEvents(startDate, endDate, {search: searchText});
  //var events = calendar.getEvents(new Date(startDate), new Date(endDate));  
 
  // Display events 
  for (var i=0; i<events.length; i++) {
    var row = i+7;
    
    var guests = events[i].getGuestList(true);
    var details = [[events[i].getCreators(), events[i].getTitle(), events[i].getStartTime(), events[i].getEndTime()]];
    
    range = sheet.getRange(row,1,1,4);
    range.setValues(details);
 
    // Format the Start and End columns
    var cell = sheet.getRange(row, 4);
    cell.setNumberFormat('mm/dd/yyyy hh:mm');
    cell = sheet.getRange(row, 5);
    cell.setNumberFormat('mm/dd/yyyy hh:mm');

  }
}

function oooStatus(){

  var dataOOO = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar Data").getRange("F7:G100").getValues();
  dataOOO = dataOOO.filter(row => 
    row[0] != '' && row[0] != 'Assignee' //filter out any blank rows or rows named 'Assignee' --- deleted row, might be able to delete this
  )
  console.log(dataOOO);

  // keeping only the people that are OOO today in dataOOO
  for(i=0; i < dataOOO.length; i++){
    if (dataOOO.length > 0 && dataOOO[i][0] == "No"){
      dataOOO.splice(i, 1);
      i --;
    }
  }
  console.log(dataOOO);

  var reviewers1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange(3, 1, 50, 3).getValues();
  reviewers1 = reviewers1.filter(row => 
    row[0] != '' && row[0] != 'Assignee'
  )
  var dict2 = {};
  reviewers1.forEach((v) => {
    let key = v[0];
    let value = v[2];

    dict2[key] = value;
  })

  console.log(Object.keys(dict2).length);
  if (dataOOO.length == 0){
    for (i=0; i < Object.keys(dict2).length; i++){
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange(3+i, 3).setValue("No");
    }
  }
  else{
    var i = 0;
    var keyOOO = "No";
    for (key in dict2){
      keyOOO = "No";
      for (j=0; j<dataOOO.length; j++){
        if (dataOOO[j][1] == key){
          keyOOO = "Yes";
        }
      }
      console.log(key);
      console.log(keyOOO);
      if (keyOOO == "Yes"){
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange(3+i, 3).setValue("Yes");
      }
      else{
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange(3+i, 3).setValue("No");
      }
      i += 1;
      
    }
  }

}
