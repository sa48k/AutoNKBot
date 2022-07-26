/* AutoNKB
   On Classroom, assign NK forms to students for each year level
   and Generate Google Slides with names and year levels
   30 May 2019 */

// Test class ID is 40492//
// Real NK class is 50331//
// Comms class is 2769283//
// Pohu HL 2020 is 556337//
// 
// groupings master sheet id = 1X_SEaakzSSLtJ*
// prettified sheet (real names) id, don't touch this = 1EYF_dqEAv_uU*
// slides id = 1j_gJCpEvNf*

function main() {
  var students = listStudents();
  for (var x=0; x<8; x++) {  // for years 1-8
    generateSlides(x);
    var returnedArray = getStudentsForForm(x, students);
    var assignees = returnedArray[0];
    var wk = returnedArray[1];
    var formUrl = returnedArray[2]  // TODO: Wipe the URL // maybe after successful post only
    var scheduleTime = returnedArray[3]
    if (assignees.length > 0) {                            // If assignees[] is empty, skip postCourseWork()
      postCourseWork(formUrl, assignees, x+1, wk, scheduleTime);
    }
  }
}

function generateSlides(yr) {
  var ss = SpreadsheetApp.openById("1EYF_dqEAv_uUx2Zw1vF1*"); // in-VLOOKUPs-out
  var sheet = ss.getSheets()[2]; // the output sheet
  var data = sheet.getDataRange().getValues();
  var c = [];
  for (var i=2; i<data.length; i++) {                      // grab data for the column yr and write to c[]
    c.push(data[i][yr]) 
  }                                                        // TODO: Sort names alphabetically and write back to range on sheet
  var col = c.filter(Boolean);                             // removes empty strings from array and sorts A-Z
  col.sort();
  var presentation = SlidesApp.openById('1j_gJCpEvNfMCZp*')
  var slide = presentation.getSlides()[yr];
  var elements = slide.getPageElements();
  for (var m=0; m<elements.length; m++) {
    if (elements[m].getPageElementType() == 'TABLE') {     // Remove existing tables on slide to make room for the new one
      elements[m].remove()
    }
  }
  if (col.length < 9) {
    var rows = col.length; } else { var rows = 8;              // table always has eight rows, unless there are <8 in the group
  }                 
  if (rows == 0) { var rows = 1; }                             
  var columns = Math.ceil(col.length/8);                             // 2 cols for 9-16, 3 cols for 17-24, etc.
  if (columns == 0) { var columns = 1; }   // hack
//  var title = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 40, 10, 400, 30); // left, top, width, height

//  table.setBorderColor('#ff0000');
//  var textRange = title.getText();
//  textRange.setText('Y' + (yr+1) + ' Number Knowledge');
  var fontSize = 10; // catchall for surprisingly large groups
  switch (columns) {case 1: var fontSize = 24; break;     // TODO: we can have more rows when fontSize is smaller (see y4, y5)
                case 2: var fontSize = 24; break;
                case 3: var fontSize = 22; break;
                case 4: var fontSize = 16; break;
                case 5: var fontSize = 14; break;
  }
//  textRange.getTextStyle().setFontFamilyAndWeight('Open Sans', 600).setFontSize(24);
  var table = slide.insertTable(rows, columns, 10, 60, 500, 200);  //
  for (var c = 0; c < columns; c++) {
    for (var r = 0; r < rows; r++) {             // write the contents of col to the new table
      if (col.length == 0) { break; }            // pop from start of array to table, until col is empty
      var paste = table.getCell(r, c).getText().setText(col.shift())
        .getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      table.getCell(r, c).getText().getTextStyle().setFontSize(fontSize).setFontFamilyAndWeight('Raleway', 300);
      table.getCell(r, c).setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    }
  }
//  var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 520, 10, 200, 300);  // 'instructions' text box
//  var textRange = shape.getText();
//  textRange.setText('Instructions');
//  textRange.getTextStyle().setFontFamilyAndWeight('Open Sans', 300).setFontSize(24);
  switch(yr){
    case 0: var bkgColor = '#FFE6F4'; break;     // tbh this doesn't need to be set each time, just like the title box and instruction box
    case 1: var bkgColor = '#F8E6FF'; break;     // can probably take this out
    case 2: var bkgColor = '#EDEBFF'; break;
    case 3: var bkgColor = '#EBF6FF'; break;
    case 4: var bkgColor = '#EBFFF8'; break;
    case 5: var bkgColor = '#F0FFEB'; break;
    case 6: var bkgColor = '#FDFFEB'; break;
    case 7: var bkgColor = '#FFF3EB'; break;      
  }
  slide.getBackground().setSolidFill(bkgColor);
}

function getStudentsForForm(yr, students) {
  // open GSheet containing students and year levels
  var ss = SpreadsheetApp.openById("1X_SEaakzSSLtJbKg*");
  SpreadsheetApp.setActiveSpreadsheet(ss);
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var formURL = data[1][yr];
  var wk = data[0][10];          // the upcoming week #
  var scheduleTime = data[1][10];// time to schedule assignment for
  var assignees = [];
     
  // first row in sheet is year level; second row is URL for form; third row onwards is student usernames
  for (var i=2; i<data.length; i++) {            // iterate over usernames in this column
    var username = data[i][yr];
    
    // if this user is in students list, add to assignees array 
    for (var j=0; j<students.length; j++) {
      if (!username) { break; }      // unless the username is blank
      var emailAddress = students[j].profile.emailAddress;
      if (emailAddress.indexOf(username) != -1) {
        Logger.log('Found ' + username + " in year " + (yr+1) + " column, pushing id to assignees[]");
        assignees.push(students[j].profile.id);
        break;
      }
    }
  }
  // TODO: Empty the cell after grabbing the form URL
  return [assignees, wk, formURL, scheduleTime]; 
}

function listCourses() {
  var response = Classroom.Courses.list();
  var courses = response.courses;
  if (courses && courses.length > 0) {
    for (i = 0; i < courses.length; i++) {
      var course = courses[i];
      Logger.log('%s (%s)', course.name, course.id);
    }
  } else {
    Logger.log('No courses found.');
  }
}

function listStudents() {
//  var classId = 40492570395;  // test
 var classId = 50331317750;  // prod 2020
//  var classId = 27692831788; // comms 2019
  var response = Classroom.Courses.Students.list(classId);
  var students = response.students;
  const roster = [],
    options = {pageSize: 30};
  do {
    var search = Classroom.Courses.Students.list(classId, options);      // Get the next page of students for this course.
    if (search.students) {                                                   // Add this page's students to the local collection of students.
      Array.prototype.push.apply(roster, search.students);
    }
    options.pageToken = search.nextPageToken;                                // Update the page for the request
  } while (options.pageToken);
  Logger.log("Found %s students", roster.length);
  for (var y=0; y<roster.length; y++) {
    Logger.log(roster[y].profile.emailAddress + "/" + roster[y].profile.name.givenName); // + "/" + roster[y].profile.name.familyName); 
  }
  return roster;
}

function postCourseWork(formUrl, assignees, yr, week, scheduleTime) {
  var ClassSource =  {
    title: ("Number Knowledge Year " + yr + " - Week " + week),
    state: "DRAFT",
    materials: [
      {
        link:{
          url: formUrl,
          title: ("Number Knowledge Year " + yr + " - Week " + week)
        }
      }
    ],
    scheduledTime: scheduleTime,
    workType: "ASSIGNMENT",
    assigneeMode: "INDIVIDUAL_STUDENTS",
    individualStudentsOptions: {
      studentIds: assignees
    }
  };
  Logger.log("Scheduling " + ClassSource.title + " to " + assignees.length + " students for " + scheduleTime);
  Classroom.Courses.CourseWork.create(ClassSource, 55633737545)      // WARNING: Uncommenting this will post actual assignments
}