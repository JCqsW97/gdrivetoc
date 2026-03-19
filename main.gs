// ======================================================================================================
/*
 * 
 * Google Apps Script - List all the folders and files in a Google Drive folder, and enter the information into a spreadsheet.
 * How it works: 
 *     - In the “Table of Contents” tab, click "Create a table of contents" and enter the folder ID in the pop-up window.
 *       You can copy the folder ID from your browser's address bar;
 *       it consists of a sequence of letters and numbers following the ‘folders/’ portion of the URL.
 *       The folder ID is likely to be everything after the ‘folders/’ portion of the URL.
 * 
 * @Source: https://www.youtube.com/watch?v=PNrkjHkgZwU. (last checked on 20/10/21)
 * @Modified by JcqsW97
 * 
 * @version 5 (updated 26/03/19)
 *
 */
// ======================================================================================================

function onOpen() 
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Table of Contents')
    .addItem('Create a table of contents', 'listAll')
    .addSeparator()
    .addItem('Add a field (Optional)', 'showFieldDialog')
    .addToUi();
  
  var ui2 = SpreadsheetApp.getUi();
  ui2.alert("Hi!", "Welcome to this Table of Contents spreadsheet!"+"\n Enjoy!", Browser.Buttons.OK);
}

function listAll() 
{
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var lastRow = currentSheet.getLastRow();
  
  try{
    var temp = currentSheet.getRange(lastRow,1).getValue();
  }
  
  catch (e) {
    var temp = '0';
  }
  
  var marker = 'All files listed on: ';
  
  if (temp.indexOf(marker) > -1 ) 
  {
    var userInput = Browser.msgBox('Working...  - Question:', 'It looks like: ' + temp + '. \n "YES" to create a new table of contents, "NO" to cancel and keep the one existing.', Browser.Buttons.YES_NO);
    if (userInput == "no") 
    {
      return;
    }
  }
  else 
  {
    var userInput = Browser.msgBox('Working... - Question:', 'Create a new table of contents OR use the existing one? "YES" to create a new table of contents, "NO" to use the existing one.', Browser.Buttons.YES_NO);
  }

  if (userInput == "yes") 
  {
    var start = new Date();
    currentSheet.clear();
    var folderId = Browser.inputBox('Enter a google folder ID (a list of letters and numbers after the "/folders/" inside the URL):', Browser.Buttons.OK_CANCEL);
    
    currentSheet.appendRow(["Gdrive ID", "Name", "Path", "URL", "Folder ID (only for folder)"]);
    
    currentSheet.setColumnWidth(1,65);
    currentSheet.setColumnWidth(2,200);
    currentSheet.setColumnWidth(3,200);
    currentSheet.setColumnWidth(4,200);
    currentSheet.setColumnWidth(5,260);
    
    currentSheet.getRange(1,1,1,5).setBackground('#ffffaa').setBorder(true,true,true,true,true,true).setFontWeight('bold');
    
    currentSheet.setFrozenRows(1);
    
    var list = [];
    var excluded = [];
    
    if (folderId === "") 
    {
      Browser.msgBox('Not a valid folder ID!');
      return;
    }
  } 
  else if (userInput == "no") 
  {
    temp = currentSheet.getRange(1, 5, lastRow).getValues();
    var folderId = Browser.inputBox('Enter a google folder ID (a list of letters and numbers after the "/folders/" inside the URL):', Browser.Buttons.OK_CANCEL);
    var list = cleanArray([].concat.apply([], temp));
    var lastID = list.pop();
    var excluded = [lastID,];
    getVoidFolderList(lastID, list, excluded);
  }
  else 
  {
    return;
  }


  var parent = DriveApp.getFolderById(folderId);
  var parentName = DriveApp.getFolderById(folderId).getName();
  var options = {weekday: "long", year: "numeric", month: "long", day: "numeric", hour: "numeric", minute: "2-digit", second: "numeric"};
  getChildFolders(parentName, parent, currentSheet, list, excluded);
  getRootFiles(parentName, parent, currentSheet);
  
  //set a marker once completed, and the duration of the execution
  var duration = (new Date().getTime()-start.getTime())/1000;
  var durationString = Utilities.formatString("%01.1f", duration)
  SpreadsheetApp.setActiveSheet(currentSheet).getRange(currentSheet.getLastRow()+1,1).setValue(marker + new Date().toLocaleDateString('en-EN',options)+ "\n Enjoy!").setFontColor("#000097");
  SpreadsheetApp.setActiveSheet(currentSheet).getRange(currentSheet.getLastRow()+1,1).setValue("Time: "+durationString+' Seconds').setFontColor('grey').setFontSize(9);
  SpreadsheetApp.flush();
  Browser.msgBox("Done!", marker + new Date().toLocaleDateString('en-EN',options), Browser.Buttons.OK);

}
                    

// ======================================================================================================
// Will remove all falsy values: undefined, null, 0, false, NaN and "" (empty string)
                    
function cleanArray(actual) 
{ 
  var newArray = new Array();
  
  for (var i = 0; i < actual.length; i++) 
  {
    if (actual[i]) 
    {
      newArray.push(actual[i]);
    }
  }
  return newArray;
}

//=====================================================================================================
// ----------code below is function to get root file names.

function getRootFiles(parentName, parent, sheet) 
{
  var files = parent.getFiles();
  var output = [];
  var path;
  var Url;
  var fileID;
  var FileOwnerEmail;
  
  while (files.hasNext()) 
  {
    var childFile = files.next();
    var fileName = childFile.getName();
    path = parentName + " |--> " + fileName;
    fileID = childFile.getId();
    Url = "https://drive.google.com/open?id=" + fileID;
    // ---------- Write
    output.push([fileID, fileName, path, Url]);
  }
  
  if (output.length) 
  {
    var last_row = sheet.getLastRow();
    sheet.getRange(last_row + 1, 1, output.length, 4).setValues(output);
  }
}

// ======================================================================================================

function getChildFolders(parentName, parent, sheet, voidFolder, excluded) 
{
  var childFolders = parent.getFolders();

  // ---------- List folders inside the folder
  while (childFolders.hasNext()) 
  {
    var childFolder = childFolders.next();
    var folderID = childFolder.getId();

    if (voidFolder.indexOf(folderID) > -1) 
    { // the folder and files in it has been listed
      continue;
    }

    var folderName = childFolder.getName();

    var data = [
      folderID,
      folderName,
      parentName + " |--> " + folderName,
      //     childFolder.getDateCreated(), // ----------cut out to save time
      ("https://drive.google.com/open?id=" + folderID),
      //     childFolder.getLastUpdated(), // ----------cut out to save time
      //     childFolder.getDescription(), // ----------cut out to save time
      //     childFolder.getSize(),  // ----------cut out to save time
      //     childFolder.getOwner().getEmail(),
      folderID, // ----------indicate this is a folder 
    ];
    
    // ---------- Write
    if (excluded.indexOf(folderID) == -1) 
    { //check the situation of folder is listed but the files have not;
      sheet.appendRow(data);
    }

    // ---------- List files inside the folder
    var files = childFolder.getFiles();
    var output = [];
    var path;
    var Url;
    var fileID;
    var FileOwnerEmail;
    while (files.hasNext()) 
    {
      var childFile = files.next();
      var fileName = childFile.getName();
      path = parentName + " |--> " + folderName + " |--> " + fileName;
      //       fileName,
      //       childFile.getDateCreated(),
      //       childFile.getLastUpdated(),
      //       childFile.getDescription(),
      //       childFile.getSize(),
      //     FileOwnerEmail=childFile.getOwner().getEmail();
      //        childFile.getParents().getUrl()
      fileID = childFile.getId();
      Url = "https://drive.google.com/open?id=" + fileID;
      // ---------- Write
      output.push([fileID, fileName, path, Url]);
    }
    
    if (output.length) 
    {
      var last_row = sheet.getLastRow();
      sheet.getRange(last_row + 1, 1, output.length, 4).setValues(output);
    }

    // ---------- Recursive call of the subfolder
    getChildFolders(parentName + " |--> " + folderName, childFolder, sheet, voidFolder, excluded);

  }
}

// ======================================================================================================
// ---------- Code below is to exclude the parent ID where the script stops so that it will catch the folder from beginning. 


function getVoidFolderList(lastID, list, excluded) {

  var folderParents = DriveApp.getFolderById(lastID).getParents();
  var folderID = new Array;

  while (folderParents.hasNext()) {
    var folderID = folderParents.next().getId();
    if (folderID == "NOT FOUND") {
      break;
    }

    var index = list.indexOf(folderID);
    if (index !== -1) {
      excluded.push(folderID);
      list.splice(index, 1);
    }

    getVoidFolderList(folderID, list, excluded);
  }
}

// ======================================================================================================
// This function call the HTML file associated that allows field selection from the user
function showFieldDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FieldDialog')
    .setWidth(800)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Field');
}

// ======================================================================================================
// This function allows the user to add field to the table of contents
function List(fieldName)
{
  var message = 'You selected:\n \n' + fieldName;
  SpreadsheetApp.getUi().alert(message);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastrow = ss.getLastRow();
  var lastcolumn = ss.getLastColumn();

  if (lastrow > 100)
  {
    var startnumber = Browser.inputBox("There are more than 100 rows. Please specify the row number to start listing the field " + fieldName +'.\\n'+'\\nP.S. Do not write "1", the header does not have a field ' + fieldName + "... \\n It starts at 2.", Browser.Buttons.OK_CANCEL);

    if (startnumber == "cancel") {return;}
    
    if (startnumber == 1) 
    {
        Browser.msgBox("Please pay attention, 1 is the header, try again! \\n ", Browser.Buttons.OK);
        return;
    }
    var x = Number(startnumber);
    if (isNaN(x))
    {
      Browser.msgBox('Nice try, however this script works better with integers!\\n ', Browser.Buttons.OK);
      return;
    }
  }
  else
  {
    var startnumber = 2;
  }

  var actualHeaders = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues().flat();

  if (!actualHeaders.includes(fieldName))
  {
    var fieldColumn = lastcolumn + 1;
  }
  else
  {
    var fieldColumn = actualHeaders.indexOf(fieldName) + 1;
  }

  ss.getRange(1, fieldColumn).setValue(fieldName).setBackground('#ffffaa').setBorder(true,true,true,true,true,true).setFontWeight('bold');

  for (var x = startnumber; x < lastrow - 1; x++)
  {
    var value = ss.getRange(x, 1).getValue();
  
    if (value == "") 
    {
        continue;
    }
  
    try 
    {
        var childFolder = DriveApp.getFolderById(value);
    } catch (e) 
    {
        ss.getRange(x, fieldColumn).setValue("Fail");
        continue;
    }

    switch(fieldName) 
    //https://developers.google.com/apps-script/reference/drive/access
    {
    case 'Owner':
      try 
      {
          var owner = childFolder.getOwner().getEmail();
          ss.getRange(x, fieldColumn).setValue(owner);
      } catch (e) 
      {
          ss.getRange(x, fieldColumn).setValue("team drive");
      }
    break;

    case 'Editor':
      var editors = childFolder.getEditors();
      var editorValues = [];

      for (var v = 0; v < editors.length; v++) 
      {
        editorValues.push(editors[v].getEmail());
      }

      var joinValues = editorValues.join(", ");
      ss.getRange(x, fieldColumn).setValue(joinValues);


      if (joinValues == "") 
      {
        ss.getRange(x, fieldColumn).setValue("-");
      }
    break;

    case 'Viewer' :
      var viewers = childFolder.getViewers();
      var viewerValues = [];

      for (v = 0; v < viewers.length; v++) {
        viewerValues.push(viewers[v].getEmail());

      }

      var joinValues = viewerValues.join(", ");
      ss.getRange(x, fieldColumn).setValue(joinValues);

      if (joinValues == "") {
        ss.getRange(x, fieldColumn).setValue("-");
      }
    break;

    case 'Description' :
      var Description = childFolder.getDescription();
      ss.getRange(x, fieldColumn).setValue(Description);
      if (Description == "" || Description == "undefined" || Description == null) {
        ss.getRange(x, fieldColumn).setValue("-");
      }
    break;

    case 'Date Created' :
      var DateCreated = childFolder.getDateCreated();
      ss.getRange(x, fieldColumn).setValue(DateCreated);
    break;

    case 'Last Updated' :
      var LastUpdated = childFolder.getLastUpdated();
      ss.getRange(x, fieldColumn).setValue(LastUpdated);
    break;

    case 'Size' :
      var Size = childFolder.getSize();
      ss.getRange(x, fieldColumn).setValue(Size);

      if (Size == "0") {
        ss.getRange(x, fieldColumn).setValue("-");
      }
    break;

    case 'Access' :
      try {
        var sharingAccess = childFolder.getSharingAccess();
        ss.getRange(x, fieldColumn).setValue(sharingAccess);
      } catch (e) {
        ss.getRange(x, fieldColumn).setValue("Fail");
        continue;
      }
    break;

    case 'Permission' :
      try {
        var sharingPermission = childFolder.getSharingPermission();
        ss.getRange(x, fieldColumn).setValue(sharingPermission);
      } catch (e) {
        ss.getRange(x, fieldColumn).setValue("Fail");
        continue;
      }
    break;

    default:
      // Code to execute if expression doesn’t match any case
      SpreadsheetApp.getUi().alert("There is an error");
      return
    }
  }
}