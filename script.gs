const SHEET_NAME = "Sheet1"
const HEADER_ROW = 1;

const GITHUB_TOKEN = ``;
const REPO_OWNER = 'caitlin-allison';
const REPO_NAME = 'SyncIssuesWithGoogleSheets';

const ROWS_HDR = {
  "Github_Id" : "A",
  "Type" : "B",
  "Title": "C",
  "Assignees": "D",
  "Labels": "E",
  "Project": "F",
  "Notes": "G",
  "Start_Date": "H",
  "End Date": "I",
  "Google_Sheets_Sync": "J",
  "Github_Sync": "K"
}
const HDR_INDEX = {
  "Github_Id" : 0,
  "Type" : 1,
  "Title": 2,
  "Assignees": 3,
  "Labels": 4,
  "Project": 5,
  "Notes": 6,
  "Start_Date": 7,
  "End Date": 8,
  "Google_Sheets_Sync": 9,
  "Github_Sync": 10
}

function handleEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var rowsData = getRowData(sheet, range.getRow())
  var row = range.getRow();

  // Dont look at row 1 (header) or another sheet
  if (sheet.getName() !== SHEET_NAME || row === HEADER_ROW) {
    return;
  }


  // Document date and time of edit for Google Sheets
  timeStampEdit(sheet, row, false);

  // Loop through arr of rows (even if 1, its the return format)
  for (const rowIndex in rowsData) {
    const rowValues = rowsData[rowIndex];

    // Is an issue, connect to Github
    if (rowValues[HDR_INDEX.Type].includes('Issue'))
    {
      var assignees = (rowValues[HDR_INDEX.Assignees] ?? '').split(", ").map(l => l.trim()).filter(l => l.length > 0);
      var labels = (rowValues[HDR_INDEX.Labels] ?? '').split(", ").map(l => l.trim()).filter(l => l.length > 0);

      // Is an existing Github Issue
      if (!!rowValues[HDR_INDEX.Github_Id]){
        updateGithubIssue(
            rowValues[HDR_INDEX.Github_Id], 
            rowValues[HDR_INDEX.Title], 
            assignees,
            labels,
            rowValues[HDR_INDEX.Project] ?? null,
            rowValues[HDR_INDEX.Notes]
          )
      }
      else {
        var assignees = (rowValues[HDR_INDEX.Assignees] ?? '').split(", ").map(l => l.trim()).filter(l => l.length > 0);
        var labels = (rowValues[HDR_INDEX.Labels] ?? '').split(", ").map(l => l.trim()).filter(l => l.length > 0);

        let issueId = createGitHubIssue(
            rowValues[HDR_INDEX.Title], 
            assignees, labels, 
            rowValues[HDR_INDEX.Project] ?? null, 
            rowValues[HDR_INDEX.Notes]
        );
        if (issueId) {
          const issueCell = sheet.getRange(ROWS_HDR.Github_Id + row.toString());
          issueCell.setValue(issueId);
        }

      }
    }
  }

}


function getRowData(sheet, rowNumber) {
  const range = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()); // Get the entire row
  const rowData = range.getValues(); // Get the values as a 2D array
  return rowData;
}

// Insert date time into specificied cell (Google Sheets Sync or Github Sync)
function timeStampEdit(sheet, row, isFromGithub){
  // Timestap either the Google Sheets or Github Sync column
  var dateModifiedColumnLetter = isFromGithub? ROWS_HDR.Github_Sync : ROWS_HDR.Google_Sheets_Sync;

  var time = new Date();
  time = Utilities.formatDate(time, "GMT-05:00", "MM/dd/yy, hh:mm:ss");

  var dateModifiedRange = sheet.getRange(dateModifiedColumnLetter + row.toString());
  dateModifiedRange.setValue(time);
}


/**
 * @id : github id, number
 * @title: string
 * @assignees: string[], (may need to replace here with their actual @)
 * @labels: string[]
 * @project: string,
 * @body: string ver of markdown
 */
function updateGithubIssue(id, title, assignees, labels, project, body){
 const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/issues/${id}`;
  const payload = JSON.stringify({ 
    title, 
    body,
    assignees, 
    labels, 
    project
  });

  const options = {
    method: "patch",
    contentType: "application/json",
    headers: {
      Authorization: `token ${GITHUB_TOKEN}`
    },
    payload: payload,
    muteHttpExceptions: true
  };

  console.log("Updating issue on Github")
  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText())
  console.log("Update a success!")


}

/**
 * @title: string
 * @assignees: string, (may need to replace here with their actual @)
 * @labels: string[]
 * @project: string,
 * @body: string ver of markdown
 */
function createGitHubIssue(title, assignees, labels, project, body) {
  const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/issues`;
  const payload = JSON.stringify({ 
    title, 
    body,
    assignees, 
    labels, 
    project,  
    });

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `token ${GITHUB_TOKEN}`
    },
    payload: payload,
    muteHttpExceptions: true
  };

  console.log("Creating github issue");
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() === 201 && result.number) {
    console.log("Creation success! Id:", result.number);
    return result.number; // Return new issue number
  } else {
    Logger.log("Issue creation failed: " + response.getContentText());
    return null;
  }

}
