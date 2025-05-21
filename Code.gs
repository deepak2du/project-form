/**
 * Utility to add CORS headers to any TextOutput
 */
function withCors(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * Handle CORS preflight requests.
 */
function doOptions(e) {
  return withCors(
    ContentService.createTextOutput('')
      .setMimeType(ContentService.MimeType.JSON)
  );
}

/**
 * Handle POST requests for all actions, with relaxed JSON detection and CORS.
 */
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data, action;

  // 1. Detect JSON vs. form-encoded
  if (e.postData && e.postData.type && e.postData.type.indexOf('application/json') === 0) {
    data = JSON.parse(e.postData.contents);
    action = data.action;
  } else {
    data = e.parameter;
    action = data.action;
  }

  // 2. Dispatch to the correct handler
  var output;
  try {
    switch (action) {
      case 'add_meeting':
        output = handleAddMeeting({ parameter: data }, ss);
        break;
      case 'edit_meeting':
        output = handleEditMeeting({ parameter: data }, ss);
        break;
      case 'delete_meeting':
        output = handleDeleteMeeting({ parameter: data }, ss);
        break;
      case 'add_action':
        output = handleAddAction({ parameter: data }, ss);
        break;
      case 'edit_action':
        output = handleEditAction({ parameter: data }, ss);
        break;
      case 'delete_action':
        output = handleDeleteAction({ parameter: data }, ss);
        break;
      case 'add_status':
        output = handleAddStatus({ parameter: data }, ss);
        break;
      case 'edit_status':
        output = handleEditStatus({ parameter: data }, ss);
        break;
      case 'delete_status':
        output = handleDeleteStatus({ parameter: data }, ss);
        break;
      case 'upload_media':
        output = handleUploadMediaJson(data, ss);
        break;
      default:
        output = ContentService
          .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
          .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    output = ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 3. Always wrap with CORS headers
  return withCors(output);
}

/**
 * Optional: Handle CORS preflight requests.
 */
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * Add a new meeting, auto‐incrementing Meeting ID with prefix "BCIEINM"
 */
function handleAddMeeting(e, ss) {
  try {
    // 1. Ensure the “Meetings” sheet exists and has headers
    var sheet = ss.getSheetByName("Meetings") || ss.insertSheet("Meetings");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Meeting ID",
        "Meeting Date",
        "Zone",
        "District",
        "Cold Room",
        "Meeting Title",
        "Conducted By",
        "Attendees",
        "Meeting Agenda",
        "Meeting Discussion",
        "Photo URL"
      ]);
    }

    // 2. Load existing IDs and extract numeric suffixes
    var headerRows = 1;
    var lastRow = sheet.getLastRow();
    var idRange = sheet.getRange(headerRows + 1, 1, Math.max(0, lastRow - headerRows), 1);
    var rawIds = idRange.getValues().flat().map(String);
    var prefix = "BCIEINM";
    var nums = rawIds
      .filter(id => id.startsWith(prefix))
      .map(id => {
        var suffix = id.slice(prefix.length);
        return parseInt(suffix, 10) || 0;
      });

    // 3. Compute next numeric suffix
    var maxNum = nums.length ? Math.max.apply(null, nums) : 0;
    var nextNum = maxNum + 1;
    // Pad to same length as the longest existing suffix (or 3 digits minimum)
    var padLength = Math.max(3, Math.max(...nums.map(n => String(n).length)));
    var nextId = prefix + String(nextNum).padStart(padLength, "0");

    // 4. Build the new row
    var data = e.parameter;
    var newRow = [
      nextId,
      data.meetingDate || "",
      data.zone || "",
      data.district || "",
      data.coldRoom || "",
      data.meetingTitle || "",
      data.conductedBy || "",
      data.attendees || "",
      data.meetingAgenda || "",
      data.meetingDiscussion || "",
      data.photoUrl || ""
    ];

    // 5. Append to sheet
    sheet.appendRow(newRow);

    // 6. Return success + the generated Meeting ID
    return ContentService
      .createTextOutput(JSON.stringify({
        message: "Meeting added successfully!",
        meetingId: nextId
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error in handleAddMeeting:", error);
    return ContentService
      .createTextOutput(JSON.stringify({
        error: "Error adding meeting: " + error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function handleEditMeeting(e, ss) {
  try {
    const sheet = ss.getSheetByName("Meetings");
    const data = e.parameter;
    const row = parseInt(data.row);
    const meetingDate = data.meetingDate;

    const updateRow = [
      data.meetingId,
      meetingDate,
      data.zone,
      data.district,
      data.coldRoom,
      data.meetingTitle,
      data.conductedBy,
      data.attendees,
      data.meetingAgenda,
      data.meetingDiscussion,
      data.photoUrl || ""
    ];

    sheet.getRange(row, 1, 1, updateRow.length).setValues([
      updateRow
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Meeting updated successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error updating meeting: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error updating meeting: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteMeeting(e, ss) {
  try {
    const sheet = ss.getSheetByName("Meetings");
    const row = parseInt(e.parameter.row);
    sheet.deleteRow(row);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Meeting deleted successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error deleting meeting: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error deleting meeting: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAddAction(e, ss) {
  try {
    const sheet = ss.getSheetByName("Action Items") || ss.insertSheet("Action Items");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Meeting ID", "Action Item", "Assigned To", "Deadline", "Status"]);
    }

    const data = e.parameter;
    sheet.appendRow([
      data.meetingId,
      data.actionItem,
      data.assignedTo,
      data.deadline,
      data.status
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Action item added successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error adding action item: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error adding action item: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleEditAction(e, ss) {
  try {
    const sheet = ss.getSheetByName("Action Items");
    const data = e.parameter;
    const row = parseInt(data.row);

    sheet.getRange(row, 1, 1, 5).setValues([
      [
        data.meetingId,
        data.actionItem,
        data.assignedTo,
        data.deadline,
        data.status
      ]
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Action item updated successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error updating action item: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error updating action item: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteAction(e, ss) {
  try {
    const sheet = ss.getSheetByName("Action Items");
    const row = parseInt(e.parameter.row);
    sheet.deleteRow(row);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Action item deleted successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error deleting action item: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error deleting action item: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAddStatus(e, ss) {
  try {
    const sheet = ss.getSheetByName("Weekly Status") || ss.insertSheet("Weekly Status");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Week", "Zone", "District", "Summary of This Week Activities", "Activities Planned for Next Week"]);
    }

    const data = e.parameter;
    sheet.appendRow([
      data.weekInfo,
      data.zone,
      data.district,
      data.currentWeekUpdate,
      data.nextWeekUpdate
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Weekly status added successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error adding weekly status: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error adding weekly status: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleEditStatus(e, ss) {
  try {
    const sheet = ss.getSheetByName("Weekly Status");
    const data = e.parameter;
    const row = parseInt(data.row);

    sheet.getRange(row, 1, 1, 5).setValues([
      [
        data.weekInfo,
        data.zone,
        data.district,
        data.currentWeekUpdate,
        data.nextWeekUpdate
      ]
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Weekly status updated successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error updating weekly status: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error updating weekly status: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteStatus(e, ss) {
  try {
    const sheet = ss.getSheetByName("Weekly Status");
    const row = parseInt(e.parameter.row);
    sheet.deleteRow(row);

    return ContentService.createTextOutput(
      JSON.stringify({ message: "Weekly status deleted successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error deleting weekly status: ", error);
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Error deleting weekly status: " + error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Decode a Base64 JSON upload and store it in Drive + your “Media” sheet.
 */
function handleUploadMediaJson(data, ss) {
  var sheet = ss.getSheetByName('Media') || ss.insertSheet('Media');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Week', 'Zone', 'District', 'File Name', 'File URL', 'File Type', 'Uploaded On'
    ]);
  }

  var folder = DriveApp.getFolderById('1Yz4uaedg3Ako_KPP9FKqCuZq1kiGLo1p');
  var b64 = data.dataBase64;
  var name = data.fileName;
  var mime = data.mimeType;
  var weekInfo = data.weekInfo;
  var zone = data.zone;
  var district = data.district;

  if (!b64 || !name || !mime) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Missing file data.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // decode + upload
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // log into sheet
  sheet.appendRow([
    weekInfo,
    zone,
    district,
    file.getName(),
    file.getUrl(),
    file.getMimeType(),
    new Date()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ message: 'Upload successful.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Serve sheet data for your gallery and tables.
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = e.parameter.sheet;
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet ? sheet.getDataRange().getValues() : [];

  return withCors(
    ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON)
  );
}

/**
 * Optional: Handle OPTIONS preflight for CORS
 */
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}


function generateNextMeetingID(sheet) {
  const prefix = "BCIEINM";
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + "001";

  const data = sheet.getRange(2, 1, lastRow - 1).getValues().reverse();
  for (let i = 0; i < data.length; i++) {
    const id = data[i][0];
    if (id && typeof id === 'string' && id.startsWith(prefix)) {
      const num = parseInt(id.replace(prefix, ""));
      if (!isNaN(num)) {
        return prefix + (num + 1).toString().padStart(3, '0');
      }
    }
  }
  return prefix + "001";
}