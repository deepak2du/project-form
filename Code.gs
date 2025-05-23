/**
 * Main function to handle HTTP POST requests.
 * Routes requests to specific handlers based on the 'action' parameter.
 * Handles both JSON and FormData (multipart/form-data) content types.
 *
 * @param {GoogleAppsScript.Events.DoPost} e The event object containing request parameters and post data.
 * @returns {GoogleAppsScript.Content.TextOutput} A JSON response indicating success or error.
 */
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let params = {};
  let action;

  console.log("--- doPost Started (V9 - Base64 Upload Attempt) ---");
  console.log("e.postData.type:", e.postData ? e.postData.type : "N/A");

  try {
    if (e.postData && e.postData.type === 'application/json') {
      params = JSON.parse(e.postData.contents);
      action = params.action;
      console.log("doPost: JSON request detected. Action:", action);
    } else {
      // For FormData (multipart/form-data which includes Base64 text fields)
      console.log("doPost: Multipart/form-data or www-form-urlencoded request detected.");
      console.log("e.parameters (RAW for FormData):", JSON.stringify(e.parameters));

      for (const key in e.parameters) {
        if (e.parameters.hasOwnProperty(key)) {
          const valueArray = e.parameters[key]; // Values are always arrays
          if (Array.isArray(valueArray) && valueArray.length > 0) {
            params[key] = valueArray[0]; // Take the first value for simple fields
          } else {
            params[key] = ""; // Or handle as needed if valueArray might be empty
          }
        }
      }
      action = params.action;
      console.log("doPost: Params from FormData:", JSON.stringify(params)); // Base64 will be very long if logged
      console.log("doPost: Derived action for routing:", action);
    }

    if (!action) {
      console.error("doPost: Action parameter is missing.");
      return ContentService.createTextOutput(JSON.stringify({ error: "Action parameter is missing." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let output;
    switch (action) {
      case 'add_meeting':
        output = handleAddMeeting(params, ss);
        break;
      case 'edit_meeting':
        output = handleEditMeeting(params, ss);
        break;
      case 'delete_meeting':
        output = handleDeleteMeeting(params, ss);
        break;
      case 'add_action':
        output = handleAddAction(params, ss);
        break;
      case 'edit_action':
        output = handleEditAction(params, ss);
        break;
      case 'delete_action':
        output = handleDeleteAction(params, ss);
        break;
      case 'add_status':
        output = handleAddStatus(params, ss);
        break;
      case 'edit_status':
        output = handleEditStatus(params, ss);
        break;
      case 'delete_status':
        output = handleDeleteStatus(params, ss);
        break;
      case 'upload_media_base64': // New action for Base64 upload
        output = handleMediaUploadBase64(params, ss);
        break;
      // case 'upload_media': // Old direct blob upload - keep commented or remove
      //   output = handleMediaFileUpload(params, ss); 
      //   break;
      default:
        console.error("doPost: Unknown action detected:", action);
        output = ContentService.createTextOutput(JSON.stringify({ error: `Unknown action: ${action}` }))
          .setMimeType(ContentService.MimeType.JSON);
    }
    return output;

  } catch (err) {
    console.error(`Critical Error in doPost for action '${action || "N/A"}': ${err.message} Stack: ${err.stack}`);
    return ContentService.createTextOutput(
      JSON.stringify({ error: `Server error processing action '${action || "N/A"}': ${err.message}` })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    console.log("--- doPost Finished ---");
  }
}


function handleAddMeeting(params, ss) {
  try {
    const sheet = ss.getSheetByName("Meetings") || ss.insertSheet("Meetings");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Meeting ID", "Meeting Date", "Zone", "District", "Cold Room",
        "Meeting Title", "Conducted By", "Attendees", "Meeting Agenda",
        "Meeting Discussion", "Photo URL"
      ]);
    }

    const prefix = "BCIEINM";
    const lastRow = sheet.getLastRow();
    let nextId = prefix + "001";

    if (lastRow >= 1) {
      const idColumnRange = sheet.getRange(2, 1, Math.max(1, lastRow - 1), 1);
      const rawIds = idColumnRange.getValues().flat().map(String);
      const nums = rawIds
        .filter(id => id && String(id).startsWith(prefix))
        .map(id => {
          const suffix = String(id).slice(prefix.length);
          return parseInt(suffix, 10) || 0;
        });

      const maxNum = nums.length > 0 ? Math.max(...nums) : 0;
      const nextNum = maxNum + 1;
      const padLength = 3;
      nextId = prefix + String(nextNum).padStart(padLength, "0");
    }

    const newRow = [
      nextId, params.meetingDate || "", params.zone || "", params.district || "", params.coldRoom || "",
      params.meetingTitle || "", params.conductedBy || "", params.attendees || "", params.meetingAgenda || "",
      params.meetingDiscussion || "", params.photoUrl || ""
    ];
    sheet.appendRow(newRow);

    return ContentService.createTextOutput(JSON.stringify({ message: "Meeting added successfully!", meetingId: nextId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error in handleAddMeeting: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error adding meeting: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleEditMeeting(params, ss) {
  try {
    const sheet = ss.getSheetByName("Meetings");
    if (!sheet) throw new Error("Meetings sheet not found.");
    const row = parseInt(params.row);
    if (isNaN(row) || row <= 1) throw new Error("Invalid or header row number for editing meeting.");

    const updateRowData = [
      params.meetingId || "", params.meetingDate || "", params.zone || "", params.district || "", params.coldRoom || "",
      params.meetingTitle || "", params.conductedBy || "", params.attendees || "", params.meetingAgenda || "",
      params.meetingDiscussion || "", params.photoUrl || ""
    ];
    sheet.getRange(row, 1, 1, updateRowData.length).setValues([updateRowData]);
    return ContentService.createTextOutput(JSON.stringify({ message: "Meeting updated successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error updating meeting: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error updating meeting: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteMeeting(params, ss) {
  try {
    const sheet = ss.getSheetByName("Meetings");
    if (!sheet) throw new Error("Meetings sheet not found.");
    const row = parseInt(params.row);
    if (isNaN(row) || row <= 1 || row > sheet.getLastRow()) {
      throw new Error("Invalid row number for deleting meeting or attempting to delete header.");
    }
    sheet.deleteRow(row);
    return ContentService.createTextOutput(JSON.stringify({ message: "Meeting deleted successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error deleting meeting: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error deleting meeting: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAddAction(params, ss) {
  try {
    const sheet = ss.getSheetByName("Action Items") || ss.insertSheet("Action Items");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Meeting ID", "Action Item", "Assigned To", "Deadline", "Status"]);
    }
    sheet.appendRow([
      params.meetingId || "", params.actionItem || "", params.assignedTo || "",
      params.deadline || "", params.status || ""
    ]);
    return ContentService.createTextOutput(JSON.stringify({ message: "Action item added successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error adding action item: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error adding action item: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleEditAction(params, ss) {
  try {
    const sheet = ss.getSheetByName("Action Items");
    if (!sheet) throw new Error("Action Items sheet not found.");
    const row = parseInt(params.row);
    if (isNaN(row) || row <= 1) throw new Error("Invalid or header row number for editing action item.");

    sheet.getRange(row, 1, 1, 5).setValues([[
      params.meetingId || "", params.actionItem || "", params.assignedTo || "",
      params.deadline || "", params.status || ""
    ]]);
    return ContentService.createTextOutput(JSON.stringify({ message: "Action item updated successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error updating action item: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error updating action item: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteAction(params, ss) {
  try {
    const sheet = ss.getSheetByName("Action Items");
    if (!sheet) throw new Error("Action Items sheet not found.");
    const row = parseInt(params.row);
    if (isNaN(row) || row <= 1 || row > sheet.getLastRow()) {
      throw new Error("Invalid row number for deleting action item or attempting to delete header.");
    }
    sheet.deleteRow(row);
    return ContentService.createTextOutput(JSON.stringify({ message: "Action item deleted successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error deleting action item: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error deleting action item: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAddStatus(params, ss) {
  try {
    const sheet = ss.getSheetByName("Weekly Status") || ss.insertSheet("Weekly Status");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Week", "Zone", "District", "Summary of This Week Activities", "Activities Planned for Next Week"]);
    }
    sheet.appendRow([
      params.weekInfo || "", params.zone || "", params.district || "",
      params.currentWeekUpdate || "", params.nextWeekUpdate || ""
    ]);
    return ContentService.createTextOutput(JSON.stringify({ message: "Weekly status added successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error adding weekly status: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error adding weekly status: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleEditStatus(params, ss) {
  try {
    const sheet = ss.getSheetByName("Weekly Status");
    if (!sheet) throw new Error("Weekly Status sheet not found.");
    const row = parseInt(params.row);
    if (isNaN(row) || row <= 1) throw new Error("Invalid or header row number for editing status.");

    sheet.getRange(row, 1, 1, 5).setValues([[
      params.weekInfo || "", params.zone || "", params.district || "",
      params.currentWeekUpdate || "", params.nextWeekUpdate || ""
    ]]);
    return ContentService.createTextOutput(JSON.stringify({ message: "Weekly status updated successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error updating weekly status: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error updating weekly status: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteStatus(params, ss) {
  try {
    const sheet = ss.getSheetByName("Weekly Status");
    if (!sheet) throw new Error("Weekly Status sheet not found.");
    const row = parseInt(params.row);
    if (isNaN(row) || row <= 1 || row > sheet.getLastRow()) {
      throw new Error("Invalid row number for deleting status or attempting to delete header.");
    }
    sheet.deleteRow(row);
    return ContentService.createTextOutput(JSON.stringify({ message: "Weekly status deleted successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error deleting weekly status: ${error.message} Stack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ error: `Error deleting weekly status: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles file uploads sent as Base64 encoded strings.
 * Expects params.dataBase64, params.fileName, params.mimeType, 
 * params.weekInfo, params.zone, params.district.
 */
function handleMediaUploadBase64(params, ss) {
  console.log("--- handleMediaUploadBase64 Started ---");
  // Avoid logging full Base64 string if too long for logs
  let loggableParams = Object.assign({}, params);
  if (loggableParams.dataBase64 && loggableParams.dataBase64.length > 100) {
    loggableParams.dataBase64 = loggableParams.dataBase64.substring(0, 100) + "... (truncated)";
  }
  console.log("Received params (Base64 potentially truncated for log):", JSON.stringify(loggableParams));

  try {
    const dataBase64 = params.dataBase64;
    const fileName = params.fileName;
    const mimeType = params.mimeType; // Make sure this is the actual MIME type e.g. "image/png"
    const weekInfo = params.weekInfo;
    const zone = params.zone;
    const district = params.district;

    if (!dataBase64 || !fileName || !mimeType) {
      let missing = [];
      if (!dataBase64) missing.push("dataBase64");
      if (!fileName) missing.push("fileName");
      if (!mimeType) missing.push("mimeType");
      console.warn("handleMediaUploadBase64: Missing required parameters:", missing.join(", "));
      return ContentService
        .createTextOutput(JSON.stringify({ error: "Upload failed: Missing required data (base64 string, filename, or mimetype). Missing: " + missing.join(", ") }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    console.log(`handleMediaUploadBase64: Processing file: ${fileName}, type: ${mimeType}, base64 length (approx): ${dataBase64.length}`);

    const decodedBytes = Utilities.base64Decode(dataBase64, Utilities.Charset.UTF_8); // Using UTF-8, ensure this is appropriate. For binary, it might not matter.
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);

    console.log("handleMediaUploadBase64: Blob created. Name:", blob.getName(), "Type:", blob.getContentType(), "Size:", blob.getBytes().length);

    let sheet = ss.getSheetByName('Media');
    if (!sheet) {
      sheet = ss.insertSheet('Media');
      sheet.appendRow(['Week', 'Zone', 'District', 'File Name', 'File URL', 'File Type', 'Uploaded On']);
    }

    const folderId = '1Yz4uaedg3Ako_KPP9FKqCuZq1kiGLo1p'; // Your Google Drive Folder ID
    const folder = DriveApp.getFolderById(folderId);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    sheet.appendRow([
      weekInfo || "",
      zone || "",
      district || "",
      file.getName(),
      file.getUrl(),
      file.getMimeType(),
      new Date()
    ]);

    console.log(`handleMediaUploadBase64: File ${file.getName()} uploaded and logged successfully.`);
    return ContentService
      .createTextOutput(JSON.stringify({ message: 'Upload successful (Base64).', fileName: file.getName(), fileUrl: file.getUrl() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error(`Error in handleMediaUploadBase64: ${error.message} Stack: ${error.stack}`);
    // Log the params again if an error occurs to see what was received
    console.error("Params at time of error in handleMediaUploadBase64 (Base64 potentially truncated):", JSON.stringify(loggableParams));
    return ContentService.createTextOutput(
      JSON.stringify({ error: `Upload failed (Base64) due to a server error: ${error.message}` })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    console.log("--- handleMediaUploadBase64 Finished ---");
  }
}

/*
// Old direct blob upload handler - commented out as we are trying Base64
function handleMediaFileUpload(params, ss) {
  // ... (previous implementation expecting params.mediaFile as a direct blob array) ...
}
*/

function doGet(e) {
  console.log("--- doGet Started ---");
  console.log("doGet: e.parameter:", JSON.stringify(e.parameter));
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = e.parameter.sheet;

    if (!sheetName) {
      console.error(`doGet error: Sheet name parameter is missing. Parameters: ${JSON.stringify(e.parameter)}`);
      return ContentService.createTextOutput(JSON.stringify({ error: "Sheet name parameter is missing." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ss.getSheetByName(sheetName);
    let data = [];

    if (sheet) {
      if (sheet.getLastRow() > 0) {
        data = sheet.getDataRange().getValues();
      } else {
        console.warn(`Sheet '${sheetName}' is empty. Returning empty data array (no headers).`);
      }
    } else {
      console.warn(`Sheet not found: '${sheetName}'. Returning empty data for client handling.`);
    }
    console.log(`doGet: Returning data for sheet '${sheetName}'. Number of rows (including potential header): ${data.length}`);
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(`Error in doGet: ${error.message} Stack: ${error.stack} Parameters: ${JSON.stringify(e.parameter)}`);
    return ContentService.createTextOutput(
      JSON.stringify({ error: `Error fetching data: ${error.message}`, sheet: e.parameter.sheet || "Unknown" })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    console.log("--- doGet Finished ---");
  }
}
