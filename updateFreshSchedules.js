function updateFreshSchedules() {
  try {
    const TEMP_FOLDER_ID = "YOUR_TEMP_FOLDER_ID_HERE";
    const SUBJECT_LINE = "YOUR_EMAIL_SUBJECT_HERE";
    const TARGET_SHEET = "FreshSchedules";

    const CONDITION_MATCH = [
      "Condition Type 1",
      "Condition Type 2",
      "Condition Type 3",
      "Condition Type 4",
      "Condition Type 5"
    ];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TARGET_SHEET);
    if (!sheet) sheet = ss.insertSheet(TARGET_SHEET);

    const threads = GmailApp.search(`subject:"${SUBJECT_LINE}"`);
    if (!threads.length) throw new Error("No email found with the provided subject.");

    const messages = threads[0].getMessages();
    const lastMessage = messages[messages.length - 1];
    const attachments = lastMessage.getAttachments();
    if (!attachments.length) throw new Error("No attachment found in the email.");

    const folder = DriveApp.getFolderById(TEMP_FOLDER_ID);
    while (folder.getFiles().hasNext()) folder.getFiles().next().setTrashed(true);

    const zipFile = folder.createFile(attachments[0]);
    const unzipped = Utilities.unzip(zipFile.getBlob());
    if (!unzipped.length) throw new Error("No files inside ZIP.");

    const csvBlob = unzipped.find(f => f.getName().toLowerCase().endsWith(".csv")) || unzipped[0];
    const csvContent = csvBlob.getDataAsString("UTF-8");
    const csvData = Utilities.parseCsv(csvContent);
    if (!csvData.length) throw new Error("CSV is empty.");

    const headers = csvData[0];
    const rows = csvData.slice(1);

    const filtered = rows.filter(r => CONDITION_MATCH.includes(String(r[9]).trim()));

    const map = {};
    for (const r of filtered) {
      const id = String(r[2] || "").trim();
      const name = String(r[5] || "").trim();
      if (!id || !name) continue;

      const key = id + "|" + name;
      if (!map[key]) map[key] = r;
    }

    const deduped = Object.values(map);
    const today = new Date();
    const todayFiltered = deduped.filter(r => {
      const dt = parseDateFlexible(r[12]);
      return (
        dt &&
        dt.getFullYear() === today.getFullYear() &&
        dt.getMonth() === today.getMonth() &&
        dt.getDate() === today.getDate()
      );
    });

    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (todayFiltered.length > 0) {
      sheet.getRange(2, 1, todayFiltered.length, headers.length).setValues(todayFiltered);
    }

    zipFile.setTrashed(true);
    while (folder.getFiles().hasNext()) folder.getFiles().next().setTrashed(true);

  } catch (err) {
    Logger.log("Error: " + err.message);
  }
}

function parseDateFlexible(input) {
  if (!input) return null;
  if (input instanceof Date && !isNaN(input)) return input;

  const s = String(input).trim();

  let m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) return new Date(m[3], m[2] - 1, m[1], m[4] || 0, m[5] || 0, m[6] || 0);

  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) return new Date(m[1], m[2] - 1, m[3], m[4] || 0, m[5] || 0, m[6] || 0);

  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}
