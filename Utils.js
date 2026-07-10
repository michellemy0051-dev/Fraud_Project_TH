
function formatDateForBQ(date) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, "yyyy-MM-dd");
}

// return group map
function getMemberGroupMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Member_Group_ID");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  const map = {};

  data.forEach(row => {
    const servers = String(row[0]).trim();
    const groupId = String(row[1]).trim();
    const description = String(row[2]).trim();

    if (servers && groupId && description) {
      // Combine server + description
      map[groupId] = `${servers} ${description}`;
    }
  });

  return map;
}


function generateSqlExcludedGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Member_Group_ID");
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return "/* No data found */";

  // Get only the GroupID column (Column B / Index 1)
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  
  // Filter out blanks and wrap each ID in double quotes
  const formattedIds = data
    .map(row => String(row[0]).trim())
    .filter(id => id.length > 0)
    .map(id => `    "${id}"`)
    .join(",\n");

  // Construct the full SQL block
  const sqlSnippet = `
  SELECT groupid
  FROM UNNEST([
  ${formattedIds}
    ]) AS groupid
    `;

  // Logger.log(sqlSnippet);
  return sqlSnippet;
}


//Waits for job completion
function runQueryStream(sql, onRow) {
  const request = {
    query: sql,
    useLegacySql: false
  };

  Logger.log("➡️ Starting streaming query job...");
  const job = BigQuery.Jobs.query(request, EXECUTION_PROJECT_ID);
  const jobId = job.jobReference.jobId;
  const location = job.jobReference.location;

  Logger.log("   Job ID: " + jobId);

  // Wait for job to complete FIRST
  let results = BigQuery.Jobs.getQueryResults(EXECUTION_PROJECT_ID, jobId, {
    maxResults: 3000,
    timeoutMs: 60000,
    location: location
  });

  while (!results.jobComplete) {
    Logger.log("⏳ Waiting for streaming job to complete...");
    Utilities.sleep(1000);
    results = BigQuery.Jobs.getQueryResults(EXECUTION_PROJECT_ID, jobId, {
      maxResults: 3000,
      timeoutMs: 60000,
      location: location
    });
  }

  Logger.log("🎯 Streaming job complete. Processing rows...");

  // Process first page
  let rowCount = 0;
  (results.rows || []).forEach(r => {
    rowCount++;
    onRow(r.f.map(c => c.v));
  });

  if (rowCount > 0) {
    Logger.log(`   Processed ${rowCount} rows from page 1`);
  }

  // Process remaining pages
  let pageToken = results.pageToken;
  let page = 1;

  while (pageToken) {
    page++;
    const res = BigQuery.Jobs.getQueryResults(EXECUTION_PROJECT_ID, jobId, {
      maxResults: 3000,
      pageToken: pageToken,
      location: location,
      timeoutMs: 60000
    });

    let pageRows = 0;
    (res.rows || []).forEach(r => {
      rowCount++;
      pageRows++;
      onRow(r.f.map(c => c.v));
    });

    if (pageRows > 0) {
      Logger.log(`   Processed ${pageRows} rows from page ${page} (total: ${rowCount})`);
    }

    pageToken = res.pageToken;
  }

  Logger.log(`✅ Stream complete. Total rows processed: ${rowCount}`);
}

/**
 * Applies vertical borders to specific column indices within a range.
 */
function applyThinBorders(range, indices) {
  const rowCount = range.getHeight();
  const colCount = range.getWidth();
  const borderColor = "#b7b7b7";
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID;

  // Clear existing vertical borders first (Optional, for a fresh look)
  range.setBorder(null, false, null, false, false, false, borderColor, borderStyle);

  // Set Left Border for the very first column
  range.offset(0, 0, rowCount, 1)
       .setBorder(null, true, null, null, null, null, borderColor, borderStyle);

  // Set Right Borders for specified indices
  indices.forEach(idx => {
    if (idx > 0 && idx <= colCount) {
      range.offset(0, idx - 1, rowCount, 1)
           .setBorder(null, null, null, true, null, null, borderColor, borderStyle);
    }
  });
}