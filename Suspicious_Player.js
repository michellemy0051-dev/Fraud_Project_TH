function insert_hardcoded_categorised_Data_phone_email() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("Suspicious_Player");
  const rawdatasheet = ss.getSheetByName("RT_Raw_Data_v4");
  const resultSheet = ss.getSheetByName("User_Result"); // <-- Added User_Result Sheet
  const statusCell = reportSheet.getRange("C6");
  const userInfoStartRow = 21 + 4;

  try {
    const toDate = reportSheet.getRange("B4").getValue();
    const toDateTime = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate()).getTime();

    const lastRow = rawdatasheet.getLastRow();
    const lastCol = rawdatasheet.getLastColumn();
    if (lastRow < 3) { statusCell.setValue("❌ No raw data found"); return; }

    statusCell.setValue("⏳ Fetching raw data...");
    const rawHeaders = rawdatasheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const rawData = rawdatasheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    const rawColMap = {};
    rawHeaders.forEach((h, i) => rawColMap[h.trim()] = i);
    const columnsToShow = [
      "category","phone/email", "earliestRegisteredDate", "registerRatePerDay", "brandCount","dpTries","successDpCnt", "dpAmt", "avgDpAmt", "maxDp", 
      "dpChannelCnt", "dpTimeoutRatio","dpSuccessRatio", "successWdCnt", "wdAmt", "avgSessionTime", "quickWDRatio","promoAmt", "Promo/Deposit", 
      "dpFreq", "wdFreq", "promoFreq", "WD x DP", "dailyBetCnt", "ValidBet", "earnAmt",
      "ValidBet/ dpAmt", "GGR", "NGR", "Win/Loss(%)", "activeBrandList", "activeUsernameList", "brand_list","usernameList", "status", "behavior_category"
    ];

    // --- NEW: Number Format Mapping based on your original request ---
    const formatMap = {
      "registerRatePerDay":"0.00","brandCount": "0","successDpCnt": "0", "dpAmt": "#,##0.00", "avgDpAmt": "#,##0.00", "maxDp": "#,##0.00",
      "dpChannelCnt": "0", "dpTimeoutRatio": "0.00%", "dpSuccessRatio": "0.00%", "successWdCnt": "0", "wdAmt": "#,##0.00",
      "avgSessionTime": "0.00", "quickWDRatio": "0.00%",
      "promoAmt": "#,##0.00", "Promo/Deposit": "0.00%", "dpFreq": "0.00", "wdFreq": "0.00",
      "promoFreq": "0.00", "WD x DP": "0.00", "dailyBetCnt": "0.00", "ValidBet": "#,##0.00",
      "earnAmt": "#,##0.00", "ValidBet/ dpAmt": "0.00", "GGR": "#,##0.00", "NGR": "#,##0.00",
      "Win/Loss(%)": "0.00%"
    };

    const renameMap = {
      "earliestRegisteredDate": "earliest_registered_date", "successDpCnt": "successDp", "dpAmt": "dpAmount", "avgDpAmt": "avgDpAmount", "dpTimeoutRatio": "dp_timeout(%)", 
      "dpSuccessRatio": "dp_success_rate(%)", "wdAmt": "wdAmount", 
      "avgSessionTime": "avg_session_withdraw_time_120m", "quickWDRatio": "utlra_fast_session_ratio", "WD x DP": "wdDpRatio",
      "ValidBet": "playAmt", "dailyBetCnt": "avgDailyBetCnt", "ValidBet/ dpAmt": "playAmt/dpAmt", "registerDate":"register_date"
    };

     // --- STEP AA: LOAD EXCLUSIONS FROM USER_RESULT SHEET ---
    statusCell.setValue("⏳ Loading exclusions from User_Result...");
    
    const blacklistEmails = {};
    const monitorEmails = {}; // Key: Phone/Email, Value: "Monitored User"
    const awaitingApproval = {}; // Key: Phone/Email, Value: "awaiting Blacklist Approval User"
    const normalCooldownEmails = {}; // Key: Phone/Email, Value: Expiration Timestamp (2 weeks later)

    if (resultSheet) {
      const resultLastRow = resultSheet.getLastRow();
      if (resultLastRow >= 5) { // Row 4 is header, data starts at row 5
        // Fetch 5 columns: Date[0], Result[1], Category[2], Phone/Email[3], Username[4]
        const resultData = resultSheet.getRange(5, 1, resultLastRow - 4, 5).getValues();
        
        resultData.forEach(row => {
          const rowDateStr   = row[0];
          const resultStatus = String(row[1]).trim();
          const emailPhone   = String(row[2]).trim(); // Index 3 is now Phone/Email

          if (!emailPhone || emailPhone === "") return; // Skip if phone is empty

          if (resultStatus === "Blacklist") {
            blacklistEmails[emailPhone] = true;
          } 
          else if (resultStatus === "Monitor") {
            monitorEmails[emailPhone] = "Monitored User";
          } 
          else if (resultStatus === "Awaiting Approval") {
            awaitingApproval[emailPhone] = true;
          } 
          else if (resultStatus === "Normal" && rowDateStr) {
            const entryDate = new Date(rowDateStr);
            // Calculate cooldown: Entry date + 14 days
            const cooldownEndDate = new Date(entryDate.getFullYear(), entryDate.getMonth(), entryDate.getDate() + 14);
            normalCooldownEmails[emailPhone] = cooldownEndDate.getTime();
          }
        });
      }
    }

    // --- STEP A: LOAD THRESHOLDS FROM REPORT SHEET ---
    statusCell.setValue("⏳ Loading thresholds from sheet...");

    const WIN_LOSS_MIN = Number(reportSheet.getRange("AD10").getValue()); 
    const WIN_LOSS_MAX = 0.1;
    const REGISTER_RATE_MAX_DIV_2 = Number(reportSheet.getRange("D16").getValue()) || 0; 
    const AVG_SESSION_TIME_P25 = Number(reportSheet.getRange("P12").getValue()) || 0; 
    const QUICK_WD_P90         = Number(reportSheet.getRange("Q15").getValue()) || 0; 
    const TIMEOUTRATIO_P90     = Number(reportSheet.getRange("L15").getValue()) || 0; 
    const TURNOVER_DP_P25      = Number(reportSheet.getRange("AA12").getValue()) || 0; 
    const DEP_FREQ_P75         = Number(reportSheet.getRange("T14").getValue()) || 0; 
    const DEP_FREQ_P90         = Number(reportSheet.getRange("T15").getValue()) || 0; 
    const PAYMENT_CHANNEL_CNT_P75 = Number(reportSheet.getRange("K14").getValue()) || 0; 
    const AVG_DEP_P25          = Number(reportSheet.getRange("I12").getValue()) || 0; 
    const AVG_DEP_P90          = Number(reportSheet.getRange("I15").getValue()) || 0; 
    const BRAND_CNT_P90        = Number(reportSheet.getRange("E15").getValue()) || 0; 
    const WD_DP_RATIO_P90      = Number(reportSheet.getRange("W15").getValue()) || 0; 


    // Logger.log(TIMEOUTRATIO_P90);

    // --- STEP B: CATEGORIZATION ---
    statusCell.setValue("⏳ Categorizing rows...");

    const givenBrandCount = reportSheet.getRange("B7").getValue().toString().trim();
    const givenDPTries = reportSheet.getRange("B8").getValue().toString().trim();
    const givenDPSuccessRatio = reportSheet.getRange("B9").getValue().toString().trim() / 100;
    const givenTurnover = reportSheet.getRange("B10").getValue().toString().trim();

    // const categoryCounts = { "QR Code Deposit Scam": 0, "Payment Channel Scraping": 0,  "PGW Crawlers" :0, "Rapid DP-WD Testing": 0, "PR Extortion": 0 , "Bonus Abuse": 0, "TurnOver Farming":0,"Normal": 0 };
    const categoryCounts = { "Suspicious": 0, "Watchlist":0, "Not match threshold": 0 };

    const groupMap = getMemberGroupMap();

    // --- DATA REDUCTION & CATEGORIZATION ---
    const filtered = rawData.reduce((acc, row) => {
      const phoneEmail = String(row[rawColMap["phone/email"]] || "").trim();

      // 1. Blacklist & Cooldown Checks
      if (blacklistEmails[phoneEmail] || awaitingApproval[phoneEmail]) return acc;
      if (normalCooldownEmails[phoneEmail] && toDateTime < normalCooldownEmails[phoneEmail]) return acc;

      const monitorLabel = monitorEmails[phoneEmail] || "";

      // 2. Metric Extraction
      const dpAmt = parseFloat(row[rawColMap["dpAmount"]]) || 0;
      const playRatio = parseFloat(row[rawColMap["playAmt/dpAmt"]]) || 0;
      const brandCount = parseInt(row[rawColMap["brandCount"]]) || 0;
      const dpTries = parseInt(row[rawColMap["dpTries"]]) || 0;
      const chanCnt = parseInt(row[rawColMap["dpChannelCnt"]]) || 0;
      const avgDep = parseFloat(row[rawColMap["avgDpAmt"]]) || 0;
      const dFreq = parseFloat(row[rawColMap["dpFrequency"]]) || 0;
      const registerRatePerDay = parseFloat(row[rawColMap["registerRatePerDay"]]) || 0;
      const quickWDRatio = parseFloat(String(row[rawColMap["quickWDRatio"]]).replace("%", "")) / 100 || 0;
      const avgSessionTime = parseFloat(row[rawColMap["avgSessionTime (Min)"]]) || 0;
      const successDpCnt = parseInt(row[rawColMap["successDpCnt"]]) || 0;
      const successWdCnt = parseInt(row[rawColMap["successWdCnt"]]) || 0;
      const wdDpRatio = parseFloat(row[rawColMap["WD x DP"]]) || 0;

      const timeoutStr = row[rawColMap["dp_timeout(%)"]] || "0";
      const timeoutRatio = typeof timeoutStr === "string" ? parseFloat(timeoutStr.replace("%", "")) / 100 : parseFloat(timeoutStr) || 0;

      const dpSuccessRatioStr = row[rawColMap["dp_success_rate(%)"]] || "0";
      const dpSuccessRatio = typeof dpSuccessRatioStr === "string" ? parseFloat(dpSuccessRatioStr.replace("%", "")) / 100 : parseFloat(dpSuccessRatioStr) || 0;

      let wlStr = row[rawColMap["Win/Loss(%)"]] || "0";
      let wlVal = typeof wlStr === "string" ? parseFloat(wlStr.replace("%", "")) / 100 : wlStr;

      // 3. Primary Status Logic (Suspicious vs Watchlist)
      let primaryStatus = "Not match threshold";
      if ((dpTries >= givenDPTries && dpSuccessRatio <= givenDPSuccessRatio && playRatio <= givenTurnover && wlVal <= WIN_LOSS_MAX) || 
          (brandCount >= givenBrandCount && playRatio <= givenTurnover && wlVal <= WIN_LOSS_MAX)) {
        primaryStatus = "Watchlist";
      }

      if ((dpAmt > 0 && wlVal <= WIN_LOSS_MAX / 2 && playRatio <= givenTurnover / 2) ||
          (dpSuccessRatio <= givenDPSuccessRatio && timeoutRatio >= 0.8 && playRatio <= givenTurnover / 2) ||
          (brandCount >= givenBrandCount && playRatio <= givenTurnover / 2 && wlVal <= WIN_LOSS_MAX / 2)) {
        primaryStatus = "Suspicious";
      }

      // 4. Sub-Category Behavior Logic
      let behaviorCat = "Unclassified"; 
      if (primaryStatus !== "Not match threshold") {
        if (dpTries >= 10 && dpSuccessRatio <= 0.1 && timeoutRatio >= TIMEOUTRATIO_P90 && playRatio <= TURNOVER_DP_P25) {
          behaviorCat = 'PGW Crawlers';
        } else if (chanCnt >= PAYMENT_CHANNEL_CNT_P75 && avgDep <= AVG_DEP_P25 && dFreq >= DEP_FREQ_P90 && playRatio <= TURNOVER_DP_P25) {
          behaviorCat = "Payment Channel Scraping";
        } else if (brandCount >= BRAND_CNT_P90 && ((quickWDRatio >= QUICK_WD_P90 || avgSessionTime <= AVG_SESSION_TIME_P25) && (successDpCnt > 0 && successWdCnt > 0)) && playRatio <= TURNOVER_DP_P25 && wlVal >= WIN_LOSS_MIN && wlVal <= WIN_LOSS_MAX) {
          behaviorCat = "PR Extortion";
        } else if (((quickWDRatio >= QUICK_WD_P90 || avgSessionTime <= AVG_SESSION_TIME_P25) && (successDpCnt > 0 && successWdCnt > 0)) && playRatio <= TURNOVER_DP_P25 && wdDpRatio >= WD_DP_RATIO_P90) {
          behaviorCat = "Rapid DP-WD Testing";
        } else if (brandCount >= BRAND_CNT_P90 && playRatio <= 1 && registerRatePerDay >= (REGISTER_RATE_MAX_DIV_2 / 2)) {
          behaviorCat = "High Brand Low Turnover";
        } 

        categoryCounts[primaryStatus]++;
        
        // 5. Build Output Row
        acc.push(columnsToShow.map(h => {
          if (h === "category") return primaryStatus;
          if (h === "status") return monitorLabel;
          // Add behavior category logic here
          if (h === "behavior_category") return behaviorCat; 
          
          return row[rawColMap[renameMap[h] || h]];
        }));
      } else {
        categoryCounts["Not match threshold"]++;
      }

      return acc;
    }, []);

  
    // --- STEP D: WRITE DATA & FORMATTING ---
    const currentMaxRows = reportSheet.getMaxRows();
    const numCols = columnsToShow.length;

    // 1. Clear previous content and remove any existing filters from row 18 down
    if (currentMaxRows >= userInfoStartRow) {
      reportSheet.getRange(userInfoStartRow, 1, currentMaxRows - (userInfoStartRow-1), numCols).clearContent().clearFormat();

      // Remove filter if one exists in the sheet
      const existingFilter = reportSheet.getFilter();
      if (existingFilter) {
        existingFilter.remove();
      }

    }
    SpreadsheetApp.flush();

    const dataBorderIndices = [3, 5, 13, 15, 17, 19, 23, 27, 30];

    if (filtered && filtered.length > 0) {
      statusCell.setValue("⏳ Writing data and formatting...");
      
      const outputRange = reportSheet.getRange(userInfoStartRow, 1, filtered.length, numCols);
      
      // 1. Write Data
      outputRange.setValues(filtered).setHorizontalAlignment("center");
      
      // 2. Apply Number Formats
      const rowFormats = columnsToShow.map(h => formatMap[h] || "@");
      const fullTableFormats = Array(filtered.length).fill(rowFormats);
      outputRange.setNumberFormats(fullTableFormats);

      // 3. SORTING LOGIC
      const categoryIdx = columnsToShow.indexOf("Category") + 1;
      const playRatioIdx = columnsToShow.indexOf("ValidBet/ dpAmt") + 1;

      if (categoryIdx > 0 && playRatioIdx > 0) {
        // Sort Priority 1: Category (Suspicious starts with 'S', Watchlist with 'W' -> Ascending)
        // Sort Priority 2: Play Ratio (Lowest first -> Ascending)
        outputRange.sort([
          { column: categoryIdx, ascending: true }, 
          { column: playRatioIdx, ascending: true }
        ]);
      } else if (playRatioIdx > 0) {
        // Fallback: Just sort by Play Ratio if category column is missing
        outputRange.sort({ column: playRatioIdx, ascending: true });
      }
      // 4. Apply Alignment & Wrapping for the last columns (status, behavior, etc.)
      // Targets the last 3 columns (status, behavior_category, etc.)
      const lastColsCount = 6;
      const metaRange = reportSheet.getRange(userInfoStartRow, numCols - (lastColsCount - 1), filtered.length, lastColsCount);
      metaRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setHorizontalAlignment("left");

      // 5. Apply Borders
      statusCell.setValue("⏳ Applying thin borders...");
      applyThinBorders(outputRange, dataBorderIndices);
      
      statusCell.setValue(`✅ Success: ${filtered.length} suspicious rows found.`);
    } else {
      statusCell.setValue(`✅ Success: 0 suspicious users detected.`);
      reportSheet.getRange(userInfoStartRow, 1).setValue("No suspicious activity detected.");
    }


    // --- STEP F: DYNAMIC SUMMARY TABLE (J2, M2, etc.) ---
    statusCell.setValue("⏳ Generating summary tables...");
    
    const summaryData = [
      ["Suspicious", categoryCounts["Suspicious"]],
      ["Watchlist", categoryCounts["Watchlist"]],
      ["Not match threshold", categoryCounts["Not match threshold"]]
    ];

    let startCol = 10; // Column J
    const rowsPerTable = 4;

    // Clear summary area (J2:O6) - covers J, K, L, M, N, O
    reportSheet.getRange("J2:O6").clearContent().clearFormat();

    for (let i = 0; i < summaryData.length; i += rowsPerTable) {
      const chunk = summaryData.slice(i, i + rowsPerTable);
      const tableHeader = [["Scenario", "# Users"]];
      const finalTable = tableHeader.concat(chunk);
      
      const destRange = reportSheet.getRange(2, startCol, finalTable.length, 2);
      destRange.setValues(finalTable);
      
      // --- Formatting ---
      // Apply thinnest border
      destRange.setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);
      
      // Alignments: Left for Scenarios (Col 1), Center for Counts (Col 2)
      destRange.offset(0, 0, finalTable.length, 1).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      destRange.offset(0, 1, finalTable.length, 1).setHorizontalAlignment("center").setVerticalAlignment("middle");
      
      // Header styling
      const headerRange = destRange.offset(0, 0, 1, 2);
      headerRange.setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

      // Move 3 columns to the right (J -> M -> P) to leave an empty column in between
      startCol += 3;
    }

    reportSheet.getRange("A2").setValue(`Updated: ${new Date()}`);
    statusCell.setValue(`✅ Success: ${filtered.length} suspicious rows found.`);
    createOrReplaceReportSheet_sus();

  } catch (err) {
    statusCell.setValue(`❌ Error: ${err.message}`);
  }
}

function createOrReplaceReportSheet_sus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Suspicious_Player");
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Suspicious_Player' sheet not found.");
    return;
  }

  // 1. Get the dates from the cells
  const fromDate = sourceSheet.getRange("B3").getValue();
  const toDate = sourceSheet.getRange("B4").getValue();
  
  if (!fromDate || !toDate) {
    SpreadsheetApp.getUi().alert("Error: Please make sure both B3 and B4 contain valid dates.");
    return;
  }

  // 2. Format the dates
  const fromDateFormatted = Utilities.formatDate(new Date(fromDate), ss.getSpreadsheetTimeZone(), "d MMMM").toLowerCase();
  const toDateFormatted = Utilities.formatDate(new Date(toDate), ss.getSpreadsheetTimeZone(), "d MMMM yyyy").toLowerCase();
  
  const newSheetName = `${fromDateFormatted} ~ ${toDateFormatted} sus user`;

  // 3. Check if the target sheet already exists
  let targetSheet = ss.getSheetByName(newSheetName);
  
  if (targetSheet) {
    targetSheet.clear();
    Logger.log("Existing sheet found and cleared: " + newSheetName);
  } else {
    targetSheet = ss.insertSheet(newSheetName);
    Logger.log("Created new sheet: " + newSheetName);
  }
  
  // 4. Copy everything from Extreme_Case_v2 to the target sheet
  const sourceRange = sourceSheet.getDataRange();
  const maxColumns = sourceRange.getLastColumn();
  
  // Ensure the target sheet has enough columns
  const targetMaxColumns = targetSheet.getMaxColumns();
  if (targetMaxColumns < maxColumns) {
    targetSheet.insertColumnsAfter(targetMaxColumns, maxColumns - targetMaxColumns);
  }

  const targetRange = targetSheet.getRange(1, 1);
  sourceRange.copyTo(targetRange);
  
  // 5. Copy column widths safely
  for (let i = 1; i <= maxColumns; i++) {
    const colWidth = sourceSheet.getColumnWidth(i);
    targetSheet.setColumnWidth(i, colWidth);
  }

  // --- NEW: Clean up empty rows at the bottom ---
  const lastRowWithData = targetSheet.getLastRow();
  const totalRowsNow = targetSheet.getMaxRows();
  const desiredTotalRows = lastRowWithData + 5; // Data rows + 5 extra buffer rows

  if (totalRowsNow > desiredTotalRows) {
    // Delete the excess rows starting right after the 5 buffer rows
    targetSheet.deleteRows(desiredTotalRows + 1, totalRowsNow - desiredTotalRows);
  } else if (totalRowsNow < desiredTotalRows) {
    // Just in case the sheet is somehow shorter, add rows up to the desired amount
    targetSheet.insertRowsAfter(totalRowsNow, desiredTotalRows - totalRowsNow);
  }
  // -----------------------------------------------

  // Bring the user to the newly generated sheet
  ss.setActiveSheet(targetSheet);
}


