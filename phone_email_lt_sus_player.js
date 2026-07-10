// --- PLACE THIS HELPER FUNCTION OUTSIDE YOUR MAIN LOOP ---
/**
 * Helper to dynamically check the 2-month historical sheet and return the proper monitor object or string fallback.
 */
function getHistoricalData(dateObj, baseLabel, ss, columnsToShow, emailPhone) {
  try {
    // 1. Calculate Start Date (2 months back)
    const startDateObj = new Date(dateObj.getFullYear(), dateObj.getMonth() - 2, dateObj.getDate());
    
    // 2. Format months to lower case text (e.g., "june", "april")
    const endMonthText = dateObj.toLocaleString('en-US', { month: 'long' }).toLowerCase();
    const startMonthText = startDateObj.toLocaleString('en-US', { month: 'long' }).toLowerCase();
    
    // 3. Construct target sheet name: "22 april ~ 22 june 2026 sus user"
    const targetSheetName = `${startDateObj.getDate()} ${startMonthText} ~ ${dateObj.getDate()} ${endMonthText} ${dateObj.getFullYear()} sus user`;
    const historicalSheet = ss.getSheetByName(targetSheetName);

    if (historicalSheet) {
      const histLastRow = historicalSheet.getLastRow();
      // Assuming data grid starts at row 14 based on your userInfoStartRow variable
      if (histLastRow >= 14) {
        const histData = historicalSheet.getRange(14, 1, histLastRow - 13, columnsToShow.length).getValues();
        
        // Find column indexes in current layout structure
        const phoneIdx = columnsToShow.indexOf("phone/email");
        const dpAmtIdx = columnsToShow.indexOf("dpAmt");
        const wdAmtIdx = columnsToShow.indexOf("wdAmt");
        const dpTriesIdx = columnsToShow.indexOf("dpTries");

        // Search for this user in the old sheet
        const oldRecord = histData.find(r => String(r[phoneIdx]).trim() === emailPhone);

        if (oldRecord && dpAmtIdx >= 0 && wdAmtIdx >= 0 && dpTriesIdx > 0) {
          // Store original values to map against rawData down the stream
          return {
            label: baseLabel,
            oldDpAmt: parseFloat(oldRecord[dpAmtIdx]) || 0,
            oldWdAmt: parseFloat(oldRecord[wdAmtIdx]) || 0,
            olddpTries: parseFloat(oldRecord[dpTriesIdx]) || 0
          };
        }
      }
    }
    return baseLabel; // Fallback if sheet or record doesn't exist/is empty
  } catch (sheetErr) {
    return baseLabel; // Resilient fallback
  }
}


function insert_categorised_lt_data() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("Suspicious_Player");
  const rawdatasheet = ss.getSheetByName("Raw_Lt_phone_email");
  const resultSheet = ss.getSheetByName("User_Result"); // <-- Added User_Result Sheet
  const statusCell = reportSheet.getRange("C6");
  const userInfoStartRow = 25 + 1;

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
      "category","phone/email", "earliestRegisteredDate", "registerRatePerDay", "brandCount","dpTries","successDpCnt", "dpAmt", "avgDpAmt",
      "dpChannelCnt", "dpTimeoutRatio","dpSuccessRatio", "successWdCnt", "wdAmt", "avgSessionTime(Min)", "quickWDRatio","promoAmt", "Promo/Deposit", 
      "dpFreq", "wdFreq", "promoFreq", "WD x DP", "ValidBet", "earnAmt",
      "turnover", "GGR", "NGR", "Win/Loss(%)", "activeBrandList", "activeUsernameList", "brandList","usernameList", "status", "behavior_category", "ReviewedBy",
      "ltDpAmt", "ltAdjAmt", "ltWdAmt",	"ltNetDpAmt",	"ltSuccessDpCnt", "ltFailDpCnt", "ltWdCnt",	"ltClaimPromo",	"ltValidBet",	"ltEarnAmt",	"ltTurnover",	"lfPlayerWlGgr"
    ];

    // --- NEW: Number Format Mapping based on your original request ---
    const formatMap = {
      // 👤 User Info
      "brandCount": "0", "registerRatePerDay": "#,##0.00",

      // 💰 Deposit
      "dpChannelCnt": "0", "dpTries": "0", "successDpCnt": "0", "dpAmt": "#,##0.00", "avgDpAmt": "#,##0.00", "maxDp": "#,##0.00", "dpTimeoutRatio": "0.00%", "dpSuccessRatio": "0.00%",

      // 💸 Withdrawal
      "successWdCnt": "0", "wdAmt": "#,##0.00",

      // 🎁 Promotion
      "promoAmt": "#,##0.00", "Promo/Deposit": "0.00%",

      // ⚡ Velocity
      "dpFreq": "0.00", "wdFreq": "0.00", "promoFreq": "0.00", "WD x DP": "0.00",

      // 🎰 Betting & Ratios
      "ValidBet": "#,##0.00", "earnAmt": "#,##0.00", "turnover": "0.00", "GGR": "#,##0.00", "NGR": "#,##0.00", "Win/Loss(%)": "0.00%",

      // ⏱️ Session Metrics
      "avgSessionTime(Min)": "0.00", "quickWDRatio": "0.00%",

      // ⏳ Lifetime Metrics
      "ltDpAmt": "#,##0.00", "ltWdAmt": "#,##0.00", "ltNetDpAmt": "#,##0.00", "ltAdjAmt": "#,##0.00", "ltSuccessDpCnt": "0", "ltFailDpCnt": "0", "ltWdCnt": "0", "ltClaimPromo": "#,##0.00", "ltValidBet": "#,##0.00","ltEarnAmt": "#,##0.00", "ltTurnover": "#,##0.00", "lfPlayerWlGgr": "#,##0.00"
    };

    const renameMap = {
      "dpTimeoutRatio": "dpTimeoutError(%)", "dpSuccessRatio": "dpSuccessRate(%)", 
      "avgSessionTime(Min)": "avg_session_withdraw_time_120m", "quickWDRatio": "utlra_fast_session_ratio", "WD x DP": "wdDpRatio",
      "ValidBet": "playAmt", "registerDate":"register_date"
    };

     // --- STEP AA: LOAD EXCLUSIONS FROM USER_RESULT SHEET ---
    statusCell.setValue("⏳ Loading exclusions from User_Result...");
    
    const blacklistPhone = {};
    const demoPhone = {};
    const monitorPhone = {}; // Key: Phone/Email, Value: "Monitored User"
    const pendingPhone = {}; // Key: Phone/Email, Value: "awaiting Blacklist Approval User"
    const normalCooldownPhone = {}; // Key: Phone/Email, Value: Expiration Timestamp (2 weeks later)

    if (resultSheet) {
      const resultLastRow = resultSheet.getLastRow();
      if (resultLastRow >= 5) { // Row 4 is header, data starts at row 5
        // Fetch 3 columns: Date[0], Phone/Email[1], Result[2]
        const resultData = resultSheet.getRange(5, 1, resultLastRow - 4, 3).getValues();
        
        resultData.forEach(row => {
          const rowDateStr   = row[0];
          
          // --- SKIP IF DATE IS TODAY ---
          if (rowDateStr) {
            const today = new Date();
            const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
            
            // Format rowDateStr to YYYY-MM-DD to ensure an accurate match
            const rowDateObj = new Date(rowDateStr);
            const formattedRowDate = `${rowDateObj.getFullYear()}-${String(rowDateObj.getMonth() + 1).padStart(2, '0')}-${String(rowDateObj.getDate()).padStart(2, '0')}`;

            if (formattedRowDate === todayStr) {
              return; // Correct way to skip to the next row inside a .forEach loop
            }
          }

          const emailPhone = String(row[1]).trim().replace(/^=/, '');
          const resultStatus = String(row[2] || "").trim().toLowerCase()

          if (!emailPhone || emailPhone === "") return; // Skip if phone is empty

          if (resultStatus === "blacklist" || resultStatus === "already blacklist") {
            blacklistPhone[emailPhone] = true;
          } 
          else if (resultStatus === "demo" || !resultStatus || resultStatus.trim() === "") {
            demoPhone[emailPhone] = true;
          } 
          else if (resultStatus === "pending") {
            pendingPhone[emailPhone] = true;
          }
          // // else if (resultStatus === "Monitor") {
          // //   if (rowDateStr) {
          // //     const dateObj = new Date(rowDateStr);
          // //     // Formats to YYYY-MM-DD manually to prevent timezone shifting
          // //     const year = dateObj.getFullYear();
          // //     const month = String(dateObj.getMonth() + 1).padStart(2, '0');
          // //     const day = String(dateObj.getDate()).padStart(2, '0');
              
          // //     monitorPhone[emailPhone] = `${year}-${month}-${day} - Monitor Acc`;
          // //   } else {
          // //     monitorPhone[emailPhone] = "Monitor Acc";
          // //   }
          // // }  
          // else if (resultStatus === "Monitor") {
          //   if (rowDateStr) {
          //     const dateObj = new Date(rowDateStr);
          //     // Formats to YYYY-MM-DD
          //     const year = dateObj.getFullYear();
          //     const month = String(dateObj.getMonth() + 1).padStart(2, '0');
          //     const day = String(dateObj.getDate()).padStart(2, '0');
          //     const baseLabel = `${year}-${month}-${day} - Monitor Acc`;

          //     // --- DYNAMIC 2-MONTH HISTORICAL SHEET CHECK ---
          //     try {
                
          //       // 1. Calculate Start Date (2 months back)
          //       const startDateObj = new Date(dateObj.getFullYear(), dateObj.getMonth() - 2, dateObj.getDate());
                
          //       // 2. Format months to lower case text (e.g., "june", "april")
          //       const endMonthText = dateObj.toLocaleString('en-US', { month: 'long' }).toLowerCase();
          //       const startMonthText = startDateObj.toLocaleString('en-US', { month: 'long' }).toLowerCase();
                
          //       // 3. Construct target sheet name: "22 april ~ 22 june 2026 sus user"
          //       const targetSheetName = `${startDateObj.getDate()} ${startMonthText} ~ ${dateObj.getDate()} ${endMonthText} ${year} sus user`;
          //       const historicalSheet = ss.getSheetByName(targetSheetName);

          //       if (historicalSheet) {
          //         const histLastRow = historicalSheet.getLastRow();
          //         // Assuming data grid starts at row 14 based on your userInfoStartRow variable
          //         if (histLastRow >= 14) {
          //           const histData = historicalSheet.getRange(14, 1, histLastRow - 13, columnsToShow.length).getValues();
                    
          //           // Find column indexes in current layout structure
          //           const phoneIdx = columnsToShow.indexOf("phone/email");
          //           const dpAmtIdx = columnsToShow.indexOf("dpAmt");
          //           const wdAmtIdx = columnsToShow.indexOf("wdAmt");

          //           // Search for this user in the old sheet
          //           const oldRecord = histData.find(r => String(r[phoneIdx]).trim() === emailPhone);

          //           if (oldRecord && dpAmtIdx >= 0 && wdAmtIdx >= 0) {
          //             // Store original values to map against rawData down the stream
          //             monitorPhone[emailPhone] = {
          //               label: baseLabel,
          //               oldDpAmt: parseFloat(oldRecord[dpAmtIdx]) || 0,
          //               oldWdAmt: parseFloat(oldRecord[wdAmtIdx]) || 0
          //             };
          //           } else {
          //             monitorPhone[emailPhone] = baseLabel; // Fallback if user missing in old sheet
          //           }
          //         } else {
          //           monitorPhone[emailPhone] = baseLabel;
          //         }
          //       } else {
          //         monitorPhone[emailPhone] = baseLabel; // Fallback if sheet doesn't exist
          //       }
          //     } catch (sheetErr) {
          //       monitorPhone[emailPhone] = baseLabel; // Resilient fallback
          //     }

          //   } else {
          //     monitorPhone[emailPhone] = "Monitor Acc";
          //   }
          // }
          // // -- Normal User
          // else if (resultStatus === "Approve" && rowDateStr) {
          //   const entryDate = new Date(rowDateStr);
          //   // Calculate cooldown: Entry date + 14 days
          //   const cooldownEndDate = new Date(entryDate.getFullYear(), entryDate.getMonth(), entryDate.getDate() + 14);
          //   normalCooldownPhone[emailPhone] = cooldownEndDate.getTime();
          // }

          else if (resultStatus === "monitor") {
            if (rowDateStr) {
              const dateObj = new Date(rowDateStr);
              // Formats to YYYY-MM-DD
              const year = dateObj.getFullYear();
              const month = String(dateObj.getMonth() + 1).padStart(2, '0');
              const day = String(dateObj.getDate()).padStart(2, '0');
              const baseLabel = `${year}-${month}-${day} - Monitor`;

              // Execute the historical check using the helper function
              monitorPhone[emailPhone] = getHistoricalData(dateObj, baseLabel, ss, columnsToShow, emailPhone);
            } else {
              monitorPhone[emailPhone] = "Monitor";
            }
          }
          // -- Normal User / Approve Status
          else if (resultStatus === "approve" && rowDateStr) {
            const entryDate = new Date(rowDateStr);
            // Calculate cooldown: Entry date + 14 days
            const cooldownEndDate = new Date(entryDate.getFullYear(), entryDate.getMonth(), entryDate.getDate() + 14);
            
            // Get today's date for the 2-week comparison
            const today = new Date();

            if (today > cooldownEndDate) {
              // --- MORE THAN 2 WEEKS: Do the same thing as the Monitor block ---
              const year = entryDate.getFullYear();
              const month = String(entryDate.getMonth() + 1).padStart(2, '0');
              const day = String(entryDate.getDate()).padStart(2, '0');
              const baseLabel = `${year}-${month}-${day} - Approve`;

              // Execute identical historical check and save to monitorPhone
              monitorPhone[emailPhone] = getHistoricalData(entryDate, baseLabel, ss, columnsToShow, emailPhone);
            } else {
              // --- LESS THAN OR EQUAL TO 2 WEEKS: Keep existing cooldown logic ---
              normalCooldownPhone[emailPhone] = cooldownEndDate.getTime();
            }
          }

        });
      }
    }

    // --- STEP A: LOAD THRESHOLDS FROM REPORT SHEET ---
    statusCell.setValue("⏳ Loading thresholds from sheet...");

    const REGISTER_RATE_MAX_DIV_2 = Number(reportSheet.getRange("D22").getValue()) || 0; 
    const BRAND_CNT_P90        = Number(reportSheet.getRange("E21").getValue()) || 0; 
    const AVG_DEP_P25          = Number(reportSheet.getRange("I18").getValue()) || 0; 
    const AVG_DEP_P90          = Number(reportSheet.getRange("I21").getValue()) || 0; 
    const PAYMENT_CHANNEL_CNT_P75 = Number(reportSheet.getRange("J20").getValue()) || 0; 
    const TIMEOUTRATIO_P90     = Number(reportSheet.getRange("K21").getValue()) || 0; 
    const AVG_SESSION_TIME_P25 = Number(reportSheet.getRange("O18").getValue()) || 0; 
    const QUICK_WD_P90         = Number(reportSheet.getRange("P21").getValue()) || 0; 
    const DEP_FREQ_P75         = Number(reportSheet.getRange("S20").getValue()) || 0; 
    const DEP_FREQ_P90         = Number(reportSheet.getRange("S21").getValue()) || 0; 
    const WD_DP_RATIO_P90      = Number(reportSheet.getRange("V21").getValue()) || 0; 
    const TURNOVER_DP_P25      = Number(reportSheet.getRange("Y18").getValue()) || 0; 

    const WIN_LOSS_MIN = Number(reportSheet.getRange("AB16").getValue()); 
    const WIN_LOSS_MAX = 0.1;


    // Logger.log(`
    //   =========================================
    //         LOADED THRESHOLD VALUES
    //   =========================================
    //   [D22] REGISTER_RATE_MAX_DIV_2 : ${REGISTER_RATE_MAX_DIV_2}
    //   [E21] BRAND_CNT_P90           : ${BRAND_CNT_P90}
    //   [I18] AVG_DEP_P25             : ${AVG_DEP_P25}
    //   [I21] AVG_DEP_P90             : ${AVG_DEP_P90}
    //   [J20] PAYMENT_CHANNEL_CNT_P75 : ${PAYMENT_CHANNEL_CNT_P75}
    //   [K21] TIMEOUTRATIO_P90        : ${TIMEOUTRATIO_P90}
    //   [O18] AVG_SESSION_TIME_P25    : ${AVG_SESSION_TIME_P25}
    //   [P21] QUICK_WD_P90            : ${QUICK_WD_P90}
    //   [S20] DEP_FREQ_P75            : ${DEP_FREQ_P75}
    //   [S21] DEP_FREQ_P90            : ${DEP_FREQ_P90}
    //   [V21] WD_DP_RATIO_P90         : ${WD_DP_RATIO_P90}
    //   [Y18] TURNOVER_DP_P25         : ${TURNOVER_DP_P25}
    //   [AB16] WIN_LOSS_MIN           : ${WIN_LOSS_MIN}
    //   [Hardcoded] WIN_LOSS_MAX       : ${WIN_LOSS_MAX}
    //   =========================================
    // `);

    // Logger.log(TIMEOUTRATIO_P90);

    // --- STEP B: CATEGORIZATION ---
    statusCell.setValue("⏳ Categorizing rows...");

    const givenDPTries          = reportSheet.getRange("B7").getValue().toString().trim();
    const givenDPSuccessRatio   = reportSheet.getRange("B8").getValue().toString().trim() / 100;
    const givenTurnover         = reportSheet.getRange("B9").getValue().toString().trim();
    const givenRegistrationRate = reportSheet.getRange("B10").getValue().toString().trim();
    const givenBrandCount       = reportSheet.getRange("B11").getValue().toString().trim();
    

    const categoryCounts = { "Suspicious": 0, "Watchlist":0, "Not match threshold": 0 };

    const groupMap = getMemberGroupMap();

    // --- Helper function to parse percentages safely ---
    const parsePercent = (val) => {
      if (val === undefined || val === null || val === "") return 0;
      // If Sheets already converted it to a decimal number (e.g., 0.9991)
      if (typeof val === "number") return val; 
      // If it's still a raw string with a '%' symbol (e.g., "99.91%")
      if (typeof val === "string") {
        return parseFloat(val.replace("%", "")) / 100 || 0;
      }
      return 0;
    };


    // --- DATA REDUCTION & CATEGORIZATION ---
    const filtered = rawData.reduce((acc, row) => {
      const phoneEmail = String(row[rawColMap["phone/email"]] || "").trim();

      // 1. Blacklist & Cooldown Checks
      if (blacklistPhone[phoneEmail] || pendingPhone[phoneEmail] || demoPhone[phoneEmail]) return acc;
      if (normalCooldownPhone[phoneEmail] && toDateTime < normalCooldownPhone[phoneEmail]) return acc;

      // const monitorLabel = monitorPhone[phoneEmail] || " ";

      // --- MODIFIED MONITOR USER EVALUATION ---
      let monitorLabel = " ";
      const monitorData = monitorPhone[phoneEmail];

      if (monitorData) {
        if (typeof monitorData === "object") {
          // Get current row amounts from raw data sheet
          const currentDpAmt = parseFloat(row[rawColMap["dpAmt"]]) || 0;
          const currentWdAmt = parseFloat(row[rawColMap["wdAmt"]]) || 0;
          const currentdpTries = parseFloat(row[rawColMap["dpTries"]]) || 0;

          // Compare with historical data
          if (currentDpAmt === monitorData.oldDpAmt && currentWdAmt === monitorData.oldWdAmt && currentdpTries <= monitorData.olddpTries + 10) {
            monitorLabel = `${monitorData.label} (SAME DP, WD)`;
          } else {
            monitorLabel = `${monitorData.label} (CHANGE DP, WD)`;
          }
        } else {
          // Fallback if the user wasn't found in the historical sheet
          monitorLabel = `${monitorData} (PAST DATA N/A)`; 
        }
      }

      //  Skip entirely if data matches exactly (Removes from output AND skips counts) ===
      if (monitorLabel.includes("(SAME DP, WD)")) {
        return acc; 
      }

      // 2. Metric Extraction
      const dpAmt = parseFloat(row[rawColMap["dpAmt"]]) || 0;
      const playRatio = parseFloat(row[rawColMap["turnover"]]) || 0;
      const brandCount = parseInt(row[rawColMap["brandCount"]]) || 0;
      const dpTries = parseInt(row[rawColMap["dpTries"]]) || 0;
      const chanCnt = parseInt(row[rawColMap["dpChannelCnt"]]) || 0;
      const avgDep = parseFloat(row[rawColMap["avgDpAmt"]]) || 0;
      const dFreq = parseFloat(row[rawColMap["dpFreq"]]) || 0;
      const registerRatePerDay = parseFloat(row[rawColMap["registerRatePerDay"]]) || 0;
      const successDpCnt = parseInt(row[rawColMap["successDpCnt"]]) || 0;
      const successWdCnt = parseInt(row[rawColMap["successWdCnt"]]) || 0;
      
      const wdDpRatio = parseFloat(row[rawColMap["wdDpRatio"]]) || 0; 
      const avgSessionTime = parseFloat(row[rawColMap["avg_session_withdraw_time_120m"]]) || 0; 

      // Percentages parsed safely without double-dividing
      const quickWDRatio   = parsePercent(row[rawColMap["utlra_fast_session_ratio"]]);
      const timeoutRatio   = parsePercent(row[rawColMap["dpTimeout(%)"]]); 
      const dpSuccessRatio = parsePercent(row[rawColMap["dpSuccessRate(%)"]]);
      const wlVal          = parsePercent(row[rawColMap["Win/Loss(%)"]]);

      // Logger.log(`Processing email: ${phoneEmail}`);
      // Logger.log(`
      //   === Metric Values ===
      //   Deposit Amount (dpAmt): ${dpAmt}
      //   Play Ratio / Turnover: ${playRatio}
      //   Brand Count: ${brandCount}
      //   Deposit Tries: ${dpTries}
      //   Channel Count: ${chanCnt}
      //   Average Deposit: ${avgDep}
      //   Deposit Frequency: ${dFreq}
      //   Register Rate Per Day: ${registerRatePerDay}
      //   Quick WD Ratio: ${quickWDRatio}
      //   Avg Session Time: ${avgSessionTime}
      //   Success Deposit Count: ${successDpCnt}
      //   Success Withdrawal Count: ${successWdCnt}
      //   WD / DP Ratio: ${wdDpRatio}
      //   Timeout Ratio: ${timeoutRatio}
      //   Deposit Success Ratio: ${dpSuccessRatio}
      //   Win/Loss Value: ${wlVal}
      //   =====================
      // `);

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
      if (primaryStatus == "Suspicious") {
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
      } else if (primaryStatus == "Watchlist") {
          categoryCounts["Watchlist"]++;
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

    const dataBorderIndices = [
      "EarliestRegisteredDate",
      "brandCount",
      "dpSuccessRatio",
      "wdAmt",
      "quickWDRatio",
      "Promo/Deposit",
      "WD x DP",
      "TurnOver",
      "Win/Loss(%)",
      "ReviewedBy"
    ]
    .map(col => columnsToShow.indexOf(col) + 1)
    .filter(idx => idx > 0); // remove columns not found

    const leftClipIndicies = [
      "activeBrandList",
      "activeUsernameList",
      "BrandList",
      "UsernameList",
      "Status"

    ].map(col => columnsToShow.indexOf(col) + 1).filter(idx => idx > 0); 

    if (filtered && filtered.length > 0) {
      
      // === SORT IN MEMORY FIRST BEFORE WRITING ===
      statusCell.setValue("⏳ Sorting by ltTurnover tiers...");
      const categoryIdx = columnsToShow.findIndex(col => col.toLowerCase() === "category");
      const playRatioIdx = columnsToShow.findIndex(col => col.toLowerCase() === "turnover");
      const ltTurnoverIdx = columnsToShow.findIndex(col => col.toLowerCase() === "ltturnover");

      if (playRatioIdx >= 0 && ltTurnoverIdx >= 0) {
        filtered.sort((rowA, rowB) => {
          const ltTurnA = parseFloat(rowA[ltTurnoverIdx]) || 0;
          const ltTurnB = parseFloat(rowB[ltTurnoverIdx]) || 0;
          const turnA = parseFloat(rowA[playRatioIdx]) || 0;
          const turnB = parseFloat(rowB[playRatioIdx]) || 0;

          // Group 1: ltTurnover is strictly less than 1
          const isGroup1A = ltTurnA < 1;
          const isGroup1B = ltTurnB < 1;

          // CONDITION 1: Rows belong to different groups -> Group 1 always goes to the top
          if (isGroup1A && !isGroup1B) return -1; // Move rowA up
          if (!isGroup1A && isGroup1B) return 1;  // Move rowB up

          // CONDITION 2: Both rows are in the SAME group -> Sort by turnover (Lowest First)
          return turnA - turnB;
        });
      } else if (playRatioIdx >= 0) {
        // Fallback: Just sort by turnover ascending if ltTurnover column is completely missing
        filtered.sort((rowA, rowB) => (parseFloat(rowA[playRatioIdx]) || 0) - (parseFloat(rowB[playRatioIdx]) || 0));
      }

      // === WRITE DATA AND FORMATTING ===
      statusCell.setValue("⏳ Writing data and formatting...");
      
      // FIX 2: Clear old data area down the sheet so old rows don't linger underneath
      const currentLastRow = reportSheet.getLastRow();
      if (currentLastRow >= userInfoStartRow) {
        reportSheet.getRange(userInfoStartRow, 1, currentLastRow - userInfoStartRow + 1, numCols).clearContent().clearFormat();
      }

      const outputRange = reportSheet.getRange(userInfoStartRow, 1, filtered.length, numCols);
      
      // 1. Write  sorted Data
      outputRange.setValues(filtered).setHorizontalAlignment("center");
      
      // 2. Apply Number Formats
      const rowFormats = columnsToShow.map(h => formatMap[h] || "@");
      const fullTableFormats = Array(filtered.length).fill(rowFormats);
      outputRange.setNumberFormats(fullTableFormats);

      // 4. Apply Alignment & Wrapping for specific clipped columns
      leftClipIndicies.forEach(colIndex => {
        const columnRange = reportSheet.getRange(userInfoStartRow, colIndex, filtered.length, 1);
        columnRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
                    .setHorizontalAlignment("left");
      });

      // 5. Apply Borders
      statusCell.setValue("⏳ Applying thin borders...");
      applyThinBorders(outputRange, dataBorderIndices);
    } else {
      statusCell.setValue(`✅ Success: 0 suspicious users detected.`);
      // Clear out the row data field completely if nothing matches
      const currentLastRow = reportSheet.getLastRow();
      if (currentLastRow >= userInfoStartRow) {
        reportSheet.getRange(userInfoStartRow, 1, currentLastRow - userInfoStartRow + 1, numCols).clearContent().clearFormat();
      }
      reportSheet.getRange(userInfoStartRow, 1).setValue("No suspicious activity detected.");
    }

    // --- STEP F: DYNAMIC SUMMARY TABLE (J2, M2, etc.) ---
    statusCell.setValue("⏳ Generating summary tables...");
    
    const summaryData = [
      ["Suspicious", categoryCounts["Suspicious"] || 0],
      ["Watchlist", categoryCounts["Watchlist"] || 0],
      ["Not match threshold", categoryCounts["Not match threshold"] || 0]
    ];

    let startCol = 10; // Column J
    const rowsPerTable = 4;

    // Clear summary area (J2:O6)
    reportSheet.getRange("J2:O6").clearContent().clearFormat();

    // Loop through chunks safely
    for (let i = 0; i < summaryData.length; i += rowsPerTable) {
      const chunk = summaryData.slice(i, i + rowsPerTable);
      const tableHeader = [["Scenario", "# Users"]];
      const finalTable = tableHeader.concat(chunk);
      
      const destRange = reportSheet.getRange(2, startCol, finalTable.length, 2);
      destRange.setValues(finalTable);
      
      // --- Formatting ---
      destRange.setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);
      
      destRange.offset(0, 0, finalTable.length, 1).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      destRange.offset(0, 1, finalTable.length, 1).setHorizontalAlignment("center").setVerticalAlignment("middle");
      
      const headerRange = destRange.offset(0, 0, 1, 2);
      headerRange.setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

      startCol += 3;
    }

    reportSheet.getRange("A2").setValue(`Updated: ${new Date()}`);
    statusCell.setValue(`✅ Success: ${filtered.length} suspicious phone/emails found`);
    // createOrReplaceReportSheet_sus();

  } catch (err) {
    statusCell.setValue(`❌ Error: ${err.message}`);
  }
}