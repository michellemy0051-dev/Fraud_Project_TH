function run_percentile_phone_email() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawdatasheet = ss.getSheetByName("Percentile_phone_email");
  const reportSheet = ss.getSheetByName("Suspicious_Player");
  
  // Define missing sheets and cells safely up front
  const statusCell = reportSheet.getRange("C5"); 
  const messageCell = reportSheet.getRange("C4");

  try {
    const fromDate = reportSheet.getRange("B3").getValue();
    const toDate = reportSheet.getRange("B4").getValue();

    // Reset previous execution messages
    statusCell.clearContent();

    // Robust date validation
    if (!fromDate || !toDate || isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
      messageCell.setValue("Please enter valid dates");
      messageCell.setFontColor("#FF0000");
      return;
    }

    if (toDate < fromDate) {
      messageCell.setValue("To date must be later than From date");
      messageCell.setFontColor("#FF0000");
      return;
    } else {
      messageCell.clearContent(); 
    }

    // 1. Subtract 1 day from B3 Safely
    const adjustedFromDate = new Date(fromDate.getTime());
    adjustedFromDate.setDate(adjustedFromDate.getDate() - 1);
    const fromBQ = formatDateForBQ(adjustedFromDate);

    // 2. Add 1 day to B4 Safely
    const adjustedToDate = new Date(toDate.getTime());
    adjustedToDate.setDate(adjustedToDate.getDate() + 1);
    const toBQ = formatDateForBQ(adjustedToDate);

    const getGroupIds = generateSqlExcludedGroups();
    // Logger.log ();

    // SQL Query
    const PERCENTILE_QUERY = `
    WITH excluded_groups AS (
      ${getGroupIds}
    )
    , all_members_base AS (
      SELECT
        COALESCE(m.id, mh.id) AS memberID,
        COALESCE(m.apiIdentifier, mh.apiIdentifier) AS apiIdentifier,
        DATE(TIMESTAMP_ADD(COALESCE(m.registerAt, mh.registerAt), INTERVAL ${userOffset} HOUR)) AS register_date,
        COALESCE(m.accountId, mh.accountId) AS accountId,
        COALESCE(m.phone, m.email, mh.phone, mh.email) AS phone_email
      FROM \`${projectId}.kz_pg_to_bq_realtime.ext_member\` m
      FULL OUTER JOIN \`${projectId}.kz_pg_to_bq_realtime.ext_member_historical\` mh ON m.id = mh.id
    )
    ,member_group_id AS (
      SELECT
        groupid,
        memberid
      FROM \`${projectId}.kz_pg_to_bq_realtime.member_group\`
      QUALIFY ROW_NUMBER() OVER (PARTITION BY memberid ORDER BY insertedAt DESC) = 1
    )
    -- memberid with demo tag
    ,initial_excluded_members AS (
      SELECT mgi.memberid
      FROM member_group_id mgi
      INNER JOIN excluded_groups eg ON mgi.groupid = eg.groupid
    )
    -- phone_email with demo tag
    ,excluded_contact_details AS (
      SELECT DISTINCT base.phone_email
      FROM all_members_base base
      INNER JOIN initial_excluded_members iem ON base.memberID = iem.memberid
      WHERE base.phone_email IS NOT NULL
    )
    -- base table to get members without demo
    ,members_raw AS (
      SELECT base.*
      FROM all_members_base base
      LEFT JOIN excluded_contact_details ecd ON base.phone_email = ecd.phone_email
      WHERE ecd.phone_email IS NULL
    )
    ,members AS (
      SELECT DISTINCT
        mr.memberID,
        mr.apiIdentifier,
        a.name AS brand,
        UPPER(CONCAT(a.gamePrefix, mr.apiIdentifier)) AS username,
        mr.register_date,
        mr.phone_email
      FROM members_raw mr
      LEFT JOIN \`${projectId}.kz_pg_to_bq_realtime.account\` a ON mr.accountId = a.id
    )
    ,phone_email_agg AS (
      SELECT
        phone_email,
        COUNT(DISTINCT brand) * 1.0 / GREATEST(DATE_DIFF(MAX(register_date), MIN(register_date), DAY), 1) AS brand_registration_rate_per_day,
        COUNT(DISTINCT brand) AS brand_count
      FROM members
      WHERE phone_email IS NOT NULL
      GROUP BY phone_email
    )
    , raw_deposit_withdrawal_promo AS (
      SELECT DISTINCT
        f.memberId,
        DATE(TIMESTAMP_ADD(f.createdAt, INTERVAL ${userOffset} HOUR)) AS order_date,
        m.phone_email,
        m.register_date,
        TIMESTAMP_ADD(f.updatedAt, INTERVAL ${userOffset} HOUR) AS updatedAt,
        m.username,
        m.brand,
        f.orderRef AS order_id,
        f.netAmount AS amount,
        f.providerkey AS payment_channel,
        f.status,
        f.type,
        f.country
      FROM \`${projectId}.crm_gold_prod.funding_tx_daily\` f
      INNER JOIN members m ON f.memberId = m.memberID 
      WHERE f.insertedDate >= '${fromBQ}'
          AND f.type IN ('deposit', 'withdraw', 'promotion')
          AND f.country = '${country_prefix}'
          AND f.createdAt BETWEEN '${fromBQ}' AND '${toBQ}'
    )
    , agg_transactions AS (
      SELECT
        memberId,
        ANY_VALUE(phone_email) AS phone_email,
        ANY_VALUE(country) AS country, 
        COUNT(DISTINCT brand) AS brand_count,
        
        -- --- DEPOSIT METRICS ---
        COUNT(DISTINCT IF(type = 'deposit', payment_channel, NULL)) AS payment_channel_count,
        COUNT(DISTINCT IF(type = 'deposit', order_id, NULL)) AS deposit_tries,
        COUNT(DISTINCT IF(type = 'deposit' AND status = 'completed', order_id, NULL)) AS complete_deposit_counts,
        COUNT(DISTINCT IF(type = 'deposit' AND status = 'timeout', order_id, NULL)) AS timeout_counts,
        SUM(IF(type = 'deposit' AND status = 'completed', CAST(amount AS FLOAT64), 0.0)) AS deposit_amount,
        
        -- --- WITHDRAW METRICS ---
        COUNT(DISTINCT IF(type = 'withdraw', order_id, NULL)) AS wd_tries,
        COUNT(DISTINCT IF(type = 'withdraw' AND status = 'completed', order_id, NULL)) AS success_wd_count,
        SUM(IF(type = 'withdraw' AND status = 'completed', CAST(amount AS FLOAT64), 0.0)) AS success_wd_amount,

        -- --- PROMOTION METRICS ---
        COUNT(DISTINCT IF(type = 'promotion' AND status = 'completed', order_id, NULL)) AS redeem_promo_count,
        SUM(IF(type = 'promotion' AND status = 'completed', CAST(amount AS FLOAT64), 0.0)) AS promo_amount,

        -- --- FREQUENCY DERIVATIONS ---
        SAFE_DIVIDE(
          COUNT(DISTINCT IF(type = 'deposit' AND status = 'completed', order_id, NULL)), 
          COUNT(DISTINCT IF(type = 'deposit' AND status = 'completed', order_date, NULL))
        ) AS deposit_frequency,
        
        SAFE_DIVIDE(
          COUNT(DISTINCT IF(type = 'withdraw' AND status = 'completed', order_id, NULL)), 
          COUNT(DISTINCT IF(type = 'withdraw' AND status = 'completed', order_date, NULL))
        ) AS withdrawal_frequency,
        
        SAFE_DIVIDE(
          COUNT(DISTINCT IF(type = 'promotion' AND status = 'completed', order_id, NULL)), 
          COUNT(DISTINCT IF(type = 'promotion' AND status = 'completed', order_date, NULL))
        ) AS promo_frequency
      FROM raw_deposit_withdrawal_promo
      GROUP BY memberId
    )
    , agg_transactions_phone AS (
      SELECT
        phone_email,
        ANY_VALUE(country) AS country, 
        SUM(payment_channel_count) AS payment_channel_count,
        SUM(deposit_tries) AS deposit_tries,
        SUM(complete_deposit_counts) AS complete_deposit_counts,
        SUM(timeout_counts) AS timeout_counts,
        SUM(deposit_amount) AS deposit_amount,
        
        SUM(wd_tries) AS wd_tries,
        SUM(success_wd_count) AS success_wd_count,
        SUM(success_wd_amount) AS success_wd_amount,
        
        SUM(redeem_promo_count) AS redeem_promo_count,
        SUM(promo_amount) AS promo_amount,
        
        AVG(deposit_frequency) AS deposit_frequency,
        AVG(withdrawal_frequency) AS withdrawal_frequency,
        AVG(promo_frequency) AS promo_frequency
      FROM agg_transactions
      WHERE phone_email IS NOT NULL
      GROUP BY phone_email
    )
    , deposit_sessions AS (
      SELECT
        memberId,
        phone_email, 
        updatedAt,
        type,
        amount,
        SUM(CASE WHEN type = 'deposit' THEN 1 ELSE 0 END) OVER (PARTITION BY memberId ORDER BY updatedAt) AS session_id
      FROM raw_deposit_withdrawal_promo
      WHERE status = 'completed'
    )
    , session_boundaries AS (
      SELECT
        *,
        MIN(IF(type = 'deposit', updatedAt, NULL)) OVER (PARTITION BY memberId, session_id) AS session_start
      FROM deposit_sessions
    )
    , session_cashout AS (
      SELECT
        phone_email,
        memberId,
        session_id,
        TIMESTAMP_DIFF( 
          MIN(IF(type = 'withdraw', updatedAt, NULL)),
          MIN(IF(type = 'deposit', updatedAt, NULL)),
          MINUTE) AS minutes_to_withdraw
      FROM session_boundaries
      WHERE session_start IS NOT NULL
        AND TIMESTAMP_DIFF(updatedAt, session_start, MINUTE) <= 120
      GROUP BY phone_email, memberId, session_id
    )
    , session_metrics_phone AS (
      SELECT
        phone_email,
        AVG(minutes_to_withdraw) AS avg_session_withdraw_time_120m,
        SUM(COUNTIF(minutes_to_withdraw <= 5)) OVER(PARTITION BY phone_email) AS ultra_fast_sessions,
        COUNT(*) AS total_sessions
      FROM session_cashout
      WHERE minutes_to_withdraw IS NOT NULL
      GROUP BY phone_email
    )
    , turnover_deduped AS (
      SELECT
        td.memberId,
        m.phone_email, 
        SUM(td.playAmt) AS total_playAmt,
        SUM(td.earnAmt) AS total_earnAmt
      FROM \`${projectId}.crm_gold_prod.turnover_daily\` td
      INNER JOIN members m ON td.memberId = m.memberID
      WHERE td.insertedDate >= '${fromBQ}'
            AND currency = 'THB'
            AND td.deletedAt IS NULL
            AND td.date IS NOT NULL
            AND td.date BETWEEN '${fromBQ}' AND '${toBQ}'
        GROUP BY td.memberId, m.phone_email
    )
    , turnover_phone AS (
      SELECT
        phone_email,
        SUM(total_playAmt) AS total_playAmt,
        SUM(total_earnAmt) AS total_earnAmt
      FROM turnover_deduped
      GROUP BY phone_email
    )
    ,active_user_agg AS (
      SELECT
          m.phone_email,
          STRING_AGG(DISTINCT m.username, ',' ORDER BY m.username) AS active_username_list,
          STRING_AGG(DISTINCT m.brand, ',' ORDER BY m.brand) AS active_brand_list
      FROM members m
      LEFT JOIN agg_transactions ar ON m.memberID = ar.memberId
      LEFT JOIN turnover_deduped t ON m.memberID = t.memberId
      WHERE ar.memberId IS NOT NULL OR t.memberId IS NOT NULL
      GROUP BY m.phone_email
    )
    , risk_metrics AS (
      SELECT
        -- 👤 User Info
        pea.phone_email,
        pea.brand_count,
        pea.brand_registration_rate_per_day,

        -- 💰 Deposit
        ar.payment_channel_count,
        ar.deposit_tries,
        ar.complete_deposit_counts,
        ar.timeout_counts,
        ar.deposit_amount,

        -- 💸 Withdrawal
        ar.wd_tries,
        ar.success_wd_count,
        ar.success_wd_amount,

        -- 🎁 Promotion
        ar.redeem_promo_count,
        ar.promo_amount,
        
        -- ⚡ Frequency
        ar.deposit_frequency,
        ar.withdrawal_frequency,
        ar.promo_frequency,

        -- 🎰 Betting
        COALESCE(t.total_playAmt, 0) AS total_playAmt,
        COALESCE(t.total_earnAmt, 0) AS total_earnAmt,

        -- Clean ratios (Calculated natively at the 1:1 phone layer)
        SAFE_DIVIDE(t.total_playAmt, ar.deposit_amount) AS turnover,
        SAFE_DIVIDE(ar.deposit_amount, ar.complete_deposit_counts) AS avg_deposit_size,
        SAFE_DIVIDE(ar.deposit_amount - ar.success_wd_amount, ar.deposit_amount) AS win_loss_ratio,
        
        -- 📊 Additional Derived Metrics
        SAFE_DIVIDE(ar.timeout_counts, ar.deposit_tries) AS deposit_timeout_rate,
        SAFE_DIVIDE(ar.complete_deposit_counts, ar.deposit_tries) AS deposit_success_rate,
        SAFE_DIVIDE(sm.ultra_fast_sessions, sm.total_sessions) AS quick_withdrawal_rate,
        SAFE_DIVIDE(ar.promo_amount, ar.deposit_amount) AS dp_promo_amount_ratio,
        SAFE_DIVIDE(ar.success_wd_count, ar.complete_deposit_counts) AS dp_wd_ratio,
        
        -- ⏱️ Session Metrics    
        sm.avg_session_withdraw_time_120m,
        sm.ultra_fast_sessions,
        sm.total_sessions,
        (t.total_playAmt - t.total_earnAmt) AS ggr,
        (t.total_playAmt - t.total_earnAmt - ar.promo_amount) AS ngr

      FROM phone_email_agg pea
      LEFT JOIN agg_transactions_phone ar ON pea.phone_email = ar.phone_email
      LEFT JOIN turnover_phone t ON pea.phone_email = t.phone_email
      LEFT JOIN session_metrics_phone sm ON pea.phone_email = sm.phone_email
      LEFT JOIN active_user_agg aua ON pea.phone_email = aua.phone_email
    )
    SELECT
        TO_JSON_STRING(APPROX_QUANTILES(brand_registration_rate_per_day, 100)) AS brand_registration_rate_per_day_q,
        TO_JSON_STRING(APPROX_QUANTILES(brand_count, 100)) AS brand_count_q,
        TO_JSON_STRING(APPROX_QUANTILES(deposit_tries, 100)) AS deposit_tries_q,
        TO_JSON_STRING(APPROX_QUANTILES(complete_deposit_counts, 100)) AS complete_deposit_counts_q,
        TO_JSON_STRING(APPROX_QUANTILES(deposit_amount, 100)) AS deposit_amount_q,
        TO_JSON_STRING(APPROX_QUANTILES(avg_deposit_size, 100)) AS avg_deposit_size_q,
        TO_JSON_STRING(APPROX_QUANTILES(payment_channel_count, 100)) AS payment_channel_count_q,
        TO_JSON_STRING(APPROX_QUANTILES(deposit_timeout_rate, 100)) AS deposit_timeout_rate_q,
        TO_JSON_STRING(APPROX_QUANTILES(deposit_success_rate, 100)) AS deposit_success_rate_q,
        TO_JSON_STRING(APPROX_QUANTILES(success_wd_count, 100)) AS success_wd_count_q,
        TO_JSON_STRING(APPROX_QUANTILES(success_wd_amount, 100)) AS success_wd_amount_q,
        TO_JSON_STRING(APPROX_QUANTILES(avg_session_withdraw_time_120m, 100)) AS session_time_q,
        TO_JSON_STRING(APPROX_QUANTILES(quick_withdrawal_rate, 100)) AS quick_withdrawal_rate_q,
        TO_JSON_STRING(APPROX_QUANTILES(promo_amount, 100)) AS promo_amount_q,
        TO_JSON_STRING(APPROX_QUANTILES(dp_promo_amount_ratio, 100)) AS dp_promo_amount_q,
        TO_JSON_STRING(APPROX_QUANTILES(deposit_frequency, 100)) AS deposit_frequency_q,
        TO_JSON_STRING(APPROX_QUANTILES(withdrawal_frequency, 100)) AS withdrawal_frequency_q,
        TO_JSON_STRING(APPROX_QUANTILES(promo_frequency, 100)) AS promo_frequency_q,
        TO_JSON_STRING(APPROX_QUANTILES(dp_wd_ratio, 100)) AS dp_wd_ratio_q,
        TO_JSON_STRING(APPROX_QUANTILES(total_playAmt, 100)) AS total_playAmt_q,
        TO_JSON_STRING(APPROX_QUANTILES(total_earnAmt, 100)) AS total_earnAmt_q,
        TO_JSON_STRING(APPROX_QUANTILES(turnover, 100)) AS turnover_ratio_q,
        TO_JSON_STRING(APPROX_QUANTILES(ggr, 100)) AS ggr_q,
        TO_JSON_STRING(APPROX_QUANTILES(ngr, 100)) AS ngr_q,
        TO_JSON_STRING(APPROX_QUANTILES(win_loss_ratio, 100)) AS win_loss_ratio_q
    FROM risk_metrics rm
    WHERE
      rm.brand_count <> 0
      AND (
          rm.complete_deposit_counts > 10 
          OR (rm.deposit_tries > 20 AND rm.complete_deposit_counts / rm.deposit_tries <= 0.1)
      );
    
    `;

    // Logger.log(PERCENTILE_QUERY);

    const lineChunkSize = 8000;
    for (let i = 0; i < PERCENTILE_QUERY.length; i += lineChunkSize) {
        Logger.log(PERCENTILE_QUERY.substring(i, i + lineChunkSize));
    }

    // ==========================================
    // CLEAR ROW 7 DOWNWARDS SAFE CLEANUP
    // ==========================================
    if (rawdatasheet.getLastRow() >= 8) {
      rawdatasheet.getRange(8, 1, rawdatasheet.getMaxRows() - 7, rawdatasheet.getMaxColumns()).breakApart();
      rawdatasheet.getRange(8, 1, rawdatasheet.getMaxRows() - 7, rawdatasheet.getMaxColumns()).clearContent();
    }
    
    statusCell.setValue(`Running...`).setFontColor("#38761d");
    SpreadsheetApp.flush();

    // Row 8 Detailed Sub-metrics
    const columnsOrder = [
      "registerRatePerDay", "brandCount", "dpTries", "successDpCnt", "dpAmt", "avgDpAmt", "dpChannelCnt", "dpTimeoutRatio", "dpSuccessRatio",
      "successWdCnt", "wdAmount", "avgSessionTime", "quickWDRatio", "promoAmt", "Promo/Deposit",
      "dpFreq", "wdFreq", "promoFreq", "WD x DP", "ValidBet", "earnAmt", "ValidBet/dpAmt",
      "GGR", "NGR", "Win/Loss(%)"
    ];
    
    // Fixed: written to row 8 column B to align directly underneath categories
    rawdatasheet.getRange(8, 2, 1, columnsOrder.length).setValues([columnsOrder]);
    
    // Style formatting for headers
    const headerRange = rawdatasheet.getRange("A7:AA8");
    headerRange.setFontWeight("bold").setBackground("#D3D3D3").setHorizontalAlignment("center");

    // ==========================================
    // RUN BIGQUERY STREAM AND TRANSPOSE DATA
    // ==========================================
    const percentileLabels = ["Min", "P10", "P25", "P50", "P75", "P90", "Max"];
    const percentileIndices = [0, 10, 25, 50, 75, 90, 100];

    // Initialize matrix table (7 rows x 26 columns)
    let matrixTable = percentileLabels.map(label => [label, ...Array(columnsOrder.length).fill(null)]);

    runQueryStream(PERCENTILE_QUERY, row => {
      row.forEach((metricJson, metricIdx) => {
        if (!metricJson) return;
        const q = (typeof metricJson === "string") ? JSON.parse(metricJson) : metricJson;
        
        percentileIndices.forEach((pctIndex, rowIdx) => {
          matrixTable[rowIdx][metricIdx + 1] = q[pctIndex];
        });
      });
    });

    // Fixed: Writes data directly to Row 9 down to Row 15, protecting rows 7 & 8
    rawdatasheet.getRange(9, 1, matrixTable.length, matrixTable[0].length).setValues(matrixTable);

    // ==========================================
    // POST-WRITING BEAUTIFUL NUMERIC FORMATTING
    // ==========================================
    // --- 1. Integer Counts (#,##0) ---
    rawdatasheet.getRange("C9:E15").setNumberFormat("#,##0");      // Brand count, Deposit Tries, Comp Deposits
    rawdatasheet.getRange("H9:H15").setNumberFormat("#,##0");      // Payment Channel Count
    rawdatasheet.getRange("K9:K15").setNumberFormat("#,##0");      // Success Withdraw Count
    

    // --- 2. Decimals & Financials (#,##0.00) ---
    rawdatasheet.getRange("B9").setNumberFormat("#,##0.00");        // Brand Registration Rate
    rawdatasheet.getRange("F9:G15").setNumberFormat("#,##0.00");    // Deposit Amount, Avg Deposit Size
    rawdatasheet.getRange("L9:L15").setNumberFormat("#,##0.00");    // Success Withdraw Amount
    rawdatasheet.getRange("M9:M15").setNumberFormat("#,##0.00");    // Avg DP-Withdraw Time
    rawdatasheet.getRange("O9:O15").setNumberFormat("#,##0.00");    // Promo Amount
    rawdatasheet.getRange("Q9:W15").setNumberFormat("#,##0.00");    // Frequencies (Dp/Wd/Promo), DpWd Ratio, Play/Earn Amt, Turnover Ratio
    rawdatasheet.getRange("X9:Y15").setNumberFormat("#,##0.00");    // GGR, NGR

    // --- 3. Rates & Percentages (0.00%) ---
    rawdatasheet.getRange("I9:J15").setNumberFormat("0.00%");       // Deposit Timeout Rate (Col I) & Success Rate (Col J)
    rawdatasheet.getRange("N9").setNumberFormat("0.00%");           // Quick Withdrawal Rate (Col N)
    rawdatasheet.getRange("P9").setNumberFormat("0.00%");           // Promo/Deposit Amount Ratio (Col P)
    rawdatasheet.getRange("Z9:Z15").setNumberFormat("0.00%");       // Win/Loss Ratio (Col Z)
    
    rawdatasheet.getRange("A2").setValue(`Updated: ${new Date()}`);
    rawdatasheet.setFrozenRows(8);
    statusCell.setValue(`✅ Finished running percentile at ${new Date()}, Now getting raw data`).setFontColor("#38761d");
    
    SpreadsheetApp.flush();

  } catch (error) {
    Logger.log(error);
    statusCell.setValue(`❌ Error: ${error.message}`);
    throw error;
  }
}