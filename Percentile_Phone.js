function percentile_calculator_phone_email() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawdatasheet = ss.getSheetByName("Percentile_phone_email");
  const reportSheet = ss.getSheetByName("RT_Report_v4");
  
  // Define missing sheets and cells safely up front
  const statusCell = rawdatasheet.getRange("C5"); 
  const messageCell = rawdatasheet.getRange("C4");

  try {
    const fromDate = rawdatasheet.getRange("B3").getValue();
    const toDate = rawdatasheet.getRange("B4").getValue();

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
      ,all_members_base AS (
        SELECT
          COALESCE(m.id, mh.id) AS memberID,
          COALESCE(m.apiIdentifier, mh.apiIdentifier) AS apiIdentifier,
          DATE(TIMESTAMP_ADD(COALESCE(m.registerAt, mh.registerAt), INTERVAL 7 HOUR)) AS register_date,
          COALESCE(m.accountId, mh.accountId) AS accountId,
          COALESCE(m.phone, m.email, mh.phone, mh.email) AS phone_email
        FROM \`kz-dp-prod.kz_pg_to_bq_realtime.ext_member\` m
        FULL OUTER JOIN \`kz-dp-prod.kz_pg_to_bq_realtime.ext_member_historical\` mh ON m.id = mh.id
      )
      ,member_group_id AS (
        SELECT
          groupid,
          memberid
        FROM \`kz-dp-prod.kz_pg_to_bq_realtime.member_group\`
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
      LEFT JOIN \`kz-dp-prod.kz_pg_to_bq_realtime.account\` a ON mr.accountId = a.id
    )
    ,phone_email_agg AS (
      SELECT
        phone_email,
        MIN(register_date) AS earliest_registered_date,
        COUNT(DISTINCT brand) * 1.0 / GREATEST(DATE_DIFF(MAX(register_date), MIN(register_date), DAY), 1) AS brand_registration_rate_per_day,
        COUNT(DISTINCT brand) AS brand_count,
        STRING_AGG(DISTINCT username, ',' ORDER BY username ASC) AS username_list,
        STRING_AGG(DISTINCT brand, ',' ORDER BY brand ASC) AS brand_list,
      FROM members
      WHERE phone_email IS NOT NULL
      GROUP BY phone_email
    )
    ,raw_deposit_withdrawal_promo AS (
        SELECT DISTINCT
          f.memberId,
          DATE(TIMESTAMP_ADD(f.createdAt, INTERVAL 7 HOUR)) AS order_date,
          m.phone_email,
          m.register_date,
          TIMESTAMP_ADD(f.updatedAt, INTERVAL 7 HOUR) AS updatedAt,
          m.username,
          m.brand,
          f.orderRef AS order_id,
          f.netAmount AS amount,
          f.providerkey AS payment_channel,
          f.status,
          f.type,
          f.country,
        CASE 
          WHEN STARTS_WITH(remark, 'Play spin tx') THEN 'Play spin tx'
          WHEN STARTS_WITH(remark, 'Referral bonus') THEN 'Referral bonus'
          WHEN STARTS_WITH(remark, 'Claim mega share reward (friend)') THEN 'Claim mega share reward (friend)'
          WHEN STARTS_WITH(remark, 'Claim mega share reward (sharer)') THEN 'Claim mega share reward (sharer)'
          WHEN STARTS_WITH(remark, 'Skinner box txs') THEN 'Skinner box txs'
          WHEN STARTS_WITH(remark, 'Secret code participation') THEN 'Secret code participation'
          WHEN STARTS_WITH(remark, 'Secret code tx') THEN 'Secret code tx (sharee)'
          WHEN STARTS_WITH(remark, 'ILV redemption order ID') THEN 'ILV redemption order'
          ELSE remark 
        END AS cleaned_remark
      FROM \`kz-dp-prod.crm_gold_prod.funding_tx_daily\` f
      INNER JOIN members m ON f.memberId = m.memberID 
      WHERE
        f.type IN ('deposit', 'withdraw', 'promotion')
      AND f.insertedAt >= '${fromBQ}'
      AND f.createdAt BETWEEN '${fromBQ}' AND '${toBQ}'
      AND f.country = '${country_prefix}'
    )
    ,agg_transactions AS (
      SELECT
        memberId,
        ANY_VALUE(phone_email) AS phone_email,
        ANY_VALUE(username) AS username,
        ANY_VALUE(brand) AS brand,
        MIN(register_date) AS register_date,
        STRING_AGG(DISTINCT username, ',' ORDER BY username ASC) AS username_list,
        ANY_VALUE(country) AS country,
        
        COUNT(DISTINCT brand) AS brand_count,
        COUNT(DISTINCT IF(type = 'deposit', payment_channel, NULL)) AS payment_channel_count,
        COUNT(DISTINCT IF(type = 'deposit', order_id, NULL)) AS deposit_tries,
        MAX(IF(type = 'deposit', amount, NULL)) AS max_received_amount,
        COUNT(DISTINCT IF(type = 'deposit' AND status = 'completed', order_id, NULL)) AS complete_deposit_counts,
        COUNT(DISTINCT IF(type = 'deposit' AND status IN ('timeout'), order_id, NULL)) AS timeout_counts,
        COUNT(DISTINCT IF(type = 'deposit' AND status IN ('error', 'timeout'), order_id, NULL)) AS error_timeout_counts,
        SUM(IF(type = 'deposit' AND status = 'completed', CAST(amount AS FLOAT64), 0)) AS deposit_amount,
        
        COUNT(DISTINCT IF(type = 'withdraw', order_id, NULL)) AS wd_tries,
        COUNT(DISTINCT IF(type = 'withdraw' AND status = 'completed', order_id, NULL)) AS success_wd_count,
        SUM(IF(type = 'withdraw' AND status = 'completed', CAST(amount AS FLOAT64), 0)) AS success_wd_amount,

        COUNT(DISTINCT IF(type = 'promotion' AND status = 'completed', order_id, NULL)) AS redeem_promo_count,
        SUM(IF(type = 'promotion' AND status = 'completed', CAST(amount AS FLOAT64), 0)) AS promo_amount,
        STRING_AGG(DISTINCT IF(type = 'promotion' AND status = 'completed', cleaned_remark, NULL), ',' 
          ORDER BY IF(type = 'promotion' AND status = 'completed', cleaned_remark, NULL) ASC) AS cleaned_remark_list,

        SAFE_DIVIDE(COUNT(DISTINCT IF(type = 'deposit' AND status = 'completed', order_id, NULL)), COUNT(DISTINCT IF(type = 'deposit' AND status = 'completed', order_date, NULL))) AS deposit_frequency,
        SAFE_DIVIDE(COUNT(DISTINCT IF(type = 'withdraw' AND status = 'completed', order_id, NULL)), COUNT(DISTINCT IF(type = 'withdraw' AND status = 'completed', order_date, NULL))) AS withdrawal_frequency,
        SAFE_DIVIDE(COUNT(DISTINCT IF(type = 'promotion' AND status = 'completed', order_id, NULL)), COUNT(DISTINCT IF(type = 'promotion' AND status = 'completed', order_date, NULL))) AS promo_frequency
      FROM raw_deposit_withdrawal_promo
      GROUP BY memberId
    )
    ,deposit_sessions AS (
      SELECT
        memberId,
        updatedAt,
        type,
        amount,
        SUM(CASE WHEN type = 'deposit' THEN 1 ELSE 0 END) OVER ( PARTITION BY memberId ORDER BY updatedAt) AS session_id
      FROM raw_deposit_withdrawal_promo
      WHERE status = 'completed'
    )
    ,session_boundaries AS (
      SELECT
        *,
        MIN( IF( type = 'deposit',updatedAt,NULL)) OVER (PARTITION BY memberId, session_id ) AS session_start
      FROM deposit_sessions
    ),
    session_filtered AS (
      SELECT *
      FROM session_boundaries
      WHERE
        session_start IS NOT NULL
        AND TIMESTAMP_DIFF(updatedAt, session_start, MINUTE) <= 120
    )
    , session_cashout AS (
      SELECT
        memberId,
        session_id,
        MIN( IF( type = 'deposit',updatedAt, NULL)) AS deposit_time,
        MIN( IF( type = 'withdraw', updatedAt, NULL)) AS first_withdraw_time,
        TIMESTAMP_DIFF( 
          MIN(IF(type = 'withdraw', updatedAt, NULL)),
          MIN(IF(type = 'deposit', updatedAt, NULL)),
          MINUTE) AS minutes_to_withdraw
      FROM session_filtered
      GROUP BY memberId, session_id
    )
    , session_metrics AS (
        SELECT
          memberId,
          AVG(minutes_to_withdraw) AS avg_session_withdraw_time_120m,
          COUNTIF(minutes_to_withdraw <= 5) AS ultra_fast_sessions,
          COUNT(*) AS total_sessions
        FROM session_cashout
        WHERE minutes_to_withdraw IS NOT NULL
        GROUP BY memberId
    )
    , turnover_deduped AS (
      SELECT
        dt.memberId,
        DATE(TIMESTAMP_ADD(dt.createdAt, INTERVAL 7 HOUR)) AS bet_date,
        playAmt,
        earnAmt
      FROM \`kz-dp-prod.kz_pg_to_bq_realtime.turnover\` dt
      INNER JOIN members m ON dt.memberId = m.memberID
      WHERE
        dt.deletedAt IS NULL
        AND dt.date IS NOT NULL
        AND LEFT(currency, 2) = '${country_prefix}'
        AND dt.insertedAt >= '${fromBQ}'
        AND dt.createdAt BETWEEN '${fromBQ}' AND '${toBQ}'
      QUALIFY ROW_NUMBER() OVER (PARTITION BY dt.id ORDER BY dt.insertedAt DESC) = 1
    ),
    turnover AS (
      SELECT
        memberId,
        COUNT(*) AS total_bet_cnt,
        COUNT(*) / NULLIF( COUNT(DISTINCT bet_date),0) AS avg_daily_bet_cnt,
        MAX(playAmt) AS max_playAmt,
        SUM(playAmt) AS total_playAmt,
        SUM(earnAmt) AS total_earnAmt
      FROM turnover_deduped
      GROUP BY memberId
    )
    ,active_user_agg AS (
        SELECT
            m.phone_email,
            STRING_AGG(DISTINCT m.username, ',' ORDER BY m.username) AS active_username_list,
            STRING_AGG(DISTINCT m.brand, ',' ORDER BY m.brand) AS active_brand_list
        FROM members m
        LEFT JOIN agg_transactions ar ON m.memberID = ar.memberId
        LEFT JOIN turnover t ON m.memberID = t.memberId
        WHERE ar.memberId IS NOT NULL OR t.memberId IS NOT NULL
        GROUP BY m.phone_email
      )
    , risk_metrics AS (
      SELECT
        pea.phone_email,
        ANY_VALUE(pea.username_list) AS username_list,
        ANY_VALUE(aua.active_username_list) AS active_username_list,
        ANY_VALUE(pea.brand_count) AS brand_count,
        ANY_VALUE(pea.earliest_registered_date) AS earliest_registered_date,
        ANY_VALUE(pea.brand_registration_rate_per_day) AS brand_registration_rate_per_day,
        ANY_VALUE(pea.brand_list) AS brand_list,
        ANY_VALUE(aua.active_brand_list) AS active_brand_list,

        ANY_VALUE(ar.country) AS country,

        SUM(ar.payment_channel_count) AS payment_channel_count,
        SUM(ar.deposit_tries) AS deposit_tries,
        SUM(ar.complete_deposit_counts) AS complete_deposit_counts,
        SUM(ar.timeout_counts) AS timeout_counts,
        SUM(ar.error_timeout_counts) AS error_timeout_counts,

        SUM(ar.deposit_amount) AS deposit_amount,
        MAX(ar.max_received_amount) AS max_received_amount,

        SUM(ar.wd_tries) AS wd_tries,
        SUM(ar.success_wd_count) AS success_wd_count,
        SUM(ar.success_wd_amount) AS success_wd_amount,

        SUM(ar.redeem_promo_count) AS redeem_promo_count,
        SUM(ar.promo_amount) AS promo_amount,

        STRING_AGG(DISTINCT ar.cleaned_remark_list, ',') AS cleaned_remark_list,
        AVG(ar.deposit_frequency) AS deposit_frequency,
        AVG(ar.withdrawal_frequency) AS withdrawal_frequency,
        AVG(ar.promo_frequency) AS promo_frequency,

        SUM(COALESCE(t.total_bet_cnt, 0)) AS total_bet_cnt,
        SUM(COALESCE(t.total_playAmt, 0)) AS total_playAmt,
        SUM(COALESCE(t.total_earnAmt, 0)) AS total_earnAmt,
        MAX(COALESCE(t.max_playAmt, 0)) AS max_playAmt,

        ROUND(AVG(COALESCE(t.avg_daily_bet_cnt, 0)), 2) AS avg_daily_bet_cnt,
        SAFE_DIVIDE(SUM(COALESCE(t.total_playAmt, 0)), SUM(ar.deposit_amount)) AS validbet_deposit_ratio,
        SAFE_DIVIDE(SUM(ar.deposit_amount), SUM(ar.complete_deposit_counts)) AS avg_deposit_size,
        SAFE_DIVIDE(SUM(ar.deposit_amount) - SUM(ar.success_wd_amount), SUM(ar.deposit_amount)) AS win_loss_ratio,
        AVG(sm.avg_session_withdraw_time_120m) AS avg_session_withdraw_time_120m,
        SUM(sm.ultra_fast_sessions) AS ultra_fast_sessions,
        SUM(sm.total_sessions) AS total_sessions,
        SUM(COALESCE(t.total_playAmt, 0)) - SUM(COALESCE(t.total_earnAmt, 0)) AS ggr,
        SUM(COALESCE(t.total_playAmt, 0)) - SUM(COALESCE(t.total_earnAmt, 0)) - SUM(COALESCE(ar.promo_amount, 0)) AS ngr
      FROM agg_transactions ar
      LEFT JOIN turnover t ON ar.memberId = t.memberId
      LEFT JOIN session_metrics sm ON ar.memberId = sm.memberId
      LEFT JOIN phone_email_agg pea ON ar.phone_email = pea.phone_email
      LEFT JOIN active_user_agg aua ON pea.phone_email = aua.phone_email
      WHERE pea.phone_email IS NOT NULL
      GROUP BY pea.phone_email
    )
    SELECT
      TO_JSON_STRING(APPROX_QUANTILES(brand_registration_rate_per_day, 100)) AS brand_registration_rate_per_day_q,
      TO_JSON_STRING(APPROX_QUANTILES(brand_count, 100)) AS brand_count_q,
      TO_JSON_STRING(APPROX_QUANTILES(deposit_tries, 100)) AS deposit_tries_q,
      TO_JSON_STRING(APPROX_QUANTILES(complete_deposit_counts, 100)) AS complete_deposit_counts_q,
      TO_JSON_STRING(APPROX_QUANTILES(deposit_amount, 100)) AS deposit_amount_q,
      TO_JSON_STRING(APPROX_QUANTILES(avg_deposit_size, 100)) AS avg_deposit_size_q,
      TO_JSON_STRING(APPROX_QUANTILES(max_received_amount, 100)) AS max_deposit_amount_q,
      TO_JSON_STRING(APPROX_QUANTILES(payment_channel_count, 100)) AS payment_channel_count_q,
      TO_JSON_STRING(APPROX_QUANTILES(timeout_counts / NULLIF(deposit_tries, 0), 100)) AS deposit_timeout_rate_q,
      TO_JSON_STRING(APPROX_QUANTILES(complete_deposit_counts / NULLIF(deposit_tries, 0), 100)) AS deposit_success_rate_q,
      TO_JSON_STRING(APPROX_QUANTILES(success_wd_count, 100)) AS success_wd_count_q,
      TO_JSON_STRING(APPROX_QUANTILES(success_wd_amount, 100)) AS success_wd_amount_q,
      TO_JSON_STRING(APPROX_QUANTILES(avg_session_withdraw_time_120m, 100)) AS session_time_q,
      TO_JSON_STRING(APPROX_QUANTILES(ultra_fast_sessions / NULLIF(total_sessions, 0), 100)) AS quick_withdrawal_rate_q,
      TO_JSON_STRING(APPROX_QUANTILES(promo_amount, 100)) AS promo_amount_q,
      TO_JSON_STRING(APPROX_QUANTILES(promo_amount / NULLIF(deposit_amount, 0), 100)) AS dp_promo_amount_q,
      TO_JSON_STRING(APPROX_QUANTILES(deposit_frequency, 100)) AS deposit_frequency_q,
      TO_JSON_STRING(APPROX_QUANTILES(withdrawal_frequency, 100)) AS withdrawal_frequency_q,
      TO_JSON_STRING(APPROX_QUANTILES(promo_frequency, 100)) AS promo_frequency_q,
      TO_JSON_STRING(APPROX_QUANTILES(success_wd_count / NULLIF(complete_deposit_counts, 0), 100)) AS dp_wd_ratio_q,
      TO_JSON_STRING(APPROX_QUANTILES(avg_daily_bet_cnt, 100)) AS avg_daily_bet_session_q,
      TO_JSON_STRING(APPROX_QUANTILES(total_playAmt, 100)) AS total_playAmt_q,
      TO_JSON_STRING(APPROX_QUANTILES(total_earnAmt, 100)) AS total_earnAmt_q,
      TO_JSON_STRING(APPROX_QUANTILES(validbet_deposit_ratio, 100)) AS validbet_deposit_ratio_q,
      TO_JSON_STRING(APPROX_QUANTILES(ggr, 100)) AS ggr_q,
      TO_JSON_STRING(APPROX_QUANTILES(ngr, 100)) AS ngr_q,
      TO_JSON_STRING(APPROX_QUANTILES(win_loss_ratio, 100)) AS win_loss_ratio_q
    FROM risk_metrics rm
    WHERE
      rm.brand_count <> 0
      AND (
          rm.complete_deposit_counts > 10 
          OR (rm.deposit_tries > 10 AND rm.complete_deposit_counts / rm.deposit_tries <= 0.1)
          OR (brand_registration_rate_per_day >= 10 AND SAFE_DIVIDE(rm.total_playAmt, rm.deposit_amount) <= 1)
      );`;

    Logger.log(PERCENTILE_QUERY);

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
      "registerRatePerDay", "brandCount", "dpTries", "successDpCnt", "dpAmt", "avgDpAmt", "maxDp", "dpChannelCnt", "dpTimeoutRatio", "dpSuccessRatio",
      "successWdCnt", "wdAmount", "avgSessionTime", "quickWDRatio", "promoAmt", "Promo/Deposit",
      "dpFreq", "wdFreq", "promoFreq", "WD x DP", "dailyBetSession", "ValidBet", "earnAmt", "ValidBet/dpAmt",
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
    rawdatasheet.getRange("B9:I15").setNumberFormat("#,##0.00");  // Cols B 
    rawdatasheet.getRange("C9:E15").setNumberFormat("#,##0");  // Cols C to E
    rawdatasheet.getRange("I9:I15").setNumberFormat("#,##0");  // Cols I
    rawdatasheet.getRange("L9:L15").setNumberFormat("#,##0");  // Cols L
    rawdatasheet.getRange("M9:N15").setNumberFormat("#,##0.00");  // Cols L to N
    rawdatasheet.getRange("P9:P15").setNumberFormat("#,##0.00");  // Col P
    rawdatasheet.getRange("R9:AA15").setNumberFormat("#,##0.00"); // Cols R to AA

    // Percentage columns conversion rules (Shifted Right)
    rawdatasheet.getRange("J9:J15").setNumberFormat("0.00%");    // Col J: dpTimeoutRatio
    rawdatasheet.getRange("K9:K15").setNumberFormat("0.00%");    // Col K: dpSuccessRatio
    rawdatasheet.getRange("O9:O15").setNumberFormat("0.00%");    // Col O: quickWithdrawalRate
    rawdatasheet.getRange("Q9:Q15").setNumberFormat("0.00%");    // Col Q: promoDepositRatio
    rawdatasheet.getRange("AB9:AB15").setNumberFormat("0.00%");  // Col AB: Win/Loss
    
    rawdatasheet.getRange("A2").setValue(`Updated: ${new Date()}`);
    rawdatasheet.setFrozenRows(8);
    statusCell.setValue(`✅ Finished at ${new Date()}`).setFontColor("#38761d");
    
    reportSheet.getRange("C4").setValue(`Finished calculating percentile at ${new Date()}, Now getting raw data`); 

  } catch (error) {
    Logger.log(error);
    statusCell.setValue(`❌ Error: ${error.message}`).setFontColor("#FF0000");
    throw error;
  }
}