function run_percentile_sus_phone_email() {

  percentile_calculator_phone_email();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("RT_Report_v4");
  const rawdatasheet = ss.getSheetByName("RT_Raw_Data_v4");
  const statusCell = reportSheet.getRange("C4"); // <-- status / error display
  
  try{
    // markTaskStart("Task fetching data from BQ");
    const fromDate = reportSheet.getRange("B3").getValue();
    const toDate = reportSheet.getRange("B4").getValue();
    const messageCell = reportSheet.getRange("C4");

    reportSheet.getRange("C5").clear();

    // Robust date validation
    if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
      messageCell.setValue("Please enter valid dates");
      return;
    }

    if (toDate < fromDate) {
      messageCell.setValue("To date must be later than From date");
      messageCell.setFontColor("#FF0000"); // red

      return;
    } else {
      messageCell.clearContent(); // Clear previous message 
    }

    // 1. Subtract 1 day from B3 (fromDate) Safely
    const adjustedFromDate = new Date(fromDate.getTime());
    // adjustedFromDate.setDate(adjustedFromDate.getDate() - 1);
    const fromBQ = formatDateForBQ(adjustedFromDate);

    // 2. Add 1 day to B4 (toDate) Safely
    const adjustedToDate = new Date(toDate.getTime());
    adjustedToDate.setDate(adjustedToDate.getDate() + 1);
    const toBQ = formatDateForBQ(adjustedToDate);
    const getGroupIds = generateSqlExcludedGroups();


    // -----------------------------------
    // Percentile thresholds
    // -----------------------------------

    const p25_avgSessionTime      = reportSheet.getRange("P12").getValue(); // O -> P
    const p90_quickWDRatio        = reportSheet.getRange("Q15").getValue(); // P -> Q
    const p90_timeoutRatio        = reportSheet.getRange("L15").getValue(); // K -> L
    const p25_validBetDpRatio     = reportSheet.getRange("AA12").getValue(); // Z -> AA
    const p75_dpFreq              = reportSheet.getRange("T14").getValue(); // S -> T
    const p90_dpFreq              = reportSheet.getRange("T15").getValue(); // S -> T
    const p75_paymentChannelCount = reportSheet.getRange("K14").getValue(); // J -> K
    const p25_avgDpAmt            = reportSheet.getRange("I12").getValue(); // H -> I
    const p90_avgDpAmt            = reportSheet.getRange("I15").getValue(); // H -> I
    const p90_brandCount          = reportSheet.getRange("E15").getValue(); // D -> E
    const p90_wdDpRatio           = reportSheet.getRange("W15").getValue(); // V -> W

    const DEPOSIT_QUERY_RT = `
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
          -- 👤 User Info
          rm.country,
          rm.phone_email,
          rm.brand_count,
          rm.earliest_registered_date,
          rm.brand_registration_rate_per_day,
          rm.brand_list,  

          -- 💰 Deposit (Raw)
          rm.payment_channel_count,
          rm.deposit_tries,
          rm.complete_deposit_counts,
          rm.timeout_counts,
          rm.error_timeout_counts,
          rm.deposit_amount,
          rm.max_received_amount,

          -- 💸 Withdrawal
          rm.wd_tries,
          rm.success_wd_count,
          rm.success_wd_amount,

          -- Promo
          rm.redeem_promo_count,
          rm.promo_amount,
          rm.cleaned_remark_list,

          -- ⚡ Frequency
          rm.deposit_frequency,
          rm.withdrawal_frequency,
          rm.promo_frequency,

          -- 🎰 Betting
          COALESCE(rm.total_bet_cnt, 0) AS total_bet_cnt,
          ROUND(COALESCE(rm.avg_daily_bet_cnt, 0), 2) AS avg_daily_bet_cnt,
          total_playAmt,
          total_earnAmt,
          max_playAmt,
          
          -- Sessions
          rm.avg_session_withdraw_time_120m,
          rm.ultra_fast_sessions,
          rm.total_sessions,

          rm.username_list,
          rm.active_username_list,
          rm.active_brand_list
          
      FROM risk_metrics rm
      WHERE
        rm.brand_count <> 0
        AND (
            (
              -- Multiple brand registered within a week
              rm.brand_count >= ${p90_brandCount}
              AND brand_registration_rate_per_day >= 10
              AND SAFE_DIVIDE(rm.total_playAmt, rm.deposit_amount) <= 1
            )
          OR
              (
                -- Payment Channel Scraping
                rm.complete_deposit_counts > 10 
                AND rm.deposit_frequency >= ${p90_dpFreq}
                AND rm.payment_channel_count >= ${p75_paymentChannelCount}
                AND SAFE_DIVIDE(rm.deposit_amount, rm.complete_deposit_counts) <= ${p25_avgDpAmt}
                AND SAFE_DIVIDE(rm.total_playAmt, rm.deposit_amount) <= ${p25_validBetDpRatio}
              )

          OR

            (
              -- PGW Crawlers
              rm.deposit_tries >= 10
              AND SAFE_DIVIDE(rm.complete_deposit_counts, rm.deposit_tries) <= 0.10
              AND SAFE_DIVIDE(rm.timeout_counts, rm.deposit_tries) >= ${p90_timeoutRatio}
              AND SAFE_DIVIDE(total_playAmt, deposit_amount) <= ${p25_validBetDpRatio}
            )

          OR

            (
              -- PR Extortion
              rm.complete_deposit_counts > 10 
              AND rm.brand_count >= ${p90_brandCount}
              AND (
                  SAFE_DIVIDE(rm.ultra_fast_sessions, rm.total_sessions) >= ${p90_quickWDRatio}
                  OR rm.avg_session_withdraw_time_120m <= ${p25_avgSessionTime}
              )
              AND SAFE_DIVIDE(total_playAmt, deposit_amount) <= ${p25_validBetDpRatio}
              AND SAFE_DIVIDE(ABS(deposit_amount - success_wd_amount), deposit_amount) BETWEEN -0.10 AND 0.10
            )

          OR

            (
              -- Rapid DP-WD Testing
              rm.complete_deposit_counts > 10 
              AND 
              (
                SAFE_DIVIDE(rm.ultra_fast_sessions, rm.total_sessions) >= ${p90_quickWDRatio}
                OR rm.avg_session_withdraw_time_120m <= ${p25_avgSessionTime}
              )
              AND SAFE_DIVIDE(total_playAmt, deposit_amount) <= ${p25_validBetDpRatio}
              AND SAFE_DIVIDE(success_wd_count, complete_deposit_counts) >= ${p90_wdDpRatio}
            )
        )
    `;

    Logger.log(DEPOSIT_QUERY_RT);

    rawdatasheet.getDataRange().breakApart();
    rawdatasheet.getRange(2, 1, rawdatasheet.getMaxRows(), rawdatasheet.getMaxColumns()).clearContent();
    rawdatasheet.getRange("A1").setValue(`Updated: ${new Date()}`).setFontWeight("bold");
    statusCell.setValue(`Started running at ${new Date()}`);
    statusCell.setFontColor("#38761d"); // green
    SpreadsheetApp.flush(); // force write immediately
    
    Logger.log(`Running reconciliation for date range: ${fromBQ} to ${toBQ}`);

   const header = [
    // 👤 User Info
    "country", "phone/email", "brandCount", "earliest_registered_date", "registerRatePerDay", "brand_list", 

    // 💰 Deposit - Raw
    "dpChannelCnt", "dpTries", "successDp", "timeoutDp", "errorTimeoutDp", "dpAmount", "maxDp", 

    // 💰 Deposit - Rates (Calculated or from SQL)
    "dp_success_rate(%)", "dp_timeout(%)", "dp_timeout_error(%)", "avgDpAmount", 

    // 💸 Withdrawal
    "wdTries", "successWdCnt", "wdAmount",

    // Promo
    "promoCnt", "promoAmt", "promoRemarks", "Promo/Deposit",

    // ⚡ Velocity
    "dpFreq", "wdFreq", "promoFreq", "wdDpRatio",

    // 🎰 Betting
    "totalBetCnt", "avgDailyBetCnt", "playAmt", "earnAmt", "maxPlayAmt", "playAmt/dpAmt", "GGR", "NGR", "Win/Loss(%)",

    // 📊 Sessions
    "avg_session_withdraw_time_120m", "ultra_fast_sessions", "total_sessions", "utlra_fast_session_ratio",

    // 📋 Lists / Appendices
    "usernameList","activeUsernameList", "activeBrandList"
  ];

    rawdatasheet.getRange(2, 1, 1, header.length)
    .setValues([header])
    .setFontWeight("bold")
    .setBackground("#D3D3D3")
    .setHorizontalAlignment("center");

    let writeRow = 3;
    const CHUNK = 4000;
    let buffer = [];

    const flush = () => {
      if (!buffer.length) return;
      rawdatasheet.getRange(writeRow, 1, buffer.length, buffer[0].length).setValues(buffer);
      writeRow += buffer.length;
      buffer = [];
      };

    statusCell.setValue(`Connecting to BQ and process rows at ${new Date()}`);
    SpreadsheetApp.flush(); // force write immediately

    runQueryStream(DEPOSIT_QUERY_RT, dp_val => {
      // 1. Mapping based on your EXACT SQL SELECT order
      // 👤 User Info
      const country              = dp_val[0];
      const phoneEmail           = dp_val[1];
      const brandCount           = Number(dp_val[2]) || 0;
      const earliestRegisterDate = dp_val[3];
      const registerRatePerDay   = Number(dp_val[4]) || 0;
      const brandList            = dp_val[5]; 

      // 💰 Deposit (Raw)
      const channelCnt           = Number(dp_val[6]) || 0;
      const dpTries              = Number(dp_val[7]) || 0;
      const successDp            = Number(dp_val[8]) || 0;
      const timeoutDp            = Number(dp_val[9]) || 0;
      const errorTimeoutDp       = Number(dp_val[10]) || 0;
      const dpAmount             = Number(dp_val[11]) || 0;
      const maxDp                = Number(dp_val[12]) || 0;

      // 💸 Withdrawal
      const wdTries              = Number(dp_val[13]) || 0;
      const successWdCnt         = Number(dp_val[14]) || 0;
      const wdAmount             = Number(dp_val[15]) || 0;

      // Promo
      const promoCnt             = Number(dp_val[16]) || 0;
      const promoAmt             = Number(dp_val[17]) || 0;
      const promoRemarks         = dp_val[18] || "";

      // ⚡ Frequency
      const dpFreq               = Number(dp_val[19]) || 0;
      const wdFreq               = Number(dp_val[20]) || 0;
      const promoFreq            = Number(dp_val[21]) || 0;

      // 🎰 Betting
      const totalBet             = Number(dp_val[22]) || 0;
      const avgDailyBet          = Number(dp_val[23]) || 0;
      const playAmt              = Number(dp_val[24]) || 0;
      const earnAmt              = Number(dp_val[25]) || 0;
      const maxPlayAmt           = Number(dp_val[26]) || 0;
      
      // 📊 Sessions
      const avgSessionWithdrawTime120m = Number(dp_val[27]) || 0;
      const ultraFastSessions    = Number(dp_val[28]) || 0;
      const totalSessions        = Number(dp_val[29]) || 0;

      // 📋 Lists
      const usernameList         = dp_val[30];
      const activeusernameList   = dp_val[31];
      const activebrandlist      = dp_val[32];

      // 2. Calculated Metrics (Logic remains the same, using new variable assignments)
      const avgDpSize             = successDp > 0 ? dpAmount / successDp : 0;
      const dpSuccessRate         = dpTries > 0 ? successDp / dpTries : 0;
      const dpTimeoutRate         = dpTries > 0 ? timeoutDp / dpTries : 0;
      const dpErrorRate           = dpTries > 0 ? errorTimeoutDp / dpTries : 0;
      const wdDpRatio             = successDp > 0 ? successWdCnt / successDp : 0;
      const promoDpRatio          = (successDp > 0) ? (promoAmt / dpAmount) : (promoAmt > 0 ? 1 : 0);
      
      const ggr                   = playAmt - earnAmt;
      const ngr                   = playAmt - earnAmt - promoAmt;
      const playToDpRatio         = dpAmount > 0 ? playAmt / dpAmount : 0;
      const winLossPct            = dpAmount > 0 ? (dpAmount - wdAmount) / dpAmount : 0;
      const ultraFastSessionRatio = totalSessions > 0 ? ultraFastSessions / totalSessions : 0;

      // 3. Push to Buffer matching your FINALISED HEADER exactly
      buffer.push([
        // 👤 User Info
        country, 
        phoneEmail, 
        brandCount, 
        earliestRegisterDate, 
        registerRatePerDay, 
        brandList, 

        // 💰 Deposit - Raw
        channelCnt, 
        dpTries, 
        successDp, 
        timeoutDp, 
        errorTimeoutDp, 
        dpAmount, 
        maxDp, 

        // 💰 Deposit - Rates / Metrics
        dpSuccessRate, 
        dpTimeoutRate, 
        dpErrorRate, 
        avgDpSize, 

        // 💸 Withdrawal
        wdTries, 
        successWdCnt, 
        wdAmount,

        // Promo
        promoCnt, 
        promoAmt, 
        promoRemarks, 
        promoDpRatio,

        // ⚡ Velocity
        dpFreq, 
        wdFreq, 
        promoFreq, 
        wdDpRatio,

        // 🎰 Betting
        totalBet, 
        avgDailyBet, 
        playAmt, 
        earnAmt, 
        maxPlayAmt, 
        playToDpRatio,
        ggr, 
        ngr, 
        winLossPct,

        // 📊 Sessions
        avgSessionWithdrawTime120m, 
        ultraFastSessions, 
        totalSessions, 
        ultraFastSessionRatio,

        // 📋 Lists / Appendices
        usernameList, 
        activeusernameList,
        activebrandlist
      ]);

      if (buffer.length >= CHUNK) flush();
    });

    // flush remaining rows
    if (buffer.length > 0) {
      flush();
    }

    statusCell.setValue(`Finished fectching all rows from BQ at ${new Date()}`);
    SpreadsheetApp.flush(); // force write immediately

    rawdatasheet.setFrozenRows(2);
    
    const lastRow = rawdatasheet.getLastRow();
    const lastCol = rawdatasheet.getLastColumn();

    if (lastRow < 3) return; // Prevent errors if no data rows exist

    // Map your columns based on the FINALIZED HEADER order:
    // A: country, B: phone/email, C: brandCount, D: earliest_registered_date, 
    // E: registerRatePerDay, F: brand_list ... and so on.

    const formats = [
      // 👤 User Info
      { range: "C3:C",  fmt: "0" },                   // brandCount
      { range: "D3:D",  fmt: "yyyy-mm-dd" },          // earliest_registered_date
      { range: "E3:E",  fmt: "#,##0.00" },            // registerRatePerDay

      // 💰 Deposit - Raw
      { range: "G3:G",  fmt: "0" },                   // dpChannelCnt
      { range: "H3:K",  fmt: "0" },                   // dpTries, successDp, timeoutDp, errorTimeoutDp
      { range: "L3:L",  fmt: "#,##0.00" },            // dpAmount
      { range: "M3:M",  fmt: "#,##0.00" },            // maxDp 

      // 💰 Deposit - Rates & Metrics
      { range: "N3:P",  fmt: "0.00%" },               // dp_success_rate(%), dp_timeout(%), dp_timeout_error(%)
      { range: "Q3:Q",  fmt: "#,##0.00" },            // avgDpAmount

      // 💸 Withdrawal
      { range: "R3:S",  fmt: "0" },                   // wdTries, successWdCnt
      { range: "T3:T",  fmt: "#,##0.00" },            // wdAmount

      // Promo
      { range: "U3:U",  fmt: "0" },                   // promoCnt
      { range: "V3:V",  fmt: "#,##0.00" },            // promoAmt
      { range: "X3:X",  fmt: "0.00%" },               // Promo/Deposit ratio (W is promoRemarks)

      // ⚡ Velocity
      { range: "Y3:AA", fmt: "0.00" },                // dpFreq, wdFreq, promoFreq
      { range: "AB3:AB", fmt: "0.00" },               // wdDpRatio

      // 🎰 Betting
      { range: "AC3:AD", fmt: "0" },                  // totalBetCnt, avgDailyBetCnt
      { range: "AE3:AG", fmt: "#,##0.00" },           // playAmt, earnAmt, maxPlayAmt
      { range: "AH3:AH", fmt: "0.00" },               // playAmt/dpAmt
      { range: "AI3:AJ", fmt: "#,##0.00" },           // GGR, NGR
      { range: "AK3:AK", fmt: "0.00%" },              // Win/Loss(%)

      // 📊 Sessions
      { range: "AL3:AL", fmt: "0.00" },               // avg_session_withdraw_time_120m
      { range: "AM3:AN", fmt: "0" },                  // ultra_fast_sessions, total_sessions
      { range: "AO3:AO", fmt: "0.00%" }               // utlra_fast_session_ratio
    ];

    // Apply batch formatting
    formats.forEach(f => {
      // Use the specific range concatenated with lastRow to avoid formatting empty rows unnecessarily
      rawdatasheet.getRange(f.range.split(':')[0] + ":" + f.range.split(':')[1].replace(/\d/g, '') + lastRow)
                  .setNumberFormat(f.fmt);
    });

    // Center align everything for better readability
    rawdatasheet.getRange(3, 1, lastRow - 2, lastCol).setHorizontalAlignment("center");

    // --- SORTING ---
    // Skip header row (row 2) and sort rows 3 → lastRow
    // Instead of hardcoding '4', find the index dynamically
    const sortCol1 = header.indexOf("dpTries") + 1; 
    const sortCol2 = header.indexOf("dpAmount") + 1;

    rawdatasheet.getRange(3, 1, rawdatasheet.getLastRow() - 2, header.length).sort([
      { column: sortCol1, ascending: false }, 
      { column: sortCol2, ascending: false }
    ]);

    insert_hardcoded_categorised_Data_phone_email();
    insert_categorised_percentile_user();

    // 🚀 FORCE Google Sheets to finish writing all data to the cells NOW
    SpreadsheetApp.flush();
        
    statusCell.setValue(`✅ Finished running at ${new Date()}`);

  } catch (error) {
    statusCell.setValue(`❌ Error: ${error.message}`);
    statusCell.setFontColor("#FF0000"); // red
    Logger.log("❌ Running Deposit failed: " + error.toString());
    
    markTaskError(error);
    throw error;
  }finally{
    markTaskComplete();
  }
}


function insert_categorised_percentile_phone_email() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("RT_Report_v4");
  const rawdatasheet = ss.getSheetByName("RT_Raw_Data_v4");
  const resultSheet = ss.getSheetByName("User_Result"); // <-- Added User_Result Sheet
  const statusCell = reportSheet.getRange("C6");
  const userInfoStartRow = 21;

  try {
    const toDate = reportSheet.getRange("B4").getValue();
    const toDateTime = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate()).getTime();

    const lastRow = rawdatasheet.getLastRow();
    const lastCol = rawdatasheet.getLastColumn();
    if (lastRow < 3) { statusCell.setValue("❌ No raw data found"); return; }

    // --- STEP AA: LOAD EXCLUSIONS FROM USER_RESULT SHEET ---
    statusCell.setValue("⏳ Loading exclusions from User_Result...");
    
    const blacklistEmails = {};
    const monitorEmails = {}; // Key: Phone/Email, Value: "Monitored User"
    const demoEmails = {};
    const normalCooldownEmails = {}; // Key: Phone/Email, Value: Expiration Timestamp (2 weeks later)

    if (resultSheet) {
      const resultLastRow = resultSheet.getLastRow();
      if (resultLastRow >= 5) { // Row 4 is header, data starts at row 5
        // Fetch 5 columns: Date[0], Result[1], Category[2], Phone/Email[3], Username[4]
        const resultData = resultSheet.getRange(5, 1, resultLastRow - 4, 5).getValues();
        
        resultData.forEach(row => {
          const rowDateStr   = row[0];
          const resultStatus = String(row[1]).trim();
          const emailPhone   = String(row[3]).trim(); // Index 3 is now Phone/Email

          if (!emailPhone || emailPhone === "") return; // Skip if phone is empty

          if (resultStatus === "Blacklist") {
            blacklistEmails[emailPhone] = true;
          } 
          else if (resultStatus === "Demo") {
            demoEmails[emailPhone] = true;
          } 
          else if (resultStatus === "Monitor") {
            monitorEmails[emailPhone] = "Monitored User";
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

    statusCell.setValue("⏳ Fetching raw data...");
    const rawHeaders = rawdatasheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const rawData = rawdatasheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    const rawColMap = {};
    rawHeaders.forEach((h, i) => rawColMap[h.trim()] = i);
    
    // Included "status" dynamically at the end to handle "Monitor User" labeling
    const columnsToShow = [
      "category","phone/email", "earliestRegisteredDate", "registerRatePerDay", "brandCount","dpTries","successDpCnt", "dpAmt", "avgDpAmt", "maxDp", 
      "dpChannelCnt", "dpTimeoutRatio","dpSuccessRatio", "successWdCnt", "wdAmt", "avgSessionTime", "quickWDRatio","promoAmt", "Promo/Deposit", 
      "dpFreq", "wdFreq", "promoFreq", "WD x DP", "dailyBetCnt", "ValidBet", "earnAmt",
      "ValidBet/ dpAmt", "GGR", "NGR", "Win/Loss(%)", "activeBrandList", "activeUsernameList",  "brand_list","usernameList", "status"
    ];

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

    // --- STEP B: CATEGORIZATION ---
    statusCell.setValue("⏳ Categorizing rows...");

    const categoryCounts = { 
      "Payment Channel Scraping": 0, "PGW Crawlers" :0, "Rapid DP-WD Testing": 0, "PR Extortion": 0 , "High Brand Low Turnover":0, "Normal": 0 
    };

    const groupMap = getMemberGroupMap();

    const filtered = rawData.reduce((acc, row) => {
      const phoneEmail = String(row[rawColMap["phone/email"]] || "").trim();

      // 1. Check Blacklist
      if (blacklistEmails[phoneEmail]) return acc;

      // 1.2 Check Demo
      if (demoEmails[phoneEmail]) return acc;

      // 2. Check Normal Cooldown
      if (normalCooldownEmails[phoneEmail]) {
        if (toDateTime < normalCooldownEmails[phoneEmail]) return acc;
      }

      // 3. Determine Status Label (Monitor check)
      const monitorLabel = monitorEmails[phoneEmail] || "";

      const avgDep = parseFloat(row[rawColMap["avgDpAmount"]]) || 0;
      const playRatio = parseFloat(row[rawColMap["playAmt/dpAmt"]]) || 0;
      const chanCnt = parseInt(row[rawColMap["dpChannelCnt"]]) || 0;
      const brandCount = parseInt(row[rawColMap["brandCount"]]) || 0;
      const dFreq = parseFloat(row[rawColMap["dpFreq"]]) || 0;
      const successDpCnt = parseFloat(row[rawColMap["successDp"]]) || 0 ;
      const wdDpRatio = parseFloat(row[rawColMap["wdDpRatio"]]) || 0 ;
      const successWdCnt = parseFloat(row[rawColMap["successWdCnt"]]) || 0 ;
      const avgSessionTime = parseFloat(row[rawColMap["avg_session_withdraw_time_120m"]]) || 0 ;
      const timeoutStr  = row[rawColMap["dp_timeout(%)"]] || "0";

      const timeoutRatio = typeof timeoutStr === "string" ? parseFloat(timeoutStr.replace("%", "")) / 100 : parseFloat(timeoutStr) || 0;
      const dpSuccessRatioStr  = row[rawColMap["dp_success_rate(%)"]] || "0";
      const dpSuccessRatio = typeof dpSuccessRatioStr === "string" ? parseFloat(dpSuccessRatioStr.replace("%", "")) / 100 : parseFloat(dpSuccessRatioStr) || 0;
      const quickWDRatioStr  = row[rawColMap["utlra_fast_session_ratio"]] || "0";
      const quickWDRatio = typeof quickWDRatioStr === "string" ? parseFloat(quickWDRatioStr.replace("%", "")) / 100 : parseFloat(quickWDRatioStr) || 0;

      const dpTries = parseInt(row[rawColMap["dpTries"]]) || 0 ;
      const registerRatePerDay = parseFloat(row[rawColMap["registerRatePerDay"]]) || 0 ;
  
      let wlStr = row[rawColMap["Win/Loss(%)"]] || "0";
      let wlVal = typeof wlStr === "string" ? parseFloat(wlStr.replace("%", "")) / 100 : wlStr;

      const rawGroupIdStr = String(row[rawColMap["groupidList"]] || "");
      let hasMatchedGroup = false;
      let unmatchedGroups = "";

      if (rawGroupIdStr && rawGroupIdStr !== "0") {
        const groupIds = rawGroupIdStr.split(",").map(id => id.trim());
        hasMatchedGroup = groupIds.some(id => groupMap[id]);
        unmatchedGroups = "";
      }

      let cat = "Normal";
      if (dpTries >= 10 && dpSuccessRatio <= 0.1 && timeoutRatio >= TIMEOUTRATIO_P90 && playRatio <= TURNOVER_DP_P25){
        cat = 'PGW Crawlers';
      } else if (chanCnt >= PAYMENT_CHANNEL_CNT_P75 && avgDep <= AVG_DEP_P25 && dFreq >= DEP_FREQ_P90 && playRatio <= TURNOVER_DP_P25) {
        cat = "Payment Channel Scraping";
      } else if (brandCount >= BRAND_CNT_P90 && playRatio <= 1 && registerRatePerDay >= (REGISTER_RATE_MAX_DIV_2/2) ) {
        cat = "High Brand Low Turnover";
      } else if (brandCount >= BRAND_CNT_P90 && ((quickWDRatio >= QUICK_WD_P90 || avgSessionTime <= AVG_SESSION_TIME_P25) && (successDpCnt > 0 && successWdCnt > 0)) && playRatio <= TURNOVER_DP_P25 && wlVal >= WIN_LOSS_MIN && wlVal <= WIN_LOSS_MAX ) {
        cat = "PR Extortion";
      } else if (((quickWDRatio >= QUICK_WD_P90 || avgSessionTime <= AVG_SESSION_TIME_P25) && (successDpCnt > 0 && successWdCnt > 0)) && playRatio <= TURNOVER_DP_P25 && wdDpRatio >= WD_DP_RATIO_P90) {
        cat = "Rapid DP-WD Testing";
      } 
  
      if (hasMatchedGroup) {
        return acc;
      }

      categoryCounts[cat]++;

      if (cat !== "Normal") {
        acc.push(columnsToShow.map(h => {
          if (h === "category") return cat;
          if (h === "groupidList") return unmatchedGroups;
          if (h === "status") return monitorLabel; // Set monitor label to column 34

          return row[rawColMap[renameMap[h] || h]];
        }));
      }
      return acc;
    }, []);

    // --- STEP D: WRITE DATA & FORMATTING ---
    const currentMaxRows = reportSheet.getMaxRows();
    const numCols = columnsToShow.length;

    if (currentMaxRows >= userInfoStartRow) {
      reportSheet.getRange(userInfoStartRow, 1, currentMaxRows - (userInfoStartRow-1), numCols).clearContent().clearFormat();
      const existingFilter = reportSheet.getFilter();
      if (existingFilter) { existingFilter.remove(); }
    }
    SpreadsheetApp.flush();

    if (filtered && filtered.length > 0) {
      const outputRange = reportSheet.getRange(userInfoStartRow, 1, filtered.length, numCols);
      outputRange.setValues(filtered).setHorizontalAlignment("center");
      outputRange.offset(0, numCols - 1, filtered.length, 1).setHorizontalAlignment("left");
      
      const rowFormats = columnsToShow.map(h => formatMap[h] || "@");
      const fullTableFormats = Array(filtered.length).fill(rowFormats);
      outputRange.setNumberFormats(fullTableFormats);

      const getCol = (name) => {
        const idx = columnsToShow.indexOf(name);
        if (idx === -1) throw new Error(`Missing column: ${name}`);
        return idx + 1;
      };

      outputRange.sort([
        { column: getCol("ValidBet/ dpAmt"), ascending: true }
      ]);
      
      statusCell.setValue(`✅ Success: ${filtered.length} suspicious rows found.`);
    } else {
      statusCell.setValue(`✅ Success: 0 suspicious users detected.`);
      reportSheet.getRange(userInfoStartRow, 1).setValue("No suspicious activity detected.");
    }

    // --- STEP E: APPLY THIN VERTICAL BORDERS ---
    statusCell.setValue("⏳ Applying thin borders...");
    const dataBorderIndices = [3, 5, 13, 15, 17, 19, 23, 27, 30];

    const applyThinBorders = (range, indices) => {
      range.offset(0, 0, range.getHeight(), 1)
           .setBorder(null, true, null, null, null, null, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);

      indices.forEach(idx => {
        if (idx <= range.getWidth()) {
          range.offset(0, idx - 1, range.getHeight(), 1)
               .setBorder(null, null, null, true, null, null, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);
        }
      });
    };

    if (filtered && filtered.length > 0) {
      const mainDataRange = reportSheet.getRange(userInfoStartRow, 1, filtered.length, numCols);
      applyThinBorders(mainDataRange, dataBorderIndices);

      const last3Range = reportSheet.getRange(userInfoStartRow, numCols - 4, filtered.length, 5);
      last3Range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setHorizontalAlignment("left");
    }

    // --- STEP F: DYNAMIC SUMMARY TABLE (J2, M2, etc.) ---
    statusCell.setValue("⏳ Generating summary tables...");
    
    const summaryData = [
      ["Payment Channel Scraping", categoryCounts["Payment Channel Scraping"]],
      ["PGW Crawlers", categoryCounts["PGW Crawlers"]],
      ["PR Extortion", categoryCounts["PR Extortion"]],
      ["Rapid DP-WD Testing", categoryCounts["Rapid DP-WD Testing"]],
      ["High Brand Low Turnover", categoryCounts["High Brand Low Turnover"]],
      ["Normal", categoryCounts["Normal"]]
    ];

    let startCol = 10; 
    const rowsPerTable = 4;

    reportSheet.getRange("J2:O6").clearContent().clearFormat();

    for (let i = 0; i < summaryData.length; i += rowsPerTable) {
      const chunk = summaryData.slice(i, i + rowsPerTable);
      const tableHeader = [["Scenario", "# Users"]];
      const finalTable = tableHeader.concat(chunk);
      
      const destRange = reportSheet.getRange(2, startCol, finalTable.length, 2);
      destRange.setValues(finalTable);
      
      destRange.setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);
      destRange.offset(0, 0, finalTable.length, 1).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      destRange.offset(0, 1, finalTable.length, 1).setHorizontalAlignment("center").setVerticalAlignment("middle");
      
      const headerRange = destRange.offset(0, 0, 1, 2);
      headerRange.setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

      startCol += 3;
    }

    reportSheet.getRange("A2").setValue(`Updated: ${new Date()}`);
    statusCell.setValue(`✅ Success: ${filtered.length} suspicious rows found.`);

  } catch (err) {
    statusCell.setValue(`❌ Error: ${err.message}`);
  }
}
