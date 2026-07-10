function run_sus_phone_email_with_lf_stats() {

  run_percentile_phone_email();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("Suspicious_Player");
  const rawdatasheet = ss.getSheetByName("Raw_Lt_phone_email");
  const statusCell = reportSheet.getRange("C5"); 
  
  try{
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
    // Given metrics
    // -----------------------------------

    const givenDPTries          = reportSheet.getRange("B7").getValue().toString().trim();
    const givenDPSuccessRatio   = reportSheet.getRange("B8").getValue().toString().trim() / 100;
    const givenTurnover         = reportSheet.getRange("B9").getValue().toString().trim();
    const givenRegistrationRate = reportSheet.getRange("B10").getValue().toString().trim();
    const givenBrandCount       = reportSheet.getRange("B11").getValue().toString().trim();
    
    const USER_LF_QUERY = `
        WITH excluded_groups AS (
            ${getGroupIds}
        )
        , brand_all AS (
            SELECT Brand, \`group\`, Sub_group, whitelabel, 1 AS priority, Country
            FROM \`${projectId}.MAPPING.brand_whitelabel_country_folderid_mapping_tbl\`
            -- 1. Change the countries / group here
            WHERE Whitelabel LIKE 'KZ%'
            UNION ALL
            SELECT  Brand, \`group\`, Sub_group, whitelabel, 2 AS priority, Country
            FROM \`${projectId}.MAPPING.brand_whitelabel_country_folderid_mapping_tbl\`
            -- 1. Change the countries / group here
            WHERE Whitelabel NOT LIKE 'KZ%'
        )
        , brand_dedup AS (
            SELECT
                Brand, 
                Country
            FROM brand_all
            QUALIFY ROW_NUMBER() OVER ( PARTITION BY Brand ORDER BY priority ASC ) = 1
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
        ,phone_email_detail AS (
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
        , adj_balance AS (
        SELECT
            m.phone_email,
            SUM(f.netAmount) AS adj_balance_amount
        FROM \`${projectId}.crm_gold_prod.funding_tx_daily\` f
        INNER JOIN members m ON f.memberId = m.memberID 
        WHERE f.insertedDate >= '${fromBQ}'
            AND f.type = 'adjustment'
            AND f.country = '${country_prefix}'
            AND status = 'completed'
            AND f.createdAt BETWEEN '${fromBQ}' AND '${toBQ}'            
        GROUP BY ALL
        )
        ,phone_email_agg AS (
        SELECT 
            ped.phone_email,
            ped.earliest_registered_date,
            ped.brand_registration_rate_per_day,
            ped.brand_count,
            ped.username_list,
            ped.brand_list,
            ab.adj_balance_amount
        FROM phone_email_detail ped
        LEFT JOIN adj_balance ab ON ped.phone_email = ab.phone_email
        )
        , user_lifetime_stats AS (
        SELECT 
            m.phone_email,
            SUM(total_deposits) as lf_deposit_amount, 
            SUM(total_withdrawals) as lf_withdrawal_amount, 
            SUM(total_dp_minus_wd) as lf_net_deposit_amount, 
            SUM(deposit_count) as lf_complete_deposit_counts, 
            SUM(failed_deposit_count) as lf_failed_deposit_counts, 
            SUM(withdrawal_count) AS lf_total_withdrawal_count, 
            SUM(claimed_bonuses) AS lf_total_claimed_bonuses, 
            SUM(total_valid_bets) AS lf_total_valid_bets, 
            SUM(player_wl_ggr + total_valid_bets) as lf_total_win,
            SUM(player_wl_ggr) AS lf_total_player_wl_ggr
        FROM \`${projectId}.crm_gold_prod.member_lifetime_stats_latest\` f
        INNER JOIN members m ON f.member_Id = m.memberID 
        INNER JOIN brand_dedup b ON m.brand = b.Brand
        WHERE m.phone_email IS NOT NULL AND b.country = '${country_prefix}'
        GROUP BY m.phone_email
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
            COUNT(DISTINCT IF(type = 'deposit' AND status IN ('error', 'timeout'), order_id, NULL)) AS error_timeout_counts,
            SUM(IF(type = 'deposit' AND status = 'completed', CAST(amount AS FLOAT64), 0.0)) AS deposit_amount,
            COUNT(DISTINCT IF(type = 'deposit' AND status = 'completed'  AND CAST(amount AS FLOAT64) < 100, order_id, NULL )) AS mini_deposit_amt_cnt,
            
            -- --- WITHDRAW METRICS ---
            COUNT(DISTINCT IF(type = 'withdraw', order_id, NULL)) AS wd_tries,
            COUNT(DISTINCT IF(type = 'withdraw' AND status = 'completed', order_id, NULL)) AS success_wd_count,
            SUM(IF(type = 'withdraw' AND status = 'completed', CAST(amount AS FLOAT64), 0.0)) AS success_wd_amount,

            -- --- PROMOTION METRICS ---
            COUNT(DISTINCT IF(type = 'promotion' AND status = 'completed', order_id, NULL)) AS redeem_promo_count,
            SUM(IF(type = 'promotion' AND status = 'completed', CAST(amount AS FLOAT64), 0.0)) AS promo_amount,

            -- Cleaned and optimized sorting evaluation block
            -- STRING_AGG(
            -- DISTINCT IF(type = 'promotion' AND status = 'completed', cleaned_remark, NULL), ',' 
            -- ORDER BY IF(type = 'promotion' AND status = 'completed', cleaned_remark, NULL) ASC
            -- ) AS cleaned_remark_list,
            

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
            SUM(error_timeout_counts) AS error_timeout_counts,
            SUM(deposit_amount) AS deposit_amount,
            SUM(mini_deposit_amt_cnt) AS mini_deposit_amt_cnt,
            
            SUM(wd_tries) AS wd_tries,
            SUM(success_wd_count) AS success_wd_count,
            SUM(success_wd_amount) AS success_wd_amount,
            
            SUM(redeem_promo_count) AS redeem_promo_count,
            SUM(promo_amount) AS promo_amount,
            --  STRING_AGG(DISTINCT cleaned_remark_list, ',') AS cleaned_remark_list,
            
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
            ar.country,
            pea.phone_email,
            pea.brand_count,
            pea.earliest_registered_date,
            pea.brand_registration_rate_per_day,

            -- 💰 Deposit
            ar.payment_channel_count,
            ar.deposit_tries,
            ar.complete_deposit_counts,
            ar.timeout_counts,
            ar.error_timeout_counts,
            ar.deposit_amount,
            ar.mini_deposit_amt_cnt,

            -- 💸 Withdrawal
            ar.wd_tries,
            ar.success_wd_count,
            ar.success_wd_amount,

            -- 🎁 Promotion
            ar.redeem_promo_count,
            ar.promo_amount,
            -- ar.cleaned_remark_list,
            
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
            
            -- ⏱️ Session Metrics    
            sm.avg_session_withdraw_time_120m,
            sm.ultra_fast_sessions,
            sm.total_sessions,
            (t.total_playAmt - t.total_earnAmt) AS ggr,
            (t.total_playAmt - t.total_earnAmt - ar.promo_amount) AS ngr,

            -- Lifetime Metrics
            uls.lf_deposit_amount, 
            uls.lf_withdrawal_amount, 
            uls.lf_net_deposit_amount, 
            uls.lf_complete_deposit_counts, 
            uls.lf_failed_deposit_counts, 
            uls.lf_total_withdrawal_count AS lf_total_withdrawal_count, 
            uls.lf_total_claimed_bonuses AS lf_total_claimed_bonuses, 
            uls.lf_total_valid_bets AS lf_total_valid_bets, 
            uls.lf_total_win AS lf_total_win,
            uls.lf_total_player_wl_ggr AS lf_total_player_wl_ggr,
            pea.adj_balance_amount AS lf_adjustment_amount,
            SAFE_DIVIDE(uls.lf_total_valid_bets , uls.lf_deposit_amount) AS lf_turnover,

            -- Detailed User Info
            aua.active_brand_list,
            aua.active_username_list,
            pea.username_list,
            pea.brand_list

        FROM phone_email_agg pea
        LEFT JOIN agg_transactions_phone ar ON pea.phone_email = ar.phone_email
        LEFT JOIN turnover_phone t ON pea.phone_email = t.phone_email
        LEFT JOIN session_metrics_phone sm ON pea.phone_email = sm.phone_email
        LEFT JOIN active_user_agg aua ON pea.phone_email = aua.phone_email
        LEFT JOIN user_lifetime_stats uls ON pea.phone_email = uls.phone_email
        )
        SELECT
            -- 👤 User Info
            rm.country,
            rm.phone_email,
            rm.brand_count,
            rm.earliest_registered_date,
            rm.brand_registration_rate_per_day,

            -- 💰 Deposit
            rm.payment_channel_count,
            rm.deposit_tries,
            rm.complete_deposit_counts,
            rm.timeout_counts,
            rm.error_timeout_counts,
            rm.deposit_amount,
            rm.avg_deposit_size,

            -- 💸 Withdrawal
            rm.wd_tries,
            rm.success_wd_count,
            rm.success_wd_amount,

            -- 🎁 Promotion
            rm.redeem_promo_count,
            rm.promo_amount,
            -- rm.cleaned_remark_list,
            rm.mini_deposit_amt_cnt,
            
            -- ⚡ Fret_frequency,
            rm.withdrawal_frequency,
            rm.promo_frequency,

            -- 🎰 Betting
            rm.total_playAmt,
            rm.total_earnAmt,
            rm.turnover,
            rm.ggr,
            rm.ngr,
            rm.win_loss_ratio,

            -- ⏱️ Session Metrics    
            rm.avg_session_withdraw_time_120m,
            rm.ultra_fast_sessions,
            rm.total_sessions,

            -- Lifetime Metrics
            rm.lf_deposit_amount,
            rm.lf_withdrawal_amount, 
            rm.lf_net_deposit_amount,
            rm.lf_adjustment_amount,
            rm.lf_complete_deposit_counts,
            rm.lf_failed_deposit_counts,
            rm.lf_total_withdrawal_count,
            rm.lf_total_claimed_bonuses,
            rm.lf_total_valid_bets,
            rm.lf_total_win,
            rm.lf_total_player_wl_ggr,

            -- Detailed User Info
            rm.active_brand_list,
            rm.active_username_list,
            rm.brand_list,
            rm.username_list
        FROM risk_metrics rm
        WHERE rm.brand_count <> 0
        AND rm.win_loss_ratio < 0.1
        AND (
            (
                rm.turnover < ${givenTurnover} 
                AND (rm.complete_deposit_counts > 10 
                    OR (rm.deposit_tries > ${givenDPTries} AND rm.complete_deposit_counts / rm.deposit_tries <= ${givenDPSuccessRatio})
                    OR (rm.brand_registration_rate_per_day >= ${givenRegistrationRate})
                    OR rm.brand_count >= ${givenBrandCount}
                    )
            )
            OR 
            (
                rm.complete_deposit_counts > 10 
                AND mini_deposit_amt_cnt >= 1
            )
        );
    `;

    // Logger.log(USER_LF_QUERY);

    const lineChunkSize = 8000;
    for (let i = 0; i < USER_LF_QUERY.length; i += lineChunkSize) {
        Logger.log(USER_LF_QUERY.substring(i, i + lineChunkSize));
    }

    rawdatasheet.getDataRange().breakApart();
    rawdatasheet.getRange(2, 1, rawdatasheet.getMaxRows(), rawdatasheet.getMaxColumns()).clearContent();
    rawdatasheet.getRange("A1").setValue(`Updated: ${new Date()}`).setFontWeight("bold");
    statusCell.setValue(`Started running at ${new Date()}`);
    statusCell.setFontColor("#38761d"); // green
    SpreadsheetApp.flush(); // force write immediately
    
    Logger.log(`Running reconciliation for date range: ${fromBQ} to ${toBQ}`);

    const header = [
        // 👤 User Info
        "country", "phone/email", "brandCount", "earliestRegisteredDate", "registerRatePerDay", 

        // 💰 Deposit - Raw
        "dpChannelCnt", "dpTries", "successDpCnt", "timeoutDp", "errorTimeoutDp", "dpAmt", "avgDpAmt", 
        
        // 💰 Deposit - Rates (Calculated or from SQL)
        "dpSuccessRate(%)", "dpTimeout(%)", "dpTimeoutError(%)",

        // 💸 Withdrawal
        "wdTries", "successWdCnt", "wdAmt",

        // 🎁 Promotion
        "promoCnt", "promoAmt", "miniDepAmtCnt", "Promo/Deposit",

        // ⚡ Velocity
        "dpFreq", "wdFreq", "promoFreq", 
        "wdDpRatio",

        // 🎰 Betting & Ratios
        "playAmt", "earnAmt", "turnover", "GGR", "NGR", "Win/Loss(%)", 

        // ⏱️ Session Metrics
        "avg_session_withdraw_time_120m", "ultra_fast_sessions", "total_sessions", "utlra_fast_session_ratio", 

        // ⏳ Lifetime Metrics 
        "ltDpAmt", "ltWdAmt", "ltNetDpAmt", "ltAdjAmt", "ltSuccessDpCnt", "ltFailDpCnt", 
        "ltWdCnt", "ltClaimPromo", "ltValidBet", "ltEarnAmt", "ltTurnover",
        "lfPlayerWlGgr",

        // 📋 Detailed User Info / Appendices
        "activeBrandList", "activeUsernameList", "brandList", "usernameList"
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

    runQueryStream(USER_LF_QUERY, dp_val => {
        // 1. Mapping based on your EXACT SQL SELECT order
        // 👤 User Info
        const country                    = dp_val[0];
        const phoneEmail                 = dp_val[1];
        const brandCount                 = Number(dp_val[2]) || 0;
        const earliestRegisterDate       = dp_val[3];
        const registerRatePerDay         = Number(dp_val[4]) || 0;

        // 💰 Deposit
        const channelCnt                 = Number(dp_val[5]) || 0;
        const dpTries                    = Number(dp_val[6]) || 0;
        const successDp                  = Number(dp_val[7]) || 0;
        const timeoutDp                  = Number(dp_val[8]) || 0;
        const errorTimeoutDp             = Number(dp_val[9]) || 0;
        const dpAmount                   = Number(dp_val[10]) || 0;
        const avgDpAmount                = Number(dp_val[11]) || 0; // Directly from SQL (rm.avg_deposit_size)

        // 💸 Withdrawal
        const wdTries                    = Number(dp_val[12]) || 0;
        const successWdCnt               = Number(dp_val[13]) || 0;
        const wdAmount                   = Number(dp_val[14]) || 0;

        // 🎁 Promotion
        const promoCnt                   = Number(dp_val[15]) || 0;
        const promoAmt                   = Number(dp_val[16]) || 0;
        const miniDepAmtCnt               = dp_val[17] || "";

        // ⚡ Frequency
        const dpFreq                     = Number(dp_val[18]) || 0;
        const wdFreq                     = Number(dp_val[19]) || 0;
        const promoFreq                  = Number(dp_val[20]) || 0;

        // 🎰 Betting
        const playAmt                    = Number(dp_val[21]) || 0;
        const earnAmt                    = Number(dp_val[22]) || 0;
        const turnover                   = Number(dp_val[23]) || 0;
        const ggr                        = Number(dp_val[24]) || 0;
        const ngr                        = Number(dp_val[25]) || 0;
        const winLossRatio               = Number(dp_val[26]) || 0; // Directly from SQL (rm.win_loss_ratio)

        // ⏱️ Session Metrics    
        const avgSessionWithdrawTime120m = Number(dp_val[27]) || 0;
        const ultraFastSessions          = Number(dp_val[28]) || 0;
        const totalSessions              = Number(dp_val[29]) || 0;

        // ⏳ Lifetime Metrics
        const lfDepositAmount            = Number(dp_val[30]) || 0;
        const lfWithdrawalAmount         = Number(dp_val[31]) || 0;
        const lfNetDepositAmount         = Number(dp_val[32]) || 0;
        const lfAdjustmentAmount         = Number(dp_val[33]) || 0;
        const lfCompleteDepositCounts    = Number(dp_val[34]) || 0;
        const lfFailedDepositCounts      = Number(dp_val[35]) || 0;
        const lfTotalWithdrawalCount     = Number(dp_val[36]) || 0;
        const lfTotalClaimedBonuses      = Number(dp_val[37]) || 0;
        const lfTotalValidBets           = Number(dp_val[38]) || 0;
        const lfTotalWin                 = Number(dp_val[39]) || 0;
        const lfTotalPlayerWlGgr         = Number(dp_val[40]) || 0;

        // 📋 Detailed User Info / Appendices
        const activeBrandList            = dp_val[41];
        const activeUsernameList         = dp_val[42];
        const brandList                  = dp_val[43];
        const usernameList               = dp_val[44];

        // ==========================================
        // 2. Calculated Metrics (Handling values requested for calculations)
        // ==========================================
        const dpSuccessRate         = dpTries > 0 ? (successDp / dpTries) : 0;
        const dpTimeoutRate         = dpTries > 0 ? (timeoutDp / dpTries) : 0;
        const dpErrorTimeoutRate    = dpTries > 0 ? (errorTimeoutDp / dpTries) : 0;
        
        const promoDepositRatio     = dpAmount > 0 ? (promoAmt / dpAmount) : (promoAmt > 0 ? 1 : 0);
        const wdDpRatio             = dpAmount > 0 ? (wdAmount / dpAmount) : 0; // Financial value ratio matching header notation
        const ultraFastSessionRatio = totalSessions > 0 ? (ultraFastSessions / totalSessions) : 0;

        const lfTurnover            = lfDepositAmount > 0 ? (lfTotalValidBets / lfDepositAmount) : 0;

        // ==========================================
        // 3. Push to Buffer matching your FINALISED HEADER exactly
        // ==========================================
        buffer.push([
            // 👤 User Info
            country, 
            phoneEmail, 
            brandCount, 
            earliestRegisterDate, 
            registerRatePerDay, 

            // 💰 Deposit - Raw
            channelCnt, 
            dpTries, 
            successDp, 
            timeoutDp, 
            errorTimeoutDp, 
            dpAmount,  
            avgDpAmount, 
            
            // 💰 Deposit - Rates (Calculated)
            dpSuccessRate, 
            dpTimeoutRate, 
            dpErrorTimeoutRate,

            // 💸 Withdrawal
            wdTries, 
            successWdCnt, 
            wdAmount,

            // 🎁 Promotion
            promoCnt, 
            promoAmt, 
            miniDepAmtCnt, 
            promoDepositRatio,

            // ⚡ Velocity
            dpFreq, 
            wdFreq, 
            promoFreq, 
            wdDpRatio,

            // 🎰 Betting & Ratios
            playAmt, 
            earnAmt, 
            turnover, 
            ggr, 
            ngr, 
            winLossRatio, 

            // ⏱️ Session Metrics
            avgSessionWithdrawTime120m, 
            ultraFastSessions, 
            totalSessions, 
            ultraFastSessionRatio, 

            // ⏳ Lifetime Metrics 
            lfDepositAmount, 
            lfWithdrawalAmount, 
            lfNetDepositAmount, 
            lfAdjustmentAmount, 
            lfCompleteDepositCounts, 
            lfFailedDepositCounts, 
            lfTotalWithdrawalCount, 
            lfTotalClaimedBonuses, 
            lfTotalValidBets, 
            lfTotalWin, 
            lfTurnover,
            lfTotalPlayerWlGgr,

            // 📋 Detailed User Info / Appendices
            activeBrandList, 
            activeUsernameList, 
            brandList, 
            usernameList
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
    const formats = [
        // 👤 User Info
        { range: "C3:C",  fmt: "0" },                  // brandCount
        { range: "D3:D",  fmt: "yyyy-mm-dd" },         // earliestRegisteredDate
        { range: "E3:E",  fmt: "#,##0.00" },           // registerRatePerDay

        // 💰 Deposit - Raw
        { range: "F3:F",  fmt: "0" },                  // dpChannelCnt
        { range: "G3:G",  fmt: "0" },                  // dpTries
        { range: "H3:H",  fmt: "0" },                  // successDp
        { range: "I3:I",  fmt: "0" },                  // timeoutDp
        { range: "J3:J",  fmt: "0" },                  // errorTimeoutDp
        { range: "K3:K",  fmt: "#,##0.00" },           // dpAmt
        { range: "L3:L",  fmt: "#,##0.00" },           // avgDpAmt

        // 💰 Deposit - Rates (Calculated)
        { range: "M3:O",  fmt: "0.00%" },              // dp_success_rate(%), dp_timeout(%), dp_timeout_error(%)

        // 💸 Withdrawal
        { range: "P3:Q",  fmt: "0" },                  // wdTries, successWdCnt
        { range: "R3:R",  fmt: "#,##0.00" },           // wdAmt

        // 🎁 Promotion
        { range: "S3:S",  fmt: "0" },                  // promoCnt
        { range: "T3:U",  fmt: "#,##0.00" },           // promoAmt
        { range: "V3:V",  fmt: "0.00%" },              // Promo/Deposit ratio (U is promoRemarks string)

        // ⚡ Velocity
        { range: "W3:Y",  fmt: "0.00" },               // dpFreq, wdFreq, promoFreq
        { range: "Z3:Z",  fmt: "0.00" },               // wdDpRatio

        // 🎰 Betting & Ratios
        { range: "AA3:AB", fmt: "#,##0.00" },          // playAmt, earnAmt
        { range: "AC3:AC", fmt: "0.00" },              // turnover
        { range: "AD3:AE", fmt: "#,##0.00" },          // GGR, NGR
        { range: "AF3:AF", fmt: "0.00%" },             // Win/Loss(%)

        // ⏱️ Session Metrics
        { range: "AG3:AG", fmt: "0.00" },              // avg_session_withdraw_time_120m
        { range: "AH3:AI", fmt: "0" },                 // ultra_fast_sessions, total_sessions
        { range: "AJ3:AJ", fmt: "0.00%" },             // utlra_fast_session_ratio

        // ⏳ Lifetime Metrics 
        { range: "AK3:AM", fmt: "#,##0.00" },          // lf_deposit_amount, lf_withdrawal_amount, lf_net_deposit_amount
        { range: "AN3:AN", fmt: "#,##0.00" },          // lf_adjustment_amount
        { range: "AO3:AS", fmt: "0" },                 // lf_complete_dp, lf_failed_dp, lf_total_wd, lf_claimed_bonuses, lf_valid_bets
        { range: "AT3:AV", fmt: "#,##0.00" }           // lf_total_win, lfTotalPlayerWlGgr, lf_total_player_wl_ggr
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
    const sortCol1 = header.indexOf("dpAmt") + 1; 
    const sortCol2 = header.indexOf("turnover") + 1;

    rawdatasheet.getRange(3, 1, rawdatasheet.getLastRow() - 2, header.length).sort([
      { column: sortCol1, ascending: false }, 
      { column: sortCol2, ascending: true }
    ]);

    insert_categorised_lt_data();

    // 🚀 FORCE Google Sheets to finish writing all data to the cells NOW
    SpreadsheetApp.flush();
        
    statusCell.setValue(`✅ Finished fetching raw data from BQ at ${new Date()}`);

  } catch (error) {
    statusCell.setValue(`❌ Error: ${error.message}`);
    Logger.log("❌ Running failed: " + error.toString());
    
    // markTaskError(error);
    throw error;
  }finally{
    // markTaskComplete();
  }
}
