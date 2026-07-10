const EXECUTION_PROJECT_ID = 'kz-dp-ops';
const projectId = "kz-dp-prod";

const country_prefix = 'TH';
const userOffset = 7;
const miniAmount = 100;


// function test_prod() {
//   try {
//     // This query runs using kz-dp-ops infrastructure, pointing to kz-dp-prod data
//     const BASE_QUERY_TEST = `
//       SELECT distinct country FROM \`kz-dp-prod.MAPPING.brand_whitelabel_country_folderid_mapping_tbl\`
//     `;

//     Logger.log("➡️ Starting environment test connection...");
//     Logger.log("SQL: " + BASE_QUERY_TEST);

//     let rowCount = 0;

//     // Stream results directly to the Apps Script Execution Log
//     runQueryStream(BASE_QUERY_TEST, row => {
//       rowCount++;
//       Logger.log(`Row [${rowCount}]: ${row[0]}`); // row[0] is the country column
//     });

//     Logger.log(`\n✅ SUCCESS: Environment connection test complete! Total rows found: ${rowCount}`);

//   } catch (error) {
//     Logger.log("\n❌ FAILURE: Environment test query failed!");
//     Logger.log("Error Details: " + error.toString());
//     throw error;
//   }
// }