function addInformation() {
  const projectId = "m2m-core";
  const queryExecute = `
    SELECT id_on_ota, ota_type, start_date, end_date
    FROM \`m2m-core.m2m_core_prod.reservation\`
    WHERE PARSE_DATE('%Y-%m-%d', end_date) BETWEEN DATE_SUB(CURRENT_DATE(), INTERVAL 3 MONTH) AND DATE_ADD(CURRENT_DATE(), INTERVAL 6 MONTH)
  `;
  Logger.log(queryExecute);

  const ss = SpreadsheetApp.openById("1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ");
  const outputSheet = ss.getSheetByName("addinformation");
  outputSheet.setFrozenRows(1);
  const headers = [["id_on_ota", "ota_type", "start_date", "end_date"]];
  outputSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  const request = {
    query: queryExecute,
    useLegacySql: false
  };

  try {
    const queryResults = BigQuery.Jobs.query(request, projectId);

    if (queryResults.jobComplete) {
      if (queryResults.rows && queryResults.rows.length > 0) {
        const rows = queryResults.rows.map(row => row.f.map(cell => cell.v));
        Logger.log(rows);
        outputSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      } else {
        Logger.log("No data returned.");
      }
    } else {
      Logger.log("The job is not complete.");
    }
  } catch (e) {
    Logger.log(`Error: ${e.message}`);
  }
}

