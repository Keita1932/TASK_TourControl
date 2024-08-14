//Basic_informationシートのPhotoTourID、ListingID、CommonareaIDを取得
function basicInfo(){
  const ss = SpreadsheetApp.openById("1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ");
  const output_sheet = ss.getSheetByName("Basic_information")
  // function_1(ss,output_sheet)
  function_2(ss,output_sheet)
  function_3(ss,output_sheet)
  function_4(ss,output_sheet)
}

// function function_1(ss,output_sheet) {
//   //読み込みたいプロジェクトの名前
//   const project_id =  "m2m-core";
  
//   //実行するクエリ
//   const query_execute =  "SELECT id,name FROM `m2m_users_prod.user` where activation_status = 2 ORDER BY name"
//   //実行するクエリーの確認
//   Logger.log(query_execute)
//   //出力先シート名


//   output_sheet.setFrozenRows(1);
//   const result = BigQuery.Jobs.query(
//     {
//       useLegacySql: false,
//       query: query_execute,
//     },
//     project_id
//   );
//   const rows = result.rows.map(row => {
//     return row.f.map(cell => cell.v)
//   })
//   Logger.log(rows)
//  output_sheet
//  .getRange(
//    2, 
//    1, 
//    rows.length, 
//    rows[0].length
//  )
//  .setValues(rows);
// }

function function_2(ss,output_sheet){
  //読み込みたいプロジェクトの名前
  const project_id =  "m2m-core";
  
  //実行するクエリ
  const query_execute =  "SELECT id,title FROM `m2m_cleaning_prod.photo_tour` where disabled is false"
  //実行するクエリーの確認
  Logger.log(query_execute)

  const result = BigQuery.Jobs.query(
    {
      useLegacySql: false,
      query: query_execute,
    },
    project_id
  );
  const rows = result.rows.map(row => {
    return row.f.map(cell => cell.v)
  })
  Logger.log(rows)
 output_sheet
 .getRange(
   2, 
   1, 
   rows.length, 
   rows[0].length
 )
 .setValues(rows);
}

function function_3(ss,output_sheet) {
  //読み込みたいプロジェクトの名前
  const project_id =  "m2m-core";
  
  //実行するクエリ
  const query_execute =  "SELECT name,id FROM `m2m_core_prod.listing` ORDER BY name"
  //実行するクエリーの確認
  Logger.log(query_execute)

  const result = BigQuery.Jobs.query(
    {
      useLegacySql: false,
      query: query_execute,
    },
    project_id
  );
  const rows = result.rows.map(row => {
    return row.f.map(cell => cell.v)
  })
  Logger.log(rows)
 output_sheet
 .getRange(
   2, 
   4, 
   rows.length, 
   rows[0].length
 )
 .setValues(rows);
}


function function_4(ss,output_sheet) {
  //読み込みたいプロジェクトの名前
  const project_id =  "m2m-core";
  
  //実行するクエリ
  const query_execute =  "SELECT name,id,note FROM `m2m_core_prod.common_area_records` ORDER BY name"
  //実行するクエリーの確認
  Logger.log(query_execute)
   
  const result = BigQuery.Jobs.query(
    {
      useLegacySql: false,
      query: query_execute,
    },
    project_id
  );
  const rows = result.rows.map(row => {
    return row.f.map(cell => cell.v)
  })
  Logger.log(rows)
 output_sheet
 .getRange(
   2, 
   7, 
   rows.length, 
   rows[0].length
 )
 .setValues(rows);
}