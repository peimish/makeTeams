function copySheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // 各シートから値をコピー
  var sheet = spreadsheet.getSheetByName("グループ分け");
  
  //非運転者列
  var non_drivers_last_row = counta(sheet, "A:A");
  //A列のラストまで取得
  var non_drivers_name_range = sheet.getRange(1, 1, non_drivers_last_row, 1);
  //B列のラストまで取得
  var non_drivers_group_range = sheet.getRange(1, 2, non_drivers_last_row, 2);
  //各グループの人数表を取得
  var non_drivers_group_table = sheet.getRange(1, 7, 2, 8);
  
  //運転者列
  var drivers_last_row = counta(sheet, "D:D");
  //D列のラストまで取得
  var drivers_name_range = sheet.getRange(1, 4, drivers_last_row, 4);
  //E列のラストまで取得
  var drivers_group_range = sheet.getRange(1, 5, drivers_last_row, 5);
  //各グループの人数表を取得
  var drivers_group_table = sheet.getRange(5, 7, 7, 8);
  
  //グループの合計人数表を取得
  var total_group_table = sheet.getRange(9, 7, 11, 8);
  
  //0からなので＋1、グループ分けは前の月に実行するだろうから更に+1
  var date = new Date();
  var next_month = date.getMonth()+1+1;
  //新しいシートを作成
  var new_sheet = spreadsheet.insertSheet(next_month + '月分');
  
  //条件付き書式だけコピペ
  non_drivers_name_range.copyTo(new_sheet.getRange(1,1), {formatOnly:true});
  non_drivers_group_range.copyTo(new_sheet.getRange(1,2), {formatOnly:true});
  drivers_name_range.copyTo(new_sheet.getRange(1,4), {formatOnly:true});
  drivers_group_range.copyTo(new_sheet.getRange(1,5), {formatOnly:true});
  drivers_group_table.copyTo(new_sheet.getRange(1, 7), {formatOnly:true});
  non_drivers_group_table.copyTo(new_sheet.getRange(5, 7), {formatOnly:true});
  total_group_table.copyTo(new_sheet.getRange(9, 7), {formatOnly:true});
  
  //値だけコピペ
  non_drivers_name_range.copyTo(new_sheet.getRange(1,1), {contentsOnly:true});
  non_drivers_group_range.copyTo(new_sheet.getRange(1,2), {contentsOnly:true});
  drivers_name_range.copyTo(new_sheet.getRange(1,4), {contentsOnly:true});
  drivers_group_range.copyTo(new_sheet.getRange(1,5), {contentsOnly:true});
  drivers_group_table.copyTo(new_sheet.getRange(1, 7), {contentsOnly:true});
  non_drivers_group_table.copyTo(new_sheet.getRange(5, 7), {contentsOnly:true});
  total_group_table.copyTo(new_sheet.getRange(9, 7), {contentsOnly:true});
  
  //セルの結合
  new_sheet.getRange("A1:B1").merge();
  new_sheet.getRange("D1:E1").merge();
  new_sheet.getRange("G1:H1").merge();
  new_sheet.getRange("G5:H5").merge();
  new_sheet.getRange("G9:H9").merge();
  
  //J~M列削除
  new_sheet.deleteColumns(10, 4);
}

function counta(sheet, range) {
  var range = sheet.getRange(range).getValues();
  Logger.log(range);
  var last_row_num = range.filter(String).length;
  Logger.log(last_row_num);
  return last_row_num;
}