function mersa_add_complete() {

  var mersa_google_sheets = SpreadsheetApp.getActiveSpreadsheet();
  var mersa_complete = mersa_google_sheets.getSheetByName("WOO_Complete");
  var mersa_orders = mersa_google_sheets.getSheetByName("WOO_Orders");

  const MERSA_MAX_COL = 20;
  const MERSA_NOTE_COL = 6;
  const MERSA_MAX_ROW = 2000;
  const MERSA_PROJECT_COL = 1;
  const MERSA_FNAME_COL = 2;
  const MERSA_LNAME_COL = 3;
  const MERSA_COMPANY_COL = 4;
  const MERSA_START_DATE_COL = 17;
  const MERSA_CONDITION_COL = 6;
  const MERSA_COMPLETE_DATE_COL = 19;
  const MERSA_BILLING_COL = 8;
  const MERSA_NUMBER_OF_ORDERS = mersa_get_number_of_filled_rows(mersa_orders, MERSA_PROJECT_COL, MERSA_MAX_ROW);
  const MERSA_CONDITION_COMPLETED = "completed";
  const MERSA_INDEX = [MERSA_MAX_COL, MERSA_NOTE_COL, MERSA_MAX_ROW, MERSA_PROJECT_COL, MERSA_FNAME_COL, MERSA_LNAME_COL, MERSA_COMPANY_COL, MERSA_START_DATE_COL, MERSA_CONDITION_COL, MERSA_COMPLETE_DATE_COL, MERSA_BILLING_COL, MERSA_NUMBER_OF_ORDERS, MERSA_CONDITION_COMPLETED];

  mersa_get_completed_orders(mersa_complete ,mersa_orders, MERSA_INDEX);

  var mersa_sheet_dest = mersa_complete.getRange(1,1,MERSA_MAX_ROW,MERSA_MAX_COL);

  mersa_sheet_dest.setFontFamily('Vazirmatn');
  mersa_sheet_dest.setFontSize(10);
  mersa_complete.getRange(1,1,1,MERSA_MAX_COL).setFontColor('White');
  mersa_complete.getRange(1,1,1,MERSA_MAX_COL).setBackground('Gray');
  mersa_sheet_dest.setVerticalAlignment('Middle');
  mersa_sheet_dest.setHorizontalAlignment('Center');
  mersa_sheet_dest.setWrap(true);


  setFilter(mersa_google_sheets, "WOO_Complete");
}


function setFilter(ss, sheet_name) {
  var filterSettings = {};

  filterSettings.range = {
    sheetId: ss.getSheetByName(sheet_name).getSheetId() // provide your sheetname to which you want to apply filter.
  };

  filterSettings.criteria = {};
  var columnIndex = 1; // column that defines criteria [A = 0]
  filterSettings['criteria'][columnIndex] = {
    'hiddenValues': ["FALSE"]
  };

  var request = {
    "setBasicFilter": {
      "filter": filterSettings
    }
  };
  
  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, ss.getId());
}


function makeArray(w, h, val) {
    var arr = [];
    for(let i = 0; i < h; i++) {
        arr[i] = [];
        for(let j = 0; j < w; j++) {
            arr[i][j] = val;
        }
    }
    return arr;
}


function mersa_get_number_of_filled_rows(sh, col, max_row) {
  let mersa_row_data = sh.getRange(1, col, max_row, 1).getValues();
  for (let counter = 1; counter <= max_row - 1; counter = counter + 1) {
    if(mersa_row_data[counter][0] === '')
      return counter - 1;
  }
  return max_row;
}


function makeArray(w, h, val) {
    var arr = [];
    for(let i = 0; i < h; i++) {
        arr[i] = [];
        for(let j = 0; j < w; j++) {
            arr[i][j] = val;
        }
    }
    return arr;
}


function mersa_get_completed_orders(sh_main ,sh_order, sh_index) {

//  const MERSA_INDEX = [MERSA_MAX_COL, MERSA_NOTE_COL, MERSA_MAX_ROW, MERSA_PROJECT_COL, MERSA_FNAME_COL, MERSA_LNAME_COL, MERSA_COMPANY_COL, MERSA_PHONE_COL, MERSA_CONDITION_COL, MERSA_COMPLETE_DATE_COL, MERSA_BILLING_COL, MERSA_NUMBER_OF_ORDERS, MERSA_CONDITION_COMPLETED];
  var mersa_project_source = sh_order.getRange(1,sh_index[3],sh_index[2],1).getValues();
  var mersa_firstname_source = sh_order.getRange(1,sh_index[4],sh_index[2],1).getValues();
  var mersa_lastname_source = sh_order.getRange(1,sh_index[5],sh_index[2],1).getValues();
  var mersa_companyname_source = sh_order.getRange(1,sh_index[6],sh_index[2],1).getValues();
  var mersa_start_date_source = sh_order.getRange(1,sh_index[7],sh_index[2],1).getValues();
  var mersa_condition_source = sh_order.getRange(1,sh_index[8],sh_index[2],1).getValues();
  var mersa_billing_source = sh_order.getRange(1,sh_index[10],sh_index[2],1).getValues();
  var mersa_complete_date_source = sh_order.getRange(1,sh_index[9],sh_index[2],1).getValues();

  var output = makeArray(7, sh_index[2], '');

  let number_of_chosen_projects = 0;

  sh_main.getRange(number_of_chosen_projects+1,1,1,1).setValue('کد پروژه');
  sh_main.getRange(number_of_chosen_projects+1,2,1,1).setValue('نام مشتری');
  sh_main.getRange(number_of_chosen_projects+1,3,1,1).setValue('نام شرکت');
  sh_main.getRange(number_of_chosen_projects+1,4,1,1).setValue('مبلغ پروژه');
  sh_main.getRange(number_of_chosen_projects+1,5,1,1).setValue('تاریخ ثبت');
  sh_main.getRange(number_of_chosen_projects+1,6,1,1).setValue('تاریخ تکمیل');
  sh_main.getRange(number_of_chosen_projects+1,7,1,1).setValue('مدیر پروژه');
  sh_main.getRange(number_of_chosen_projects+1,8,1,1).setValue('مشتری جدید');
  for (let counter = 1; counter <= sh_index[11]; counter = counter + 1) {

    if(mersa_condition_source[counter][0] === sh_index[12]) {
      const edited_date = Utilities.formatDate(mersa_complete_date_source[counter][0], 'GMT-0700', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
      const edited_date2 = Utilities.formatDate(mersa_start_date_source[counter][0], 'GMT-0700', 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');

      output[number_of_chosen_projects][0] =  mersa_project_source[counter][0];
      output[number_of_chosen_projects][1] =  mersa_firstname_source[counter][0];
      output[number_of_chosen_projects][2] =  mersa_lastname_source[counter][0];
      output[number_of_chosen_projects][3] =  mersa_companyname_source[counter][0];
      output[number_of_chosen_projects][4] =  mersa_billing_source[counter][0];
      output[number_of_chosen_projects][5] =  edited_date2;
      output[number_of_chosen_projects][6] =  edited_date;

      sh_main.getRange(number_of_chosen_projects+2,1,1,1).setValue(output[number_of_chosen_projects][0]);
      let str_name = output[number_of_chosen_projects][1] + ' ' + output[number_of_chosen_projects][2];
      sh_main.getRange(number_of_chosen_projects+2,2,1,1).setValue(str_name);
      sh_main.getRange(number_of_chosen_projects+2,3,1,1).setValue(output[number_of_chosen_projects][3]);
      sh_main.getRange(number_of_chosen_projects+2,4,1,1).setValue(output[number_of_chosen_projects][4]);
      sh_main.getRange(number_of_chosen_projects+2,5,1,1).setValue(output[number_of_chosen_projects][5]);
      sh_main.getRange(number_of_chosen_projects+2,6,1,1).setValue(output[number_of_chosen_projects][6]);

      number_of_chosen_projects = number_of_chosen_projects + 1;

    }

  }

  var range = sh_main.getRange(2,1,number_of_chosen_projects+1,6);
  range.sort(6);


}


