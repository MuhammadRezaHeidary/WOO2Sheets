function mersa_add_data() {

  var mersa_google_sheets = SpreadsheetApp.getActiveSpreadsheet();
  var mersa_woocommerce = mersa_google_sheets.getSheetByName("WOO");
  var mersa_orders = mersa_google_sheets.getSheetByName("WOO_Orders");
  var mersa_customers = mersa_google_sheets.getSheetByName("WOO_Customers");

  const MERSA_MAX_COL = 1000;
  const MERSA_NOTE_COL = 6;
  const MERSA_MAX_ROW = 2000;
  const MERSA_PROJECT_COL = 1;
  const MERSA_USERNAME_COL = 4;
  const MERSA_PHONE_COL = 5;
  const MERSA_CONDITION_COL = 6;
  const MERSA_ID_COL = 7;
  const MERSA_BILLING_COL = 8;
  const MERSA_NUMBER_OF_ORDERS = mersa_get_number_of_filled_rows(mersa_orders, MERSA_PROJECT_COL, MERSA_MAX_ROW);
  const MERSA_NUMBER_OF_CUSTOMERS = mersa_get_number_of_filled_rows(mersa_customers, MERSA_USERNAME_COL, MERSA_MAX_ROW);
  const MERSA_CONDITION_COMPLETED = "completed";
  const MERSA_CONDITION_PROCESSING = "processing";
  const MERSA_CONDITION_PENDING = "pending";

  var mersa_index_array = [MERSA_PHONE_COL, MERSA_CONDITION_COL, MERSA_ID_COL, MERSA_BILLING_COL, MERSA_MAX_ROW, MERSA_MAX_COL, MERSA_NUMBER_OF_ORDERS, MERSA_NUMBER_OF_CUSTOMERS];
  var mersa_condition_array = [MERSA_CONDITION_COMPLETED, MERSA_CONDITION_PROCESSING, MERSA_CONDITION_PENDING];

  var mersa_firstname_source = mersa_customers.getRange(1,1,MERSA_MAX_ROW,1);
  var mersa_lastname_source = mersa_customers.getRange(1,2,MERSA_MAX_ROW,1);
  var mersa_companyname_source = mersa_customers.getRange(1,3,MERSA_MAX_ROW,1);
  var mersa_username_source = mersa_customers.getRange(1,4,MERSA_MAX_ROW,1);
  var mersa_phone_source = mersa_customers.getRange(1,5,MERSA_MAX_ROW,1);
  var mersa_mail_source = mersa_customers.getRange(1,6,MERSA_MAX_ROW,1);
  var mersa_userid_source = mersa_customers.getRange(1,7,MERSA_MAX_ROW,1);
  var mersa_userinfo_source = mersa_customers.getRange(1,8,MERSA_MAX_ROW,1);

  var mersa_firstname_dest = mersa_woocommerce.getRange(1,1,MERSA_MAX_ROW,1);
  var mersa_lastname_dest = mersa_woocommerce.getRange(1,2,MERSA_MAX_ROW,1);
  var mersa_companyname_dest = mersa_woocommerce.getRange(1,12,MERSA_MAX_ROW,1);
  var mersa_username_dest = mersa_woocommerce.getRange(1,3,MERSA_MAX_ROW,1);
  var mersa_phone_dest = mersa_woocommerce.getRange(1,13,MERSA_MAX_ROW,1);
  var mersa_mail_dest = mersa_woocommerce.getRange(1,14,MERSA_MAX_ROW,1);
  var mersa_userid_dest = mersa_woocommerce.getRange(1,15,MERSA_MAX_ROW,1);
  var mersa_userinfo_dest = mersa_woocommerce.getRange(1,19,MERSA_MAX_ROW,1);

  var mersa_sheet_dest = mersa_woocommerce.getRange(1,1,MERSA_MAX_ROW,MERSA_MAX_COL);

  var mersa_notes = mersa_woocommerce.getRange(1,13,MERSA_NUMBER_OF_CUSTOMERS,MERSA_NOTE_COL).getValues();

  mersa_firstname_source.copyTo(mersa_firstname_dest);
  mersa_lastname_source.copyTo(mersa_lastname_dest);
  mersa_companyname_source.copyTo(mersa_companyname_dest);
  mersa_username_source.copyTo(mersa_username_dest);
  mersa_phone_source.copyTo(mersa_phone_dest);
  mersa_mail_source.copyTo(mersa_mail_dest);
  mersa_userid_source.copyTo(mersa_userid_dest);
  mersa_userinfo_source.copyTo(mersa_userinfo_dest);

  var mersa_notes_update = mersa_woocommerce.getRange(1,13,MERSA_NUMBER_OF_CUSTOMERS,MERSA_NOTE_COL).getValues();

  mersa_woocommerce.getRange(1,4,1,1).setValue('تعداد پروژه های تکمیل شده');
  mersa_woocommerce.getRange(1,5,1,1).setValue('هزینه پروژه های تکمیل شده');
  mersa_woocommerce.getRange(1,6,1,1).setValue('تعداد پروژه های در حال انجام');
  mersa_woocommerce.getRange(1,7,1,1).setValue('هزینه پروژه های در حال انجام');
  mersa_woocommerce.getRange(1,8,1,1).setValue('تعداد پروژه های در انتظار پرداخت');
  mersa_woocommerce.getRange(1,9,1,1).setValue('هزینه پروژه های در انتظار پرداخت');
  mersa_woocommerce.getRange(1,10,1,1).setValue('تعداد پروژه های لغو شده');
  mersa_woocommerce.getRange(1,11,1,1).setValue('هزینه پروژه های لغو شده');

  mersa_sheet_dest.setFontFamily('Calibri');
  mersa_sheet_dest.setFontSize(16);
  mersa_woocommerce.getRange(1,1,1,MERSA_MAX_COL).setFontColor('White');
  mersa_woocommerce.getRange(1,1,1,MERSA_MAX_COL).setBackground('Black');
  mersa_sheet_dest.setVerticalAlignment('Middle');
  mersa_sheet_dest.setHorizontalAlignment('Center');
  mersa_sheet_dest.setWrap(true);

  mersa_get_customer_orders(mersa_woocommerce, mersa_orders, mersa_customers, mersa_index_array, mersa_condition_array);
  var notes = mersa_set_notes(mersa_notes, mersa_notes_update, MERSA_NUMBER_OF_CUSTOMERS);
  mersa_woocommerce.getRange(1,16,MERSA_NUMBER_OF_CUSTOMERS,3).setValues(notes);

  mersa_woocommerce.getRange(1,16,1,1).setValue('یادداشت 1');
  mersa_woocommerce.getRange(1,17,1,1).setValue('یادداشت 2');
  mersa_woocommerce.getRange(1,18,1,1).setValue('یادداشت 3');
  mersa_woocommerce.getRange(1,19,1,1).setValue('یادداشت ووکامرس (اصلاح نشود)');

  setFilter(mersa_google_sheets, "WOO");

}


function setFilter(ss, sheet_name) {
  var filterSettings = {};

  filterSettings.range = {
    sheetId: ss.getSheetByName(sheet_name).getSheetId() // provide your sheetname to which you want to apply filter.
  };

  filterSettings.criteria = {};
  var columnIndex = 9; // column that defines criteria [A = 0]
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

/*
  buf[0]: phone number
  buf[1]: email
  buf[2]: id
  buf[3:5]: note
 */
function mersa_set_notes(note_buf, note_update_buf, max_row) {
  let retval = makeArray(3, max_row, '');
  for (let counter = 1; counter <= max_row - 1; counter = counter + 1) {
    for(let counter2 = 1; counter2 <= max_row - 1; counter2 = counter2 + 1) {
      if(
        (note_update_buf[counter][0] === note_buf[counter2][0]) &&
        (note_update_buf[counter][2] === note_buf[counter2][2])
      ) {
        retval[counter][0] = note_buf[counter2][3];
        retval[counter][1] = note_buf[counter2][4];
        retval[counter][2] = note_buf[counter2][5];
        break;
      }
    }
  }
  return retval;
}

function mersa_get_number_of_filled_rows(sh, col, max_row) {
  let mersa_row_data = sh.getRange(1, col, max_row, 1).getValues();
  for (let counter = 1; counter <= max_row - 1; counter = counter + 1) {
    if(mersa_row_data[counter][0] === '')
      return counter - 1;
  }
  return max_row;
}

function mersa_get_customer_orders(sh_main ,sh_order, sh_customer, indx_arr, condition_arr) {
  /*
    // indx_arr
      const MERSA_PHONE_COL = 5; [0]
      const MERSA_CONDITION_COL = 6; [1]
      const MERSA_ID_COL = 7; [2]
      const MERSA_BILLING_COL = 8; [3]
      const MERSA_MAX_COL = 2000; [4]
      const MERSA_MAX_ROW = 1000; [5]
      const MERSA_NUMBER_OF_ORDERS = mersa_get_number_of_filled_rows(mersa_orders, MERSA_PROJECT_COL, MERSA_MAX_ROW); [6]
      const MERSA_NUMBER_OF_CUSTOMERS = mersa_get_number_of_filled_rows(mersa_customers, MERSA_USERNAME_COL, MERSA_MAX_ROW); [7]
  */
  /*
    // condition_arr
      const MERSA_CONDITION_COMPLETED = "completed"; [0]
      const MERSA_CONDITION_PROCESSING = "processing"; [1]
      const MERSA_CONDITION_PENDING = "pending"; [2]
  */
  var mersa_orders_data = sh_order.getRange(1, indx_arr[0], indx_arr[6]+1, 4).getValues();
  var mersa_customers_id = sh_customer.getRange(1, indx_arr[2], indx_arr[7]+1, 1).getValues();
  var mersa_customers_phone = sh_customer.getRange(1, indx_arr[0], indx_arr[7]+1, 1).getValues();


  for (let counter = 1; counter <= indx_arr[7]; counter = counter + 1) {
    let id = mersa_customers_id[counter][0];
    let phone = mersa_customers_phone[counter][0];
    console.log("Customer: " + counter);
    console.log("id: " + id);
    console.log("phone: " + phone);

    let single_customer = mersa_get_single_customer_orders(mersa_orders_data, id, phone, indx_arr, condition_arr);
    sh_main.getRange(counter+1,4,1,1).setValue(single_customer[0]);
    sh_main.getRange(counter+1,5,1,1).setValue(single_customer[1]);
    sh_main.getRange(counter+1,6,1,1).setValue(single_customer[2]);
    sh_main.getRange(counter+1,7,1,1).setValue(single_customer[3]);
    sh_main.getRange(counter+1,8,1,1).setValue(single_customer[4]);
    sh_main.getRange(counter+1,9,1,1).setValue(single_customer[5]);
    sh_main.getRange(counter+1,10,1,1).setValue(single_customer[6]);
    sh_main.getRange(counter+1,11,1,1).setValue(single_customer[7]);
  }
}

function mersa_get_single_customer_orders(orders_data, id, phone, indx_arr, condition_arr) {
  let count_comleted = 0;
  let cost_comleted = 0;
  let count_processing = 0;
  let cost_processing = 0;
  let count_pending = 0;
  let cost_pending = 0;
  let count_cancelled = 0;
  let cost_cancelled = 0;
  /*
    // indx_arr
      const MERSA_PHONE_COL = 5; [0]
      const MERSA_CONDITION_COL = 6; [1]
      const MERSA_ID_COL = 7; [2]
      const MERSA_BILLING_COL = 8; [3]
      const MERSA_MAX_COL = 2000; [4]
      const MERSA_MAX_ROW = 1000; [5]
      const MERSA_NUMBER_OF_ORDERS = mersa_get_number_of_filled_rows(mersa_orders, MERSA_PROJECT_COL, MERSA_MAX_ROW); [6]
      const MERSA_NUMBER_OF_CUSTOMERS = mersa_get_number_of_filled_rows(mersa_customers, MERSA_USERNAME_COL, MERSA_MAX_ROW); [7]
  */
  /*
    // condition_arr
      const MERSA_CONDITION_COMPLETED = "completed"; [0]
      const MERSA_CONDITION_PROCESSING = "processing"; [1]
      const MERSA_CONDITION_PENDING = "pending"; [2]
  */
  for (let counter = 1; counter <= indx_arr[6]; counter = counter + 1) {
    if(
      (orders_data[counter][2] === id) && 
      (orders_data[counter][0] === phone)
    )
    {
      if(orders_data[counter][1] === condition_arr[0]) {
        count_comleted = count_comleted + 1;
        cost_comleted = cost_comleted + orders_data[counter][3];
      }
      else if(orders_data[counter][1] === condition_arr[1]) {
        count_processing = count_processing + 1;
        cost_processing = cost_processing + orders_data[counter][3];
      }
      else if(orders_data[counter][1] === condition_arr[2]) {
        count_pending = count_pending + 1;
        cost_pending = cost_pending + orders_data[counter][3];
      }
      else {
        count_cancelled = count_cancelled + 1;
        cost_cancelled = cost_cancelled + orders_data[counter][3];
      }
    }
  }
  let retval = [count_comleted, cost_comleted, count_processing, cost_processing, count_pending, cost_pending, count_cancelled, cost_cancelled];
  return retval;
}
