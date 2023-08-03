function mersa_add_data() {
  var mersa_google_sheets = SpreadsheetApp.getActiveSpreadsheet();
  var mersa_woocommerce = mersa_google_sheets.getSheetByName("WOO");
  var mersa_orders = mersa_google_sheets.getSheetByName("WOO_Orders");
  var mersa_customers = mersa_google_sheets.getSheetByName("WOO_Customers");

  const MERSA_MAX_COL = 2000;
  const MERSA_MAX_ROW = 1000;
  const MERSA_PHONE_COL = 5;
  const MERSA_CONDITION_COL = 6;
  const MERSA_ID_COL = 7;
  const MERSA_BILLING_COL = 8;
  const MERSA_CONDITION_COMPLETED = "completed";
  const MERSA_CONDITION_PROCESSING = "processing";
  const MERSA_CONDITION_PENDING = "pending";

  var mersa_index_array = [MERSA_PHONE_COL, MERSA_CONDITION_COL, MERSA_ID_COL, MERSA_BILLING_COL, MERSA_MAX_COL, MERSA_MAX_ROW];
  var mersa_condition_array = [MERSA_CONDITION_COMPLETED, MERSA_CONDITION_PROCESSING, MERSA_CONDITION_PENDING];

  var mersa_firstname_source = mersa_customers.getRange(1,1,MERSA_MAX_COL,1);
  var mersa_lastname_source = mersa_customers.getRange(1,2,MERSA_MAX_COL,1);
  var mersa_companyname_source = mersa_customers.getRange(1,3,MERSA_MAX_COL,1);
  var mersa_username_source = mersa_customers.getRange(1,4,MERSA_MAX_COL,1);
  var mersa_phone_source = mersa_customers.getRange(1,5,MERSA_MAX_COL,1);
  var mersa_mail_source = mersa_customers.getRange(1,6,MERSA_MAX_COL,1);
  var mersa_userinfo_source = mersa_customers.getRange(1,7,MERSA_MAX_COL,1);

  var mersa_firstname_dest = mersa_woocommerce.getRange(1,1,MERSA_MAX_COL,1);
  var mersa_lastname_dest = mersa_woocommerce.getRange(1,2,MERSA_MAX_COL,1);
  var mersa_companyname_dest = mersa_woocommerce.getRange(1,12,MERSA_MAX_COL,1);
  var mersa_username_dest = mersa_woocommerce.getRange(1,3,MERSA_MAX_COL,1);
  var mersa_phone_dest = mersa_woocommerce.getRange(1,13,MERSA_MAX_COL,1);
  var mersa_mail_dest = mersa_woocommerce.getRange(1,14,MERSA_MAX_COL,1);
  var mersa_userinfo_dest = mersa_woocommerce.getRange(1,15,MERSA_MAX_COL,1);

  var mersa_sheet_dest = mersa_woocommerce.getRange(1,1,MERSA_MAX_COL,MERSA_MAX_ROW);

  mersa_firstname_source.copyTo(mersa_firstname_dest);
  mersa_lastname_source.copyTo(mersa_lastname_dest);
  mersa_companyname_source.copyTo(mersa_companyname_dest);
  mersa_username_source.copyTo(mersa_username_dest);
  mersa_phone_source.copyTo(mersa_phone_dest);
  mersa_mail_source.copyTo(mersa_mail_dest);
  mersa_userinfo_source.copyTo(mersa_userinfo_dest);

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
  mersa_woocommerce.getRange(1,1,1,MERSA_MAX_ROW).setFontColor('White');
  mersa_woocommerce.getRange(1,1,1,MERSA_MAX_ROW).setBackground('Black');
  mersa_sheet_dest.setVerticalAlignment('Middle');
  mersa_sheet_dest.setHorizontalAlignment('Center');
  mersa_sheet_dest.setWrap(true);

  mersa_get_customer_orders(mersa_woocommerce, mersa_orders, mersa_customers, mersa_index_array, mersa_condition_array);

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
  */
  /*
    // condition_arr
      const MERSA_CONDITION_COMPLETED = "completed"; [0]
      const MERSA_CONDITION_PROCESSING = "processing"; [1]
      const MERSA_CONDITION_PENDING = "pending"; [2]
  */
  for (let counter = 2; counter <= indx_arr[5]; counter = counter + 1) {
    console.log("Customer: " + counter);
    let id = sh_customer.getRange(counter,indx_arr[2]).getValue();
    let phone = sh_customer.getRange(counter,indx_arr[0]).getValue();
    let single_customer = mersa_get_single_customer_orders(sh_order, id, phone, indx_arr, condition_arr);
    sh_main.getRange(counter,4,1,1).setValue(single_customer[0]);
    sh_main.getRange(counter,5,1,1).setValue(single_customer[1]);
    sh_main.getRange(counter,6,1,1).setValue(single_customer[2]);
    sh_main.getRange(counter,7,1,1).setValue(single_customer[3]);
    sh_main.getRange(counter,8,1,1).setValue(single_customer[4]);
    sh_main.getRange(counter,9,1,1).setValue(single_customer[5]);
    sh_main.getRange(counter,10,1,1).setValue(single_customer[6]);
    sh_main.getRange(counter,11,1,1).setValue(single_customer[7]);
  }
}

function mersa_get_single_customer_orders(sh_order, id, phone, indx_arr, condition_arr) {
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
  */
  /*
    // condition_arr
      const MERSA_CONDITION_COMPLETED = "completed"; [0]
      const MERSA_CONDITION_PROCESSING = "processing"; [1]
      const MERSA_CONDITION_PENDING = "pending"; [2]
  */
  for (let counter = 2; counter <= indx_arr[4]; counter = counter + 1) {
    if(
      (sh_order.getRange(counter,indx_arr[2]).getValue() === id) &&
      (sh_order.getRange(counter,indx_arr[0]).getValue() === phone)
    )
    {
      if(sh_order.getRange(counter,indx_arr[1]).getValue() === condition_arr[0]) {
        count_comleted = count_comleted + 1;
        cost_comleted = cost_comleted + sh_order.getRange(counter,indx_arr[3]).getValue();
      }
      else if(sh_order.getRange(counter,indx_arr[1]).getValue() === condition_arr[1]) {
        count_processing = count_processing + 1;
        cost_processing = cost_processing + sh_order.getRange(counter,indx_arr[3]).getValue();
      }
      else if(sh_order.getRange(counter,indx_arr[1]).getValue() === condition_arr[2]) {
        count_pending = count_pending + 1;
        cost_pending = cost_pending + sh_order.getRange(counter,indx_arr[3]).getValue();
      }
      else {
        count_cancelled = count_cancelled + 1;
        cost_cancelled = cost_cancelled + sh_order.getRange(counter,indx_arr[3]).getValue();
      }
    }
  }
  let retval = [count_comleted, cost_comleted, count_processing, cost_processing, count_pending, cost_pending, count_cancelled, cost_cancelled];
  return retval;
}
