function churnData() {
  
 var conn = Jdbc.getConnection(url, username, password);
 var stmt = conn.createStatement();
 


 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Dashboard');
  var cell = sheet.getRange('B4').getValues();

 var sheet = spreadsheet.getSheetByName('Membership Churn Data');
sheet.clearContents();



// start of volunteering function
function stats(results, title)
{

 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Membership Churn Data');
 
sheet.appendRow([title]);


  //console.log(results +"sd");
 var metaData=results.getMetaData();
  var numCols = metaData.getColumnCount();
 var arr=[];
  for (var col = 0; col < numCols; col++) {
   arr.push(metaData.getColumnLabel(col + 1));
 }
  sheet.appendRow(arr);
 while (results.next()) {
 arr=[];
 for (var col = 0; col < numCols; col++) {
   arr.push(results.getString(col + 1));
 }
 sheet.appendRow(arr);
 }


sheet.autoResizeColumns(1, numCols+1);



} // end of function
var results = stmt.executeQuery("SELECT   joined.Month,   joined.`Members Joined this Month`,   IFNULL(lost.`Members Lost this Month`, 0) AS `Members Lost this Month` FROM   (SELECT     DATE_FORMAT(pd.order_created, '%Y-%m-01') AS 'Month',     COUNT(pd.order_id) AS 'Members Joined this Month'   FROM     wp_member_db db     JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id   WHERE     product_id = 2118     AND db.cc_member = 'yes'     AND status IN ('wc-processing', 'wc-completed')     AND DATE(pd.order_created) NOT IN ('2022-01-02', '2023-01-02', '2024-01-02')   GROUP BY     DATE_FORMAT(pd.order_created, '%Y-%m-01')   ORDER BY     DATE_FORMAT(pd.order_created, '%Y-%m-01') ASC) AS joined LEFT JOIN   (SELECT     COALESCE(       DATE_FORMAT(FROM_UNIXTIME(cc_membership_cancellation_intent_date), '%Y-%m-01'),       DATE_FORMAT(membership_leaving_date, '%Y-%m-01'),       IF(         db.cc_member = 'expired' AND DATE(pd.order_created) < DATE('2022-01-01'),         '2021-12-01',         NULL       )     ) AS 'Month',     COUNT(pd.order_id) AS 'Members Lost this Month'   FROM     wp_member_db db     JOIN wp_order_product_customer_lookup pd ON pd.user_id = db.id   WHERE     product_id = 2118     AND (db.cc_member = 'expired' OR cc_membership_cancellation_intent_date <> '')     AND status IN ('wc-processing', 'wc-completed')     AND DATE(pd.order_created) <> '2022-01-02'   GROUP BY     `Month`) AS lost ON   joined.Month = lost.Month ORDER BY   joined.Month ASC")

stats(results,"Member orders");




// select distinct  `first_name` "Forenames",`last_name` "Surname", `admin-dob` "Dob", `shipping_address_1` "Address 1",`shipping_address_2` "Address 2",`shipping_city` "Town", `shipping_postcode` "Postcode",`billing_email` "Email",`admin-phone-number` "Home Tel.",pd.user_id "Club Ref", `admin-bmc-membership-number`  "BMC Ref",`admin-membership-type` AS "Membership Type", (SELECT IFNULL( (select `dbi`.`committee_current` from `wp_member_db` dbi where dbi.id = `db`.`id` LIMIT 1) ,"")) AS `Membership Title`, (SELECT (CASE WHEN `dbi`.`committee_current`="secretary" THEN "Main Club Contact" WHEN `dbi`.`committee_current`="chair" THEN "Communications Contact" ELSE "Club Member" END) FROM `wp_member_db` dbi where dbi.id = `db`.`id` LIMIT 1) AS `Contact Type`, pd.order_id "Order Ref" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1 order by `first_name` ASC;





} 

