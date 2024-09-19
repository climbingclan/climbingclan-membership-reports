var server = '18.168.242.164';
var port = 3306;
var dbName = 'bitnami_wordpress';
var username = 'gsheets';
var password = 'eyai4yohF4uX8eeP7phoob';
var url = 'jdbc:mysql://'+server+':'+port+'/'+dbName;

function readData() {
  
 var conn = Jdbc.getConnection(url, username, password);
 var stmt = conn.createStatement();
 


 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Dashboard');
  var cell = sheet.getRange('B4').getValues();

 var sheet = spreadsheet.getSheetByName('Membership Report');
sheet.clearContents();



// start of volunteering function
function stats(results, title)
{

 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Membership Report');
 
sheet.appendRow([title]);


  //console.log(results +"sd");
 var metaData=results.getMetaData();
  var numCols = metaData.getColumnCount();
 var arr=[];
  for (var col = 0; col < numCols; col++) {
   arr.push(metaData.getColumnLabel(col + 1));
 }
 // sheet.appendRow(arr);
 while (results.next()) {
 arr=[];
 for (var col = 0; col < numCols; col++) {
   arr.push(results.getString(col + 1));
 }
 sheet.appendRow(arr);
 }


sheet.autoResizeColumns(1, numCols+1);



} // end of function
var results = stmt.executeQuery('select count(pd.order_id) as "Current Membership" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1 order by `first_name` ASC')

stats(results,"Number of members");

var results = stmt.executeQuery('select distinct count(db.id) as "Non Members" from wp_member_db db where (db.cc_member IS NULL or db.cc_member<>"yes")')
stats(results,"Number of registered community without Membership");

var results = stmt.executeQuery('select distinct count(db.id) as "Everyone" from wp_member_db db')
stats(results,"Members +  NonMembers community total");

var results = stmt.executeQuery('select distinct count(db.id) as "Non Members Never" from wp_member_db db where (db.cc_member IS NULL or db.cc_member<>"yes") AND stats_attendance_attended_cached=0')
stats(results,"Number of Community who have never attended");

var results = stmt.executeQuery('select count(pd.order_id) as "Members Never" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached=0')
stats(results,"Members who have never attended");


var results = stmt.executeQuery('select COUNT(id) "Members" from wp_member_db where cc_member="yes" AND  `skills-trad-climbing` like "%lead%" order by CAST(stats_attendance_attended_cached AS UNSIGNED INTEGER) DESC')
stats(results,"Members who lead trad");
// select nickname "FB Name",`climbing-trad-grades` as "Grades", stats_attendance_attended_cached "number of events attended", DATE_FORMAT(DATE(order_created), '%d-%m') "Member since" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND  `skills-trad-climbing` like "%lead%" AND DATE(pd.order_created)<>"2022-01-02" order by CAST(stats_attendance_attended_cached AS UNSIGNED INTEGER) DESC


var results = stmt.executeQuery('select count(order_id) "Members" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND  `skills-belaying` like "%lead%" AND DATE(pd.order_created)<>"2022-01-02" AND  `skills-trad-climbing` not like "%lead%" AND `skills-trad-climbing` not like "%learn%"')

stats(results,"Members who have joined this year who lead belay but don't lead trad yet");
// select nickname "FB Name", stats_attendance_attended_cached "number of events attended", DATE_FORMAT(DATE(order_created), '%d-%m') "Member since" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND  `skills-belaying` like "%lead%" AND DATE(pd.order_created)<>"2022-01-02" AND  `skills-trad-climbing` not like "%lead%" AND `skills-trad-climbing` not like "%learn%" order by CAST(stats_attendance_attended_cached AS UNSIGNED INTEGER) DESC

var results = stmt.executeQuery('select count(pd.order_id) as "Once" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached=1')
stats(results,"Members who have attended once");

var results = stmt.executeQuery('select count(pd.order_id) as "Twice" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached=2')
stats(results,"Members who have attended twice");

var results = stmt.executeQuery('select count(pd.order_id) as "Less than 5" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached<=5')
stats(results,"Members who have attended less than 5");

var results = stmt.executeQuery('select count(pd.order_id) as "More than 10" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached>=10')
stats(results,"Members who have attended more than 10");

var results = stmt.executeQuery('select count(pd.order_id) as "More than 20" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached>=20')
stats(results,"Members who have attended more than 20");

var results = stmt.executeQuery('select count(pd.order_id) as "More than 30" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached>=30')
stats(results,"Members who have attended more than 30");

var results = stmt.executeQuery('select count(pd.order_id) as "More than 40" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND stats_attendance_attended_cached>=40')
stats(results,"Members who have attended more than 40");


// select distinct  `first_name` "Forenames",`last_name` "Surname", `admin-dob` "Dob", `shipping_address_1` "Address 1",`shipping_address_2` "Address 2",`shipping_city` "Town", `shipping_postcode` "Postcode",`billing_email` "Email",`admin-phone-number` "Home Tel.",pd.user_id "Club Ref", `admin-bmc-membership-number`  "BMC Ref",`admin-membership-type` AS "Membership Type", (SELECT IFNULL( (select `dbi`.`committee_current` from `wp_member_db` dbi where dbi.id = `db`.`id` LIMIT 1) ,"")) AS `Membership Title`, (SELECT (CASE WHEN `dbi`.`committee_current`="secretary" THEN "Main Club Contact" WHEN `dbi`.`committee_current`="chair" THEN "Communications Contact" ELSE "Club Member" END) FROM `wp_member_db` dbi where dbi.id = `db`.`id` LIMIT 1) AS `Contact Type`, pd.order_id "Order Ref" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1 order by `first_name` ASC;




membershipData();
} 

