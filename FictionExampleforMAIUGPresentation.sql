/*Creates a shelf list of all items with a fiction call number in either the bib or the item. 
Handles missing publication dates and excludes paperback items records that may be 
attached to bibs that fall within that range.
Created by Eric McCarthy 10/2017 for the Greenwich Library to be used with the Sierra ILS/ */
SELECT
call.index_entry AS "Bib Call #",
itemcall.call_number AS "Item Call #",
title.best_title AS "Title",
title.best_author AS "Author",
--How to handle missing data in the publication field
CASE
 WHEN title.publish_year IS NULL THEN 0
 ELSE title.publish_year END AS "Pub. Year",
item.record_creation_date_gmt AS "Item Created",
--How to handle no data in the last checked out field
CASE
 WHEN to_char (item.last_checkin_gmt, 'yyyy-mm-dd hh:mi AM') IS NULL THEN 'No data'
 ELSE to_char (item.last_checkin_gmt, 'yyyy-mm-dd hh:mi AM')
 END AS "Last Checkin",
item.checkout_total AS "Total Checkouts",
item.renewal_total AS "Total Renewals",
item.year_to_date_checkout_total AS "YTD",
item.last_year_to_date_checkout_total AS "LYR",
--Checks to see if an item is checked out
CASE
 WHEN checkout.loanrule_code_num > 1 THEN 'YES'
 ELSE 'NO' END as "Checked out?",
item.barcode AS "Barcode",
item.item_status_code AS "Status"

FROM
sierra_view.phrase_entry AS call
JOIN sierra_view.bib_record_property AS title ON title.bib_record_id = call.record_id
JOIN sierra_view.bib_record_item_record_link AS link ON call.record_id = link.bib_record_id
JOIN sierra_view.item_view AS item ON item.id = link.item_record_id 
LEFT JOIN sierra_view.item_record_property AS itemcall ON itemcall.item_record_id = link.item_record_id
FULL OUTER JOIN sierra_view.checkout AS checkout ON checkout.item_record_id = link.item_record_id

WHERE

call.varfield_type_code = 'c' AND
(call.index_entry LIKE 'fiction%' OR
itemcall.call_number LIKE 'FICTION%') AND
item.location_code = 'gmad2' 



ORDER BY

call.index_entry, itemcall.call_number;