source('./monthly_excel_util.R')

require(xlsx)
require(dplyr)
require(lubridate)


my_db <- src_postgres(dbname = 'cxtd', host = 'localhost', port = 5432, user = 'cxtd', password = 'xintiandi')

# end_time <- ymd_hms(Sys.time())
# end_time <- ymd_hms('2015-10-01 00:00:00')
# prep time 
end_time <- floor_date(today(), unit = 'month')
rpt_dur <- ddays(7)
start_time <- floor_date(end_time - rpt_dur, unit = 'month')

pg_end_time <- as.character(end_time)
pg_start_time <- as.character(start_time)

# load shop data
shop <- tbl(my_db, 'shop') 
contract <- tbl(my_db, 'contract')

shop_IDS <- shop %>% filter(status == 0) %>% select(shop_id) %>% collect()

for (current_shop_id in unique(shop_IDS$shop_id)){
    print_xls_output(shop, current_shop_id, my_db, pg_start_time, pg_end_time)
}