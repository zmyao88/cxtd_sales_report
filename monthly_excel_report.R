source('~/src/cxtd_sales_report/monthly_excel_util_v2.R')

require(xlsx)
require(dplyr)
require(lubridate)


my_db <- src_postgres(dbname = 'cxtd', host = 'localhost', port = 5432, user = 'cxtd', password = 'CxTd1234!')

# end_time <- ymd_hms(Sys.time())
# end_time <- ymd_hms('2017-04-01 00:00:00')
# prep time 
end_time <- floor_date(today(), unit = 'month')
rpt_dur <- ddays(7)
start_time <- floor_date(end_time - rpt_dur, unit = 'month')

pg_end_time <- as.character(end_time)
pg_start_time <- as.character(start_time)

# load shop data
shop <- tbl(my_db, 'shop') %>% filter(status == 0)
partner <- tbl(my_db, 'partner') 
contract_shop_mapping <- tbl(my_db, 'contract_shop_mapping') %>% select(contract_id, shop_id) 
contract <- tbl(my_db, 'contract') %>% 
                filter(contract_end_datetime >= pg_start_time) %>% 
                filter(contract_type %in% c(0,1,2,3,4) & status == 0) %>% 
                inner_join(contract_shop_mapping)
# frozen members
suspended_members <- tbl(my_db, 'member') %>% filter(status == 99) %>% select(member_id)

shop_IDS <- contract %>% select(shop_id) %>% collect()
# shop_IDS <- shop_IDS %>% filter(row_number()<5)
if (nrow(shop_IDS) > 0){
    for (current_shop_id in unique(shop_IDS$shop_id)){
        print_xls_output(shop, partner, contract, suspended_members, current_shop_id, my_db, pg_start_time, pg_end_time)
    }
    print(paste(now(), 'Monthly sales reoprts done!', sep = " "))
    
}else{
    print(paste(now(), 'No sales reports for this month', sep = " "))
}


