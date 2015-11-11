require(xlsx)
file <- system.file("tests", "test_import.xlsx", package = "xlsx")
res <- read.xlsx(file, 1)  # read first sheet
head(res[, 1:6])

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

# prep report infos
report_date <- today()
report_month <- month(start_time, label = T)
report_title <- "i天地合作商户对账单"
    

current_shop_id = 20128
current_shop <- shop %>% filter(shop_id == current_shop_id) %>% collect()

#shop_name 
shop_name <- current_shop$name_sc
shop_address <- current_shop$address_sc
report_id <- paste(current_shop$shop_code, 
                   year(start_time), 
                   month(start_time, label = T), 
                   sep = "")
contact_person <- ""
mailling_address <- ""



# real sales_report data 
monthly_sales <- tbl(my_db, 'sales_report') %>% 
    filter(transaction_datetime >= pg_start_time &
               transaction_datetime < pg_end_time &
               shop_id == current_shop_id) %>%
    select(transaction_datetime, member_card_no,
           original_amount, point_cashredeem_amount, 
           coupon_discount_amount, actual_final_amount,
           point_issue) %>% 
    collect()

my_colnames <- c("消费日期", "会员卡号", "消费金额", "积分抵扣金额",
                             "优惠券抵扣金额", "实际支付金额", "实际累计积分")
colnames(monthly_sales) <- my_colnames


# if (nrow(monthly_sales) == 0){
#     monthly_sales <- data.frame(member_id = numeric(),
#                                 transaction_datetime = character()
#                                 )
# }

test <- monthly_sales 

write.xlsx(test, file="myworkbook.xlsx",
           sheetName="USA-ARRESTS", append=FALSE)

mon