require(dplyr)
require(lubridate)

# prerp time
# end_time <- ymd_hms(Sys.time())
# end_time <- ymd_hms('2015-11-11 00:00:00')
end_time <- floor_date(today(), unit = 'day')
rpt_dur <- ddays(1)
start_time <- end_time - rpt_dur
pg_end_time <- as.character(end_time)
pg_start_time <- as.character(start_time)

# pulling data from db
my_db <- src_postgres(dbname = 'cxtd', host = 'localhost', port = 5432, user = 'cxtd', password = 'xintiandi')

# frozen_user <- tbl(my_db, 'member') %>% 
#     filter(status != 0) %>%
#     select(member_id) %>% collect()


# getting coupon_discount rate
coupon <- tbl(my_db, 'coupon') %>% 
    filter(coupon_redeem_datetime >= pg_start_time & 
               coupon_redeem_datetime < pg_end_time) %>% 
    select(coupon_campaign_id, redeem_sales_id)
    
coupon_campaign <- tbl(my_db, 'coupon_campaign') %>% 
                        select(coupon_campaign_id, 
                               coupon_discount_rate)

coupon_discount <- coupon %>% 
    inner_join(coupon_campaign) %>% 
    #mutate(coupon_off_rate = coupon_discount_rate / 100) %>% 
    collect() #%>%
#     group_by(redeem_sales_id) %>% 
#     summarise(coupon_off_rate = 1 - prod(1-coupon_discount_rate/100))
    # select(redeem_sales_id, coupon_off_rate) 

# create fake coupon_discount table
if (nrow(coupon_discount) == 0){
    coupon_discount <- data.frame(redeem_sales_id = integer(),
                                  coupon_off_rate = numeric()
    )
}else{
    coupon_discount <- coupon_discount %>%
        group_by(redeem_sales_id) %>% 
        summarise(coupon_off_rate = 1 - prod(1-coupon_discount_rate/100))
}

# getting other tables
member_card_df <- tbl(my_db, 'member') %>% select(member_id, member_card_no)
sales <- tbl(my_db, 'sales') %>% 
    filter(transaction_datetime >= pg_start_time & 
               transaction_datetime < pg_end_time &
               invoice_original_amount > 0) # &
               # !member_id %in% frozen_user$member_id)  


# calculate point issue and point redeem
sales_point_issue <- tbl(my_db, 'sales_point_issue') %>% 
    filter(created_datetime >= pg_start_time & 
               created_datetime < pg_end_time) %>% 
    transmute(sales_id, 
              point_issue = point)

sales_point_redemption <- tbl(my_db, 'sales_point_redemption') %>% 
    filter(created_datetime >= pg_start_time & 
               created_datetime < pg_end_time) %>% 
    transmute(sales_id, 
              point_redeemed = point,
              point_transaction_amount)
# commented out part
# member_point_transaction <- tbl(my_db, 'member_point_transaction') %>% 
#     select(member_point_transaction_id,increment_point, 
#            decrement_point, transaction_flow_type)
# 
# # prep lean df
# point_issue_redeem <- dplyr::union(sales_point_issue, sales_point_redemption) %>% 
#                             inner_join(member_point_transaction, by="member_point_transaction_id") 
                            
if (nrow(sales) > 0){
    report <- sales %>% transmute(member_id, sales_id, shop_id, transaction_datetime,
                                  original_amount = invoice_original_amount,
                                  actual_final_amount = sales_settlement_amount,
                                  total_discount_amount = invoice_chargeable_amount) %>%
        left_join(sales_point_issue, by = "sales_id") %>%
        left_join(sales_point_redemption, by = "sales_id") %>%
        inner_join(member_card_df, by = 'member_id') %>% 
        collect() %>%
        left_join(coupon_discount, by = c("sales_id" = 'redeem_sales_id')) %>%
        mutate(#coupon_discount_amount = original_amount * coupon_off_rate,
            point_cashredeem_amount = point_transaction_amount) %>%
        select(#-redeem_sales_id, 
            -point_transaction_amount) %>% 
        
        transmute(member_id, sales_id, shop_id, transaction_datetime,
                  original_amount = ifelse(is.na(original_amount), 0, original_amount), 
                  actual_final_amount = ifelse(is.na(actual_final_amount), 0, actual_final_amount), 
                  total_discount_amount = ifelse(is.na(total_discount_amount), 0, total_discount_amount), 
                  point_issue = ifelse(is.na(point_issue), 0, point_issue), 
                  point_redeemed = ifelse(is.na(point_redeemed), 0, point_redeemed), 
                  coupon_off_rate = ifelse(is.na(coupon_off_rate), 0, coupon_off_rate), 
                  member_card_no, 
                  coupon_discount_amount = original_amount * coupon_off_rate,
                  #coupon_discount_amount = ifelse(is.na(coupon_discount_amount), 0, coupon_discount_amount), 
                  point_cashredeem_amount = ifelse(is.na(point_cashredeem_amount), 0, point_cashredeem_amount) 
        )
    
    escape.POSIXt <- dplyr:::escape.Date
    db_insert_into( con = my_db$con, 
                    table = "sales_report", 
                    values = report)
    print(paste(now(), 'Success!', sep = " "))
    
}else{
    print(paste(now(), 'no sales today!', sep = " "))
}



    

            

# View(report)

