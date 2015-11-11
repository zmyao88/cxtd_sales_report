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
report_date <- as.character(today())
report_month <- as.character(month(start_time, label = T))
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
contact_person <- "*"
mailling_address <- "*"
company_name <- "*"



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

####3 write to xlsx
outwb <- createWorkbook()
# Define some cell styles within that workbook
csSheetTitle <- CellStyle(outwb) + Font(outwb, heightInPoints=12,
                                        isBold=TRUE)
csSheetSubTitle <- CellStyle(outwb) + Font(outwb,
                                           heightInPoints=12, isItalic=TRUE, isBold=FALSE)
# column and row name style
csTableRowNames <- CellStyle(outwb) + 
    Font(outwb, isBold=FALSE, heightInPoints=10) + 
                         Border(color="black",
                                position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
                                pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))

csTableColNames <- CellStyle(outwb) + 
    Font(outwb, isBold=TRUE, heightInPoints = 10) +
    Alignment(wrapText=TRUE,
              h="ALIGN_CENTER") + 
    Border(color="black",
           position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
           pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) +
    Fill(backgroundColor="blue")

# define column stytle
csDeclColumn <- CellStyle(outwb, dataFormat=DataFormat("0.0")) + 
    Border(color="black",
           position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
           pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))

csIntColumn <- CellStyle(outwb, dataFormat=DataFormat("0")) +
    Border(color="black",
           position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
           pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) 

csDateColumn <- CellStyle(outwb, dataFormat=DataFormat("YYYY-MM-DD")) + 
    Border(color="black",
           position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
           pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))

# column index and correspoinding format
date_col = list('1' = csDateColumn)
int_col = list('2' = csIntColumn,
               '7' = csIntColumn)
dec_col = list('3' = csDeclColumn,
               '4' = csDeclColumn,
               '5' = csDeclColumn,
               '6' = csDeclColumn)

# create a sheet
sheet <- createSheet(outwb, sheetName = "Sales Report")
# Add the table calculated above to the new sheet
addDataFrame(monthly_sales, sheet, startRow=8,startColumn=1, 
             colStyle= c(date_col, int_col, dec_col),
             colnamesStyle = csTableColNames, 
             rownamesStyle = csTableRowNames)
# set column width

setColumnWidth(sheet, colIndex=3, colWidth=15)
setColumnWidth(sheet, colIndex=c(1, 2, 4:5, 7:8), colWidth=10)
setColumnWidth(sheet, colIndex=6, colWidth=12)
# setColumnWidth(sheet, colIndex=c(2:15), colWidth=11)
# setColumnWidth(sheet, colIndex=16, colWidth=13)
# setColumnWidth(sheet, colIndex=17, colWidth=6)
# setColumnWidth(sheet, colIndex=1, colWidth= 0.8*max(length
#                                                     (rownames(x.RiskStats))))

# set Titles
rows <- createRow(sheet,rowIndex=1)
sheetTitle <- createCell(rows, colIndex=1)
setCellValue(sheetTitle[[1,1]], report_title)
setCellStyle(sheetTitle[[1,1]], csSheetTitle)
## set var names col
rows <- createRow(sheet,rowIndex=c(3:5))
cells <- createCell(rows,colIndex=c(1:6))
values <- c("公司名称：", "商铺名称：", "联系人：",
            company_name, shop_name, contact_person,
            "", "商铺位置：", "邮件地址：",
            "", shop_address, mailling_address,
            "账单月份：",  "账单编号：","制单日期：",
            report_month,  report_id, report_date
            )
mapply(setCellValue, cells, values)

# set some random shit
rows <- createRow(sheet,rowIndex=7)
sheetTitle <- createCell(rows, colIndex=8)
setCellStyle(sheetTitle[[1,1]], CellStyle(outwb) + 
                 Font(outwb, isBold=FALSE, heightInPoints=8) +
                 Alignment(h="ALIGN_RIGHT"))
setCellValue(sheetTitle[[1,1]], "（币种：人民币／RMB）")

# foot note
row_position <- 8 + nrow(monthly_sales) + 2

note_value <- c("备注：", 
                "1. 中国新天地每月10日前向合作商户提供上月的积分对账单，合作商户请于每月15日前进行确认并通知中国新天地，逾期则视为对中国新天地所提供的对账单无异议。",
                "2. 如合作商户对积分对账单有任何问题，请联系中国新天地 企业传讯及推广部 客户关系管理组，邮件: itiandi@xintiandi.com")
rows <- createRow(sheet,rowIndex=c(row_position:(row_position+2)))
cells <- createCell(rows,colIndex=1)
mapply(setCellValue, cells, note_value)


saveWorkbook(outwb, "new_test_work_book.xlsx")
# write.xlsx(test, file="myworkbook.xlsx",
#            sheetName="USA-ARRESTS", append=FALSE)
