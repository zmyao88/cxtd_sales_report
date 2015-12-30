require(xlsx)
require(dplyr)
require(lubridate)

substrRight <- function(x, n){
    substr(x, nchar(x)-n+1, nchar(x))
}

print_xls_output <- function(shop_tbl, partner_tbl, contract_tbl, suspended_members_tbl, current_shop_id, db_con, pg_start_time, pg_end_time){
    
    current_shop <- shop_tbl %>% filter(shop_id == current_shop_id) %>% collect()
    current_contract <- contract_tbl %>% filter(shop_id == current_shop_id) %>% collect()
    current_partner <- partner_tbl %>% filter(partner_id == current_contract$partner_id) %>% collect()
    
    # prep report infos
    report_date <- as.character(today())
    report_month <- as.character(month(start_time, label = T))
    report_title <- "i天地合作商户对账单"
    
    #shop_name 
    shop_name <- current_shop$name_sc
    shop_address <- current_shop$address_sc
    report_id <- paste('DZ',
                       substr(current_shop$shop_code, start = 1, stop = 6),
                       substr(year(start_time), start = 3, stop = 4),
                       ifelse(nchar(month(start_time)) == 1, paste("0", month(start_time), sep = ""), month(start_time)), 
                       substrRight(current_shop$shop_code, n = 3),
                       sep = "")
    contact_person <- current_partner$contact_person_1
    mailling_address <- current_partner$address_sc
    company_name <- current_partner$name_sc

    
    
    # real sales_report data 
    monthly_sales <- tbl(db_con, 'sales_report') %>% 
        filter(transaction_datetime >= pg_start_time &
                   transaction_datetime < pg_end_time &
                   shop_id == current_shop_id) %>% 
        anti_join(suspended_members_tbl, by = "member_id") %>% 
        select(transaction_datetime, member_card_no,
               original_amount, point_cashredeem_amount, 
               coupon_discount_amount, actual_final_amount,
               point_issue) %>% 
        collect()
    
    # calculate total
    if(nrow(monthly_sales) > 0){
        monthly_total <- monthly_sales %>% 
            summarise(original_amount_total = sum(original_amount, na.rm = T),
                      point_cashredeem_amount_total = sum(point_cashredeem_amount, na.rm = T),
                      coupon_discount_amount_total = sum(coupon_discount_amount, na.rm = T),
                      actual_final_amount_total = sum(actual_final_amount, na.rm = T),
                      point_issue_total = sum(point_issue, na.rm = T))
    }else{
        # create "fake monthly_total" if there is no monthly_sales_data
        monthly_sales <- data.frame(transaction_datetime = character(),
                                    member_card_no = character(),
                                    original_amount = character(),
                                    point_cashredeem_amount = character(),
                                    coupon_discount_amount = character(),
                                    actual_final_amount = character(),
                                    point_issue = character()
        )
        
        monthly_total <- data.frame(original_amount_total = 0,
                                    point_cashredeem_amount_total = 0,
                                    coupon_discount_amount_total = 0,
                                    actual_final_amount_total = 0,
                                    point_issue_total = 0)
    }
    
    # Rename column name for presentation purpose
    my_colnames <- c("消费日期", "会员卡号", "消费金额", "积分抵扣金额",
                     "优惠券抵扣金额", "实际支付金额", "实际累计积分")
    colnames(monthly_sales) <- my_colnames
    
    
    # if (nrow(monthly_sales) == 0){
    #     monthly_sales <- data.frame(member_id = numeric(),
    #                                 transaction_datetime = character()
    #                                 )
    # }
    
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
    
    # set some random shit on upper right corner
    rows <- createRow(sheet,rowIndex=7)
    sheetTitle <- createCell(rows, colIndex=8)
    setCellStyle(sheetTitle[[1,1]], CellStyle(outwb) + 
                     Font(outwb, isBold=FALSE, heightInPoints=8) +
                     Alignment(h="ALIGN_RIGHT"))
    setCellValue(sheetTitle[[1,1]], "（币种：人民币／RMB）")
    
    ## monthly total 
    total_row_position <- 8 + nrow(monthly_sales) + 1
    pt1_value <- c("总计：", "", "")
    pt2_value <- c(monthly_total$original_amount_total, monthly_total$point_cashredeem_amount_total,
                   monthly_total$coupon_discount_amount_total, monthly_total$actual_final_amount_total)
    pt3_value <- c(monthly_total$point_issue_total)
    #rows <- createRow(sheet,rowIndex=total_position)
    # cb <- CellBlock(sheet, startRow = total_row_position, startColumn = 1,
    #                 noRows = nrow(monthly_total), 
    #                 noColumns = (ncol(monthly_sales) + 1))
    
    #total_cells <- createCell(rows,colIndex=c(1:(ncol(monthly_sales) + 1)))
    # total_cells_values <- c("总计：", "", "",
    #                         monthly_total$original_amount_total, monthly_total$point_cashredeem_amount_total,
    #                         monthly_total$coupon_discount_amount_total, monthly_total$actual_final_amount_total,
    #                         monthly_total$point_issue_total)
    # total_cells_style <-  CellStyle(outwb, d/aataFormat=DataFormat("0.0")) + 
    #                              Font(outwb, isBold=FALSE, heightInPoints=10) +
    #                              Border(color="black",
    #                                     position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
    #                                     pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))
    pt1_style <-  CellStyle(outwb) + 
        Font(outwb, isBold=FALSE, heightInPoints=10) +
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))
    pt2_style <-  CellStyle(outwb, dataFormat=DataFormat("0.0")) + 
        Font(outwb, isBold=FALSE, heightInPoints=10) +
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))
    pt3_style <-  CellStyle(outwb, dataFormat=DataFormat("0")) + 
        Font(outwb, isBold=FALSE, heightInPoints=10) +
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) + 
        Alignment(h="ALIGN_RIGHT")
    
    pt1 <- CellBlock(sheet, startRow = total_row_position, 
                     startColumn = 1,
                     noRows = 1, 
                     noColumns = length(pt1_value))
    
    pt2 <- CellBlock(sheet, startRow = total_row_position, 
                     startColumn = length(pt1_value) + 1,
                     noRows = 1, 
                     noColumns = length(pt2_value))
    
    pt3 <- CellBlock(sheet, startRow = total_row_position, 
                     startColumn = length(c(pt1_value, pt2_value)) + 1,
                     noRows = 1, 
                     noColumns = length(pt3_value))
    
    CB.setMatrixData(pt1, matrix(pt1_value, nrow = 1), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = pt1_style)
    CB.setMatrixData(pt2, matrix(pt2_value, nrow = 1), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = pt2_style)
    CB.setMatrixData(pt3, matrix(pt3_value, nrow = 1), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = pt3_style)
    
    # mapply(setCellStyle, total_cells[1,], total_cells_style)
    # apply(total_cells, 1:2, setCellStyle, total_cells_style)
    # mapply(setCellValue, total_cells, total_cells_values)
    
    # foot note
    row_position <- 8 + nrow(monthly_sales) + 3
    
    note_value <- c("备注：", 
                    "1. 中国新天地每月10日前向合作商户提供上月的积分对账单，合作商户请于每月15日前进行确认并通知中国新天地，逾期则视为对中国新天地所提供的对账单无异议。",
                    "2. 如合作商户对积分对账单有任何问题，请联系中国新天地 企业传讯及推广部 客户关系管理组，邮件: itiandi@xintiandi.com")
    rows <- createRow(sheet,rowIndex=c(row_position:(row_position+2)))
    cells <- createCell(rows,colIndex=1)
    mapply(setCellValue, cells, note_value)
    
    
    # prep folder 
    base_dir <- getwd()
    # base_dir <- "/home/ubuntu/src/monthly_sales_report"
    # base_dir <- "/home/zaiming/src/monthly_sales_report"
    base_dir <- normalizePath("~/src/monthly_sales_report/")
    file_name <- paste(current_shop$shop_code, '.xlsx', sep = '')
    output_dir <- paste(base_dir, 
                        paste(substr(year(start_time), start = 3, stop = 4),
                        ifelse(nchar(month(start_time)) == 1, paste("0", month(start_time), sep = ""), month(start_time)), sep = ''), 
                        file_name, sep = '/')
    
    # create dir if not exists
    if (!file.exists(file.path(output_dir))){
        dir.create(file.path(dirname(output_dir)))
    }
    # write.file
    saveWorkbook(outwb, output_dir)
}

