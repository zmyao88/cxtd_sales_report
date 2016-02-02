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
    report_month <- paste(format(ymd(pg_start_time), format="%Y%m%d"),
                          format(ymd(pg_end_time) - ddays(1), format="%Y%m%d"),
                          sep = "-")
    report_title <- "i天地合作商户对账单"
    
    #shop_name 
    shop_name <- ifelse(is.null(current_shop$name_sc), "", current_shop$name_sc)
    shop_address <- ifelse(is.null(current_shop$address_sc), "", current_shop$address_sc) 
    report_id <- paste('DZ',
                       substr(current_shop$shop_code, start = 1, stop = 6),
                       substr(year(start_time), start = 3, stop = 4),
                       ifelse(nchar(month(start_time)) == 1, paste("0", month(start_time), sep = ""), month(start_time)), 
                       substrRight(current_shop$shop_code, n = 3),
                       sep = "")
    contact_person <- ifelse(is.null(current_partner$contact_person_1), "", current_partner$contact_person_1)  
    mailling_address <- ifelse(is.null(current_partner$address_sc), "", current_partner$address_sc) 
    company_name <- ifelse(is.null(current_partner$name_sc), "", current_partner$name_sc)  

    
    
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
        mutate(point_cost = point_issue / 100) %>% 
        collect()
    
    # calculate total
    if(nrow(monthly_sales) > 0){
        monthly_total <- monthly_sales %>% 
            summarise(original_amount_total = sum(original_amount, na.rm = T),
                      point_cashredeem_amount_total = sum(point_cashredeem_amount, na.rm = T),
                      coupon_discount_amount_total = sum(coupon_discount_amount, na.rm = T),
                      actual_final_amount_total = sum(actual_final_amount, na.rm = T),
                      point_issue_total = sum(point_issue, na.rm = T),
                      point_cost_total = sum(point_cost, na.rm = T))
    }else{
        # create "fake monthly_total" if there is no monthly_sales_data
        monthly_sales <- data.frame(transaction_datetime = character(),
                                    member_card_no = character(),
                                    original_amount = character(),
                                    point_cashredeem_amount = character(),
                                    coupon_discount_amount = character(),
                                    actual_final_amount = character(),
                                    point_issue = character(),
                                    point_cost = character()
        )
        
        monthly_total <- data.frame(original_amount_total = 0,
                                    point_cashredeem_amount_total = 0,
                                    coupon_discount_amount_total = 0,
                                    actual_final_amount_total = 0,
                                    point_issue_total = 0,
                                    point_cost_total = 0)
    }
    
    # Rename column name for presentation purpose
    my_colnames <- c("消费日期", "会员卡号", "消费金额", "积分抵扣金额",
                     "优惠券抵扣金额", "实际支付金额", "实际累计积分", "实际产生积分成本")
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
                                            isBold=TRUE, , name = "Microsoft YaHei")
    csSheetSubTitle <- CellStyle(outwb) + Font(outwb, heightInPoints=12,
                                               isItalic=TRUE, isBold=FALSE, name = "Microsoft YaHei")
    # column and row name style
    csTableRowNames <- CellStyle(outwb) + 
        Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") + 
        Alignment(wrapText=TRUE,
                  h="ALIGN_CENTER") + 
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))
    
    csTableColNames <- CellStyle(outwb) + 
        Font(outwb, isBold=TRUE, heightInPoints = 10, name = "Microsoft YaHei") +
        Alignment(wrapText=TRUE,
                  h="ALIGN_CENTER") + 
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) +
        Fill(backgroundColor="blue")
    
    # define column stytle
    csDeclColumn <- CellStyle(outwb, dataFormat=DataFormat("0.0")) +  
        Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) 
    
    csIntColumn <- CellStyle(outwb, dataFormat=DataFormat("0")) +
        Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) 
    
    csDateColumn <- CellStyle(outwb, dataFormat=DataFormat("YYYY-MM-DD")) + 
        Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) + 
        Alignment(wrapText=FALSE, horizontal="ALIGN_CENTER")
    
    csCardNoColumn <- CellStyle(outwb, dataFormat=DataFormat("0")) +
        Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) + 
        Alignment(wrapText=FALSE, horizontal="ALIGN_CENTER")
    
    
    # column index and correspoinding format
    date_col = list('1' = csDateColumn)
    int_col = list('2' = csCardNoColumn,
                   '7' = csIntColumn)
    dec_col = list('3' = csDeclColumn,
                   '4' = csDeclColumn,
                   '5' = csDeclColumn,
                   '6' = csDeclColumn,
                   '8' = csDeclColumn)
    
    # create a sheet
    sheet <- createSheet(outwb, sheetName = "Sales Report")
    # Add the table calculated above to the new sheet
    addDataFrame(monthly_sales, sheet, startRow=8,startColumn=1, 
                 colStyle= c(date_col, int_col, dec_col),
                 colnamesStyle = csTableColNames, 
                 rownamesStyle = csTableRowNames)
    # set column width
    
    setColumnWidth(sheet, colIndex=c(1, 2, 3, 9), colWidth=17)
    setColumnWidth(sheet, colIndex=c(4:5, 7:8), colWidth=12)
    setColumnWidth(sheet, colIndex=c(6), colWidth=14)
   
    ### BIG HACKS HERE 
    rows <- createRow(sheet,rowIndex=8)
    magic_cell <- createCell(rows, colIndex=1)
    setCellStyle(magic_cell[[1,1]], csTableColNames)
    
    magic_cb <- CellBlock(sheet, startRow = 8, 
                         startColumn = 2,
                         noRows = 1, 
                         noColumns = 8)
    CB.setMatrixData(magic_cb, matrix(my_colnames, nrow = 1), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = csTableColNames)
    ### BIG HACKS HERE 
    
    ### set Titles
    rows <- createRow(sheet,rowIndex=1)
    sheetTitle <- createCell(rows, colIndex=1)
    setCellValue(sheetTitle[[1,1]], report_title)
    setCellStyle(sheetTitle[[1,1]], csSheetTitle)
    
    ## set var names col
    ## Meta data Block
    # rows <- createRow(sheet,rowIndex=c(3:5))
    # cells <- createCell(rows,colIndex=c(1:8))
    
    meta_value_pt1 <- c("公司名称：", "商铺名称：", "联系人：",
                company_name, shop_name, contact_person)
    meta_value_pt2 <- c("商铺位置：", "邮件地址：",
                shop_address, mailling_address)
    meta_value_pt3 <- c("账单周期：",  "账单编号：","制单日期：",
                        report_month,  report_id, report_date)
    meta_style <- CellStyle(outwb) + 
                Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") +
                Alignment(h="ALIGN_LEFT")
                
    meta_cb1 <- CellBlock(sheet, startRow = 3, 
                    startColumn = 1,
                    noRows = 3, 
                    noColumns = 2)
    meta_cb2 <- CellBlock(sheet, startRow = 4, 
                          startColumn = 3,
                          noRows = 2, 
                          noColumns = 2)
    meta_cb3 <- CellBlock(sheet, startRow = 3, 
                          startColumn = 7,
                          noRows = 3, 
                          noColumns = 2)
    
    CB.setMatrixData(meta_cb1, matrix(meta_value_pt1, nrow = 3), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = meta_style)  
    CB.setMatrixData(meta_cb2, matrix(meta_value_pt2, nrow = 2), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = meta_style)  
    
    CB.setMatrixData(meta_cb3, matrix(meta_value_pt3, nrow = 3), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = meta_style)  
#     values <- c("公司名称：", "商铺名称：", "联系人：",
#                 company_name, shop_name, contact_person,
#                 "", "商铺位置：", "邮件地址：",
#                 "", shop_address, mailling_address,
#                 "", "", "",
#                 "", "", "",
#                 "账单周期：",  "账单编号：","制单日期：",
#                 report_month,  report_id, report_date
#     )
#     # mapply(setCellValue, cells, values)
#     
#     bs_style <- CellStyle(outwb) + 
#         Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") +
#         Alignment(h="ALIGN_LEFT")
#     
#     bs <- CellBlock(sheet, startRow = 3, 
#               startColumn = 1,
#               noRows = 3, 
#               noColumns = 8)
#     CB.setMatrixData(bs, matrix(values, nrow = 3), 
#                      startRow = 1, 
#                      startColumn = 1, 
#                      cellStyle = bs_style)
    
    ## Meta data Block Done
    # set some random shit on upper right corner
    rows <- createRow(sheet,rowIndex=7)
    sheetTitle <- createCell(rows, colIndex=9)
    setCellStyle(sheetTitle[[1,1]], CellStyle(outwb) + 
                     Font(outwb, isBold=FALSE, heightInPoints=8, name = "Microsoft YaHei") +
                     Alignment(h="ALIGN_RIGHT"))
    setCellValue(sheetTitle[[1,1]], "（币种：人民币／RMB）")
    
    ## monthly total 
    total_row_position <- 8 + nrow(monthly_sales) + 1
    pt1_value <- c("总计：", "", "")
    pt2_value <- c(monthly_total$original_amount_total, monthly_total$point_cashredeem_amount_total,
                   monthly_total$coupon_discount_amount_total, monthly_total$actual_final_amount_total)
    pt3_value <- c(monthly_total$point_issue_total)
    pt4_value <- c(monthly_total$point_cost_total)
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
        Font(outwb, isBold=TRUE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))
    pt2_style <-  CellStyle(outwb, dataFormat=DataFormat("0.0")) + 
        Font(outwb, isBold=TRUE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))
    pt3_style <-  CellStyle(outwb, dataFormat=DataFormat("0")) + 
        Font(outwb, isBold=TRUE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" )) + 
        Alignment(h="ALIGN_RIGHT")
    pt4_style <-  CellStyle(outwb, dataFormat=DataFormat("0.0")) + 
        Font(outwb, isBold=TRUE, heightInPoints=10, name = "Microsoft YaHei") +  
        Border(color="black",
               position=c("TOP", "BOTTOM", "LEFT", "RIGHT"),
               pen=c("BORDER_THIN", "BORDER_THIN", "BORDER_THIN", "BORDER_THIN" ))
    
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
    
    pt4 <- CellBlock(sheet, startRow = total_row_position, 
                     startColumn = length(c(pt1_value, pt2_value, pt3_value)) + 1,
                     noRows = 1, 
                     noColumns = length(pt4_value))
    
    
    
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
    CB.setMatrixData(pt4, matrix(pt4_value, nrow = 1), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = pt4_style)
    
    
    # mapply(setCellStyle, total_cells[1,], total_cells_style)
    # apply(total_cells, 1:2, setCellStyle, total_cells_style)
    # mapply(setCellValue, total_cells, total_cells_values)
    
    # foot note
    row_position <- 8 + nrow(monthly_sales) + 3
    
    note_value <- c("备注：", 
                    "1. 中国新天地每月10日前向合作商户提供上月的积分对账单，合作商户请于每月15日前进行确认并通知中国新天地，逾期则视为对中国新天地所提供的对账单无异议。",
                    "2. 如合作商户对积分对账单有任何问题，请联系中国新天地 企业传讯及推广部 客户关系管理组，邮件: itiandi@xintiandi.com")
    # rows <- createRow(sheet,rowIndex=c(row_position:(row_position+2)))
    # cells <- createCell(rows,colIndex=1)
    # mapply(setCellValue, cells, note_value)
    
    note_style <- CellStyle(outwb) + 
        Font(outwb, isBold=FALSE, heightInPoints=10, name = "Microsoft YaHei") +
        Alignment(h="ALIGN_LEFT")
    
    note_cb <- CellBlock(sheet, startRow = row_position, 
              startColumn = 1,
              noRows = 3, 
              noColumns = 1)
    CB.setMatrixData(note_cb, matrix(note_value, nrow = 3), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = note_style)
    
    
    # confirmation
    row_position <- 8 + nrow(monthly_sales) + 3 + 3 + 4
    
    confirm_value <- c("商户负责人签字_______________________", 
                    "",
                    "",
                    "",
                    "签字日期_______________________")    
    
    confirm_cb <- CellBlock(sheet, startRow = row_position, 
              startColumn = 1,
              noRows = 5, 
              noColumns = 1)
    CB.setMatrixData(confirm_cb, matrix(confirm_value, nrow = 5), 
                     startRow = 1, 
                     startColumn = 1, 
                     cellStyle = note_style)
#     rows <- createRow(sheet,rowIndex=c(row_position:(row_position+length(confirm_value) - 1)))
#     cells <- createCell(rows,colIndex=1)
#     mapply(setCellValue, cells, confirm_value)
    
    
    
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
    
    download_dir <- paste(base_dir, 
                    paste(substr(year(start_time), start = 3, stop = 4),
                          ifelse(nchar(month(start_time)) == 1, paste("0", month(start_time), sep = ""), month(start_time)), sep = ''), 
                    "download",
                    paste(sub("[[:punct:]]", "", shop_name, perl = F), ".xlsx", sep=""), 
                    sep = '/')
    
    # create dir if not exists
    if (!file.exists(file.path(output_dir))){
        dir.create(file.path(dirname(output_dir)))
    }
    if (!file.exists(file.path(download_dir))){
        dir.create(file.path(dirname(download_dir)))
    }
    # write.file
    saveWorkbook(outwb, output_dir)
    saveWorkbook(outwb, download_dir)
    
}

