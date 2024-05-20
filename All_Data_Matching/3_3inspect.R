library(openxlsx)
library(dplyr)
# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)
#zb=1,nonfound
#zb=2,incorrect
#zb=3,missing
#zb=4,NAandWrong
#zb=0,correct


# 设置文件夹路径和文件名
input_folder_1 <- ""
input_folder_2 <- ""
output_folder <- ""

excel_file <- paste0(input_folder_2, "指标说明.xlsx")
output_text_file1 <- paste0(output_folder, "output_info_sdxm.txt")
output_text_file2 <- paste0(output_folder, "output_info_yzbm.txt")
output_text_file3 <- paste0(output_folder, "output_info_xzqdm.txt")

# 在这里假设有两个表格，你需要根据实际情况修改表格名称和列名
sheet1_1 <- read.xlsx(excel_file, sheet = "省地县代码98")
sheet1_2 <- read.xlsx(excel_file, sheet = "省地县代码00")
sheet1_3 <- read.xlsx(excel_file, sheet = "省地县代码07")
sheet1_4 <- read.xlsx(excel_file, sheet = "省地县代码13")
sheet2 <- read.xlsx(excel_file, sheet = "邮政编码")

# 打开文件以写入模式
file_con1 <- file(output_text_file1, "w")
file_con2 <- file(output_text_file2, "w")
file_con3 <- file(output_text_file3, "w")

# 需要处理的年份列表
years <- c("1998","1999","2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010") 
# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_1, "address_data_",year, ".csv")
  output_file1 <- paste0(output_folder, "ins_all_sdxm_",year,".xlsx")
  output_file2 <- paste0(output_folder, "ins_all_yzbm_",year,".xlsx")
  output_file3 <- paste0(output_folder, "ins_all_xzqdm_",year,".xlsx")
  
  if(year=="1998"||year=="1999"){
    sheet1 <- sheet1_1
  }else if(year=="2000"||year=="2001"||year=="2002"){
    sheet1 <- sheet1_2
  }else if(year=="2014"||year=="2013"){
    sheet1 <- sheet1_4
  }else{
    sheet1 <- sheet1_3
  }
  
  all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  # 查找包含非法字符的行和列
  illegal_rows <- which(apply(all_data, 1, function(row) any(grepl("[[:cntrl:]]", row))))
  illegal_cols <- which(apply(all_data, 2, function(col) any(grepl("[[:cntrl:]]", col))))
  
  # 打印非法字符所在的行和列
  print(paste("Illegal rows:", illegal_rows))
  print(paste("Illegal columns:", illegal_cols))
  
  # 删除包含非法字符的行
  # if (length(illegal_rows) > 0){
  #   # clean_data <- all_data[-illegal_rows, ]
  #   # all_data <- clean_data
  #   clean_data <- data.frame(apply(all_data, 2, function(col) gsub("[[:cntrl:]]", " ", col)))
  #   }
  
  all_data_df <- as.data.frame(all_data)
  all_data <- all_data_df
  
  # 创建一个Excel工作簿
  wb1 <- createWorkbook()
  wb2 <- createWorkbook()
  wb3 <- createWorkbook()
  
  # 创建一个新的工作表，并命名为当前年
  addWorksheet(wb1, "correct")
  addWorksheet(wb1, "incorrect")
  addWorksheet(wb1, "nonfound")
  addWorksheet(wb1, "missing")
  addWorksheet(wb1, "NAandWrong")
  
  addWorksheet(wb2, "correct")
  addWorksheet(wb2, "incorrect")
  addWorksheet(wb2, "nonfound")
  addWorksheet(wb2, "missing")
  addWorksheet(wb2, "NAandWrong")
  
  addWorksheet(wb3, "correct")
  addWorksheet(wb3, "incorrect")
  addWorksheet(wb3, "nonfound")
  addWorksheet(wb3, "missing")
  addWorksheet(wb3, "NAandWrong")
  
  # 创建两个空的数据框用于存放结果
  incorrect1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(incorrect1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  incorrect2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(incorrect2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  incorrect3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(incorrect3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  correct1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(correct1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  correct2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(correct2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  correct3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(correct3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  nonfound1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(nonfound1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  nonfound2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(nonfound2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  nonfound3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(nonfound3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  NAandWrong1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(NAandWrong1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  NAandWrong2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(NAandWrong2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  NAandWrong3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(NAandWrong3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  missing1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(missing1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  missing2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(missing2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  missing3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(missing3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  # 将所有空值替换为 "NA"
  all_data_na <- all_data %>% replace(is.na(.), "NA")
  
  # 选择特定的10个列
  selected_cols <- all_data_na %>%
    select(qymc,frdbxm, province, 方位, city, town, county, sdxm, yzbm, 行政区代码)
  
  selected_cols_rows <- nrow(selected_cols)
  
  # 逐行提取sdxm、yzbm和行政区代码，并命名
  for (i in 1:nrow(selected_cols)) {
    row <- selected_cols[i, ]
    row_new <- row
    information <- c()
    information2 <- c()
    
    # 初始化一个空字符向量来存储有效值
    valid_information <- character()
    valid_information2 <- character()
    
    for (col_name in c("province", "city", "town", "county")) {
      col_values <- row[[col_name]]
      # 选择除了 "NA", "wrong province", "wrong city" 之外的值
      valid_values <- col_values[!(col_values %in% c("NA", "wrong province", "wrong city"))]
      # 将这些值添加到 valid_information 中
      valid_information <- c(valid_information, valid_values)
    }
    
    # 将 valid_information 合并成一个长字符串，用 "|" 分隔
    information <- paste(valid_information, collapse = "|")
    
    for (col_name in c("city", "town", "county")) {
      col_values <- row[[col_name]]
      # 选择除了 "NA", "wrong province", "wrong city" 之外的值
      valid_values <- col_values[!(col_values %in% c("NA", "wrong province", "wrong city"))]
      # 将这些值添加到 valid_information 中
      valid_information2 <- c(valid_information2, valid_values)
    }
    
    # 将 valid_information 合并成一个长字符串，用 "|" 分隔
    information2 <- paste(valid_information2, collapse = "|")
    
    sdxm_individual <- row$sdxm
    yzbm_individual <- row$yzbm
    行政区代码_individual <- row$行政区代码
    
    zb <- 0
    
    # 判断sdxm_individual是否不是字符串"NA"
    if (sdxm_individual != "NA"&&sdxm_individual != "wrong sdxm") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(substr(sheet1$代码, 1, 6) == substr(sdxm_individual, 1, 6))
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$sdxm_reference<-"nonfound"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$sdxm_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(reference, information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          last_char_reference <- substr(reference, nchar(reference), nchar(reference))
          is_duplicate2 <- any(grepl(last_char_reference, information2))
          if (!is_duplicate2) {
          zb <- 3
          }else{
          zb <- 2
          }
        }
      }
    } else {
      zb <- 4
      row_new$sdxm_reference<-"NA"
      print("sdxm_individual 的值是 NA 或者 wrong sdxm")
    }
    
    if (zb == 0){
      correct1 <- rbind(correct1, row_new)
    } else if(zb == 1){
      nonfound1 <- rbind(nonfound1, row_new)
    } else if(zb == 2){
      incorrect1 <- rbind(incorrect1, row_new)
    } else if(zb == 3){
      missing1 <- rbind(missing1, row_new)
    } else if(zb == 4){
      NAandWrong1 <- rbind(NAandWrong1, row_new)
    } 
    
    zb <- 0
    zbb <- 2
    # 判断yzbm_individual是否不是字符串"NA"
    if (yzbm_individual != "NA"&&yzbm_individual != "wrong yzbm") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which((substr(sheet2$D, 1, 4) == substr(yzbm_individual, 1, 4) & nchar(sheet2$D) == nchar(yzbm_individual)) | 
                               (substr(sheet2$D, 1, 3) == substr(yzbm_individual, 1, 3) & nchar(sheet2$D) == nchar(yzbm_individual)))
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        zbb <- 1
      } else {
        for(j in matching_rows){
          reference1 <- sheet2$B[j]
          reference2 <- sheet2$A[j]
          reference <- paste(reference1, reference2, sep = " ")
          
          # 判断reference是否是information中某个值的子字符串
          is_duplicate1 <- any(grepl(reference1, information))#城市级是否匹配
          
          if (is_duplicate1) {
            zbb <- 0
            row_new$yzbm_reference<- reference1
          }
          # 判断information2是否包含"市"等字符
          last_char_reference <- substr(reference1, nchar(reference1), nchar(reference1))
          is_duplicate2 <- any(grepl(last_char_reference, information2))
          if(is_duplicate2){
            zb <- 2
          }
          }
        }
        if(zbb==0){
          zb <- 0
        }else if(zbb==1){
          zb <- 1
          row_new$yzbm_reference<-"nonfound"
        }else{
          # 判断information2是否为空值
          is_empty <- is.na(information2) || information2 == ""
          
          # 根据条件设置zb的值
          if (is_empty|| zb!=2) {
            zb <- 3
            row_new$yzbm_reference<- reference1
          } else {
            zb <- 2
            row_new$yzbm_reference<- reference1
          }
        
      }
    } else {
      zb <- 4
      row_new$yzbm_reference<-"NA"
      print("yzbm_individual 的值是 NA 或者 wrong yzbm")
    }
    
    if (zb == 0){
      correct2 <- rbind(correct2, row_new)
    } else if(zb == 1){
      nonfound2 <- rbind(nonfound2, row_new)
    } else if(zb == 2){
      incorrect2 <- rbind(incorrect2, row_new)
    } else if(zb == 3){
      missing2 <- rbind(missing2, row_new)
    } else if(zb == 4){
      NAandWrong2 <- rbind(NAandWrong2, row_new)
    } 
    
    zb <- 0
    # 判断行政区代码_individual是否不是字符串"NA"
    if (行政区代码_individual != "NA"&&行政区代码_individual != "wrong 行政区代码") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(sheet1$代码 == 行政区代码_individual)
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$行政区代码_reference<-"nonfound"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$行政区代码_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(reference, information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          last_char_reference <- substr(reference, nchar(reference), nchar(reference))
          is_duplicate2 <- any(grepl(last_char_reference, information2))
          if (!is_duplicate2) {
            zb <- 3
          }else{
            zb <- 2
          }
        }
      }
    } else {
      zb <- 4
      row_new$行政区代码_reference<-"NA"
      print("行政区代码_individual 的值是 NA 或者 wrong 行政区代码")
    }
    
    if (zb == 0){
      correct3 <- rbind(correct3, row_new)
    } else if(zb == 1){
      nonfound3 <- rbind(nonfound3, row_new)
    } else if(zb == 2){
      incorrect3 <- rbind(incorrect3, row_new)
    } else if(zb == 3){
      missing3 <- rbind(missing3, row_new)
    } else if(zb == 4){
      NAandWrong3 <- rbind(NAandWrong3, row_new)
    } 
    
   
  }
  # 计算正确和不正确数据框的行数
  nonfound_rows <- nrow(nonfound1)
  missing_rows <- nrow(missing1)
  correct_rows <- nrow(correct1)
  incorrect_rows <- nrow(incorrect1)
  NAandWrong_rows <- nrow(NAandWrong1)
  
  # 计算所有行数之和
  all_number <- correct_rows + incorrect_rows + nonfound_rows + missing_rows + NAandWrong_rows
  
  # 计算正确数据框占总行数的百分比
  percentage_nonfound <- (nonfound_rows / all_number) * 100
  percentage_missing <- (missing_rows / all_number) * 100
  percentage_correct <- (correct_rows / all_number) * 100
  percentage_incorrect <- (incorrect_rows / all_number) * 100
  percentage_NAandWrong <- (NAandWrong_rows / all_number) * 100
  
  # 打印百分比
  print(paste("正确数据占总数的百分比:", percentage_correct, "%"))
  print(paste("行数验证:",  selected_cols_rows, "=",all_number))
  
  # 将数据写入工作表
  writeData(wb1, sheet = "correct", correct1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "incorrect", incorrect1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "nonfound", nonfound1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "missing", missing1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "NAandWrong", NAandWrong1, startCol = 1, startRow = 1)
  
  # 保存工作簿为Excel文件
  saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
  
  # 将信息写入文件
  cat(paste("无法匹配占总数的百分比:", percentage_nonfound, "%"), "\n", file = file_con1)
  cat(paste("数据缺失占总数的百分比:", percentage_missing, "%"), "\n", file = file_con1)
  cat(paste("正确数据占总数的百分比:", percentage_correct, "%"), "\n", file = file_con1)
  cat(paste("错误数据占总数的百分比:", percentage_incorrect, "%"), "\n", file = file_con1)
  cat(paste("编码缺失占总数的百分比:", percentage_NAandWrong, "%"), "\n", file = file_con1)
  
  cat(paste("行数验证:", selected_cols_rows, "=", all_number), "\n", file = file_con1)
  cat(paste("完成", year, "的工作"), "\n", file = file_con1)
  
  # 计算正确和不正确数据框的行数
  nonfound_rows <- nrow(nonfound2)
  missing_rows <- nrow(missing2)
  correct_rows <- nrow(correct2)
  incorrect_rows <- nrow(incorrect2)
  NAandWrong_rows <- nrow(NAandWrong2)
  
  # 计算所有行数之和
  all_number <- correct_rows + incorrect_rows + nonfound_rows + missing_rows + NAandWrong_rows
  
  # 计算正确数据框占总行数的百分比
  percentage_nonfound <- (nonfound_rows / all_number) * 100
  percentage_missing <- (missing_rows / all_number) * 100
  percentage_correct <- (correct_rows / all_number) * 100
  percentage_incorrect <- (incorrect_rows / all_number) * 100
  percentage_NAandWrong <- (NAandWrong_rows / all_number) * 100
  
  # 打印百分比
  print(paste("正确数据占总数的百分比:", percentage_correct, "%"))
  print(paste("行数验证:",  selected_cols_rows, "=",all_number))
  
  # 将数据写入工作表
  writeData(wb2, sheet = "correct", correct2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "incorrect", incorrect2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "nonfound", nonfound2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "missing", missing2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "NAandWrong", NAandWrong2, startCol = 1, startRow = 1)
  
  # 保存工作簿为Excel文件
  saveWorkbook(wb2, file = output_file2, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
  
  # 将信息写入文件
  cat(paste("无法匹配占总数的百分比:", percentage_nonfound, "%"), "\n", file = file_con2)
  cat(paste("数据缺失占总数的百分比:", percentage_missing, "%"), "\n", file = file_con2)
  cat(paste("正确数据占总数的百分比:", percentage_correct, "%"), "\n", file = file_con2)
  cat(paste("错误数据占总数的百分比:", percentage_incorrect, "%"), "\n", file = file_con2)
  cat(paste("编码缺失占总数的百分比:", percentage_NAandWrong, "%"), "\n", file = file_con2)
  
  cat(paste("行数验证:", selected_cols_rows, "=", all_number), "\n", file = file_con2)
  cat(paste("完成", year, "的工作"), "\n", file = file_con2)
  
  # 计算正确和不正确数据框的行数
  nonfound_rows <- nrow(nonfound3)
  missing_rows <- nrow(missing3)
  correct_rows <- nrow(correct3)
  incorrect_rows <- nrow(incorrect3)
  NAandWrong_rows <- nrow(NAandWrong3)
  
  # 计算所有行数之和
  all_number <- correct_rows + incorrect_rows + nonfound_rows + missing_rows + NAandWrong_rows
  
  # 计算正确数据框占总行数的百分比
  percentage_nonfound <- (nonfound_rows / all_number) * 100
  percentage_missing <- (missing_rows / all_number) * 100
  percentage_correct <- (correct_rows / all_number) * 100
  percentage_incorrect <- (incorrect_rows / all_number) * 100
  percentage_NAandWrong <- (NAandWrong_rows / all_number) * 100
  
  # 打印百分比
  print(paste("正确数据占总数的百分比:", percentage_correct, "%"))
  print(paste("行数验证:",  selected_cols_rows, "=",all_number))
  
  # 将数据写入工作表
  writeData(wb3, sheet = "correct", correct3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "incorrect", incorrect3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "nonfound", nonfound3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "missing", missing3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "NAandWrong", NAandWrong3, startCol = 1, startRow = 1)
  
  # 保存工作簿为Excel文件
  saveWorkbook(wb3, file = output_file3, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
  
  # 将信息写入文件
  cat(paste("无法匹配占总数的百分比:", percentage_nonfound, "%"), "\n", file = file_con3)
  cat(paste("数据缺失占总数的百分比:", percentage_missing, "%"), "\n", file = file_con3)
  cat(paste("正确数据占总数的百分比:", percentage_correct, "%"), "\n", file = file_con3)
  cat(paste("错误数据占总数的百分比:", percentage_incorrect, "%"), "\n", file = file_con3)
  cat(paste("编码缺失占总数的百分比:", percentage_NAandWrong, "%"), "\n", file = file_con3)
  
  cat(paste("行数验证:", selected_cols_rows, "=", all_number), "\n", file = file_con3)
  cat(paste("完成", year, "的工作"), "\n", file = file_con3)
}

#对于2011至2014，用行政区划代码
# 需要处理的年份列表
years <- c("2011","2012","2013","2014") 
# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_1, "address_data_",year, ".csv")
  output_file1 <- paste0(output_folder, "ins_all_sdxm_",year,".xlsx")
  output_file2 <- paste0(output_folder, "ins_all_yzbm_",year,".xlsx")
  output_file3 <- paste0(output_folder, "ins_all_xzqdm_",year,".xlsx")
  
  if(year=="1998"||year=="1999"){
    sheet1 <- sheet1_1
  }else if(year=="2000"||year=="2001"||year=="2002"){
    sheet1 <- sheet1_2
  }else if(year=="2014"||year=="2013"){
    sheet1 <- sheet1_4
  }else{
    sheet1 <- sheet1_3
  }
  
  all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  # 查找包含非法字符的行和列
  illegal_rows <- which(apply(all_data, 1, function(row) any(grepl("[[:cntrl:]]", row))))
  illegal_cols <- which(apply(all_data, 2, function(col) any(grepl("[[:cntrl:]]", col))))
  
  # 打印非法字符所在的行和列
  print(paste("Illegal rows:", illegal_rows))
  print(paste("Illegal columns:", illegal_cols))
  
  # 删除包含非法字符的行
  # if (length(illegal_rows) > 0){
  #   # clean_data <- all_data[-illegal_rows, ]
  #   # all_data <- clean_data
  #   clean_data <- data.frame(apply(all_data, 2, function(col) gsub("[[:cntrl:]]", " ", col)))
  #   }
  
  all_data_df <- as.data.frame(all_data)
  all_data <- all_data_df
  
  # 创建一个Excel工作簿
  wb1 <- createWorkbook()
  wb2 <- createWorkbook()
  wb3 <- createWorkbook()
  
  # 创建一个新的工作表，并命名为当前年
  addWorksheet(wb1, "correct")
  addWorksheet(wb1, "incorrect")
  addWorksheet(wb1, "nonfound")
  addWorksheet(wb1, "missing")
  addWorksheet(wb1, "NAandWrong")
  
  addWorksheet(wb2, "correct")
  addWorksheet(wb2, "incorrect")
  addWorksheet(wb2, "nonfound")
  addWorksheet(wb2, "missing")
  addWorksheet(wb2, "NAandWrong")
  
  addWorksheet(wb3, "correct")
  addWorksheet(wb3, "incorrect")
  addWorksheet(wb3, "nonfound")
  addWorksheet(wb3, "missing")
  addWorksheet(wb3, "NAandWrong")
  
  # 创建两个空的数据框用于存放结果
  incorrect1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(incorrect1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  incorrect2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(incorrect2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  incorrect3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(incorrect3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  correct1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(correct1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  correct2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(correct2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  correct3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(correct3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  nonfound1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(nonfound1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  nonfound2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(nonfound2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  nonfound3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(nonfound3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  NAandWrong1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(NAandWrong1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  NAandWrong2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(NAandWrong2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  NAandWrong3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(NAandWrong3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  missing1 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(missing1) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  missing2 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(missing2) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  missing3 <- data.frame(matrix(ncol = 13, nrow = 0))
  colnames(missing3) <- c("qymc","frdbxm","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区划代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  # 将所有空值替换为 "NA"
  all_data_na <- all_data %>% replace(is.na(.), "NA")
  
  # 选择特定的10个列
  selected_cols <- all_data_na %>%
    select(qymc,frdbxm, province, 方位, city, town, county, sdxm, yzbm, 行政区划代码)
  
  selected_cols_rows <- nrow(selected_cols)
  
  # 逐行提取sdxm、yzbm和行政区代码，并命名
  for (i in 1:nrow(selected_cols)) {
    row <- selected_cols[i, ]
    row_new <- row
    information <- c()
    information2 <- c()
    
    # 初始化一个空字符向量来存储有效值
    valid_information <- character()
    valid_information2 <- character()
    
    for (col_name in c("province", "city", "town", "county")) {
      col_values <- row[[col_name]]
      # 选择除了 "NA", "wrong province", "wrong city" 之外的值
      valid_values <- col_values[!(col_values %in% c("NA", "wrong province", "wrong city"))]
      # 将这些值添加到 valid_information 中
      valid_information <- c(valid_information, valid_values)
    }
    
    # 将 valid_information 合并成一个长字符串，用 "|" 分隔
    information <- paste(valid_information, collapse = "|")
    
    for (col_name in c("city", "town", "county")) {
      col_values <- row[[col_name]]
      # 选择除了 "NA", "wrong province", "wrong city" 之外的值
      valid_values <- col_values[!(col_values %in% c("NA", "wrong province", "wrong city"))]
      # 将这些值添加到 valid_information 中
      valid_information2 <- c(valid_information2, valid_values)
    }
    
    # 将 valid_information 合并成一个长字符串，用 "|" 分隔
    information2 <- paste(valid_information2, collapse = "|")
    
    sdxm_individual <- row$sdxm
    yzbm_individual <- row$yzbm
    行政区代码_individual <- row$行政区划代码
    
    zb <- 0
    
    # 判断sdxm_individual是否不是字符串"NA"
    if (sdxm_individual != "NA"&&sdxm_individual != "wrong sdxm") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(substr(sheet1$代码, 1, 6) == substr(sdxm_individual, 1, 6))
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$sdxm_reference<-"nonfound"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$sdxm_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(reference, information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          last_char_reference <- substr(reference, nchar(reference), nchar(reference))
          is_duplicate2 <- any(grepl(last_char_reference, information2))
          if (!is_duplicate2) {
            zb <- 3
          }else{
            zb <- 2
          }
        }
      }
    } else {
      zb <- 4
      row_new$sdxm_reference<-"NA"
      print("sdxm_individual 的值是 NA 或者 wrong sdxm")
    }
    
    if (zb == 0){
      correct1 <- rbind(correct1, row_new)
    } else if(zb == 1){
      nonfound1 <- rbind(nonfound1, row_new)
    } else if(zb == 2){
      incorrect1 <- rbind(incorrect1, row_new)
    } else if(zb == 3){
      missing1 <- rbind(missing1, row_new)
    } else if(zb == 4){
      NAandWrong1 <- rbind(NAandWrong1, row_new)
    } 
    
    zb <- 0
    zbb <- 2
    # 判断yzbm_individual是否不是字符串"NA"
    if (yzbm_individual != "NA"&&yzbm_individual != "wrong yzbm") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which((substr(sheet2$D, 1, 4) == substr(yzbm_individual, 1, 4) & nchar(sheet2$D) == nchar(yzbm_individual)) | 
                               (substr(sheet2$D, 1, 3) == substr(yzbm_individual, 1, 3) & nchar(sheet2$D) == nchar(yzbm_individual)))
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        zbb <- 1
      } else {
        for(j in matching_rows){
          reference1 <- sheet2$B[j]
          reference2 <- sheet2$A[j]
          reference <- paste(reference1, reference2, sep = " ")
          
          # 判断reference是否是information中某个值的子字符串
          is_duplicate1 <- any(grepl(reference1, information))#城市级是否匹配
          
          if (is_duplicate1) {
            zbb <- 0
            row_new$yzbm_reference<- reference1
          }
          # 判断information2是否包含"市"等字符
          last_char_reference <- substr(reference1, nchar(reference1), nchar(reference1))
          is_duplicate2 <- any(grepl(last_char_reference, information2))
          if(is_duplicate2){
            zb <- 2
          }
        }
      }
      if(zbb==0){
        zb <- 0
      }else if(zbb==1){
        zb <- 1
        row_new$yzbm_reference<-"nonfound"
      }else{
        # 判断information2是否为空值
        is_empty <- is.na(information2) || information2 == ""
        
        # 根据条件设置zb的值
        if (is_empty|| zb!=2) {
          zb <- 3
          row_new$yzbm_reference<- reference1
        } else {
          zb <- 2
          row_new$yzbm_reference<- reference1
        }
        
      }
    } else {
      zb <- 4
      row_new$yzbm_reference<-"NA"
      print("yzbm_individual 的值是 NA 或者 wrong yzbm")
    }
    
    if (zb == 0){
      correct2 <- rbind(correct2, row_new)
    } else if(zb == 1){
      nonfound2 <- rbind(nonfound2, row_new)
    } else if(zb == 2){
      incorrect2 <- rbind(incorrect2, row_new)
    } else if(zb == 3){
      missing2 <- rbind(missing2, row_new)
    } else if(zb == 4){
      NAandWrong2 <- rbind(NAandWrong2, row_new)
    } 
    
    zb <- 0
    # 判断行政区代码_individual是否不是字符串"NA"
    if (行政区代码_individual != "NA"&&行政区代码_individual != "wrong 行政区代码") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(sheet1$代码 == 行政区代码_individual)
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$行政区代码_reference<-"nonfound"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$行政区代码_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(reference, information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          last_char_reference <- substr(reference, nchar(reference), nchar(reference))
          is_duplicate2 <- any(grepl(last_char_reference, information2))
          if (!is_duplicate2) {
            zb <- 3
          }else{
            zb <- 2
          }
        }
      }
    } else {
      zb <- 4
      row_new$行政区代码_reference<-"NA"
      print("行政区代码_individual 的值是 NA 或者 wrong 行政区代码")
    }
    
    if (zb == 0){
      correct3 <- rbind(correct3, row_new)
    } else if(zb == 1){
      nonfound3 <- rbind(nonfound3, row_new)
    } else if(zb == 2){
      incorrect3 <- rbind(incorrect3, row_new)
    } else if(zb == 3){
      missing3 <- rbind(missing3, row_new)
    } else if(zb == 4){
      NAandWrong3 <- rbind(NAandWrong3, row_new)
    } 
    
    
  }
  # 计算正确和不正确数据框的行数
  nonfound_rows <- nrow(nonfound1)
  missing_rows <- nrow(missing1)
  correct_rows <- nrow(correct1)
  incorrect_rows <- nrow(incorrect1)
  NAandWrong_rows <- nrow(NAandWrong1)
  
  # 计算所有行数之和
  all_number <- correct_rows + incorrect_rows + nonfound_rows + missing_rows + NAandWrong_rows
  
  # 计算正确数据框占总行数的百分比
  percentage_nonfound <- (nonfound_rows / all_number) * 100
  percentage_missing <- (missing_rows / all_number) * 100
  percentage_correct <- (correct_rows / all_number) * 100
  percentage_incorrect <- (incorrect_rows / all_number) * 100
  percentage_NAandWrong <- (NAandWrong_rows / all_number) * 100
  
  # 打印百分比
  print(paste("正确数据占总数的百分比:", percentage_correct, "%"))
  print(paste("行数验证:",  selected_cols_rows, "=",all_number))
  
  # 将数据写入工作表
  writeData(wb1, sheet = "correct", correct1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "incorrect", incorrect1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "nonfound", nonfound1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "missing", missing1, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "NAandWrong", NAandWrong1, startCol = 1, startRow = 1)
  
  # 保存工作簿为Excel文件
  saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
  
  # 将信息写入文件
  cat(paste("无法匹配占总数的百分比:", percentage_nonfound, "%"), "\n", file = file_con1)
  cat(paste("数据缺失占总数的百分比:", percentage_missing, "%"), "\n", file = file_con1)
  cat(paste("正确数据占总数的百分比:", percentage_correct, "%"), "\n", file = file_con1)
  cat(paste("错误数据占总数的百分比:", percentage_incorrect, "%"), "\n", file = file_con1)
  cat(paste("编码缺失占总数的百分比:", percentage_NAandWrong, "%"), "\n", file = file_con1)
  
  cat(paste("行数验证:", selected_cols_rows, "=", all_number), "\n", file = file_con1)
  cat(paste("完成", year, "的工作"), "\n", file = file_con1)
  
  # 计算正确和不正确数据框的行数
  nonfound_rows <- nrow(nonfound2)
  missing_rows <- nrow(missing2)
  correct_rows <- nrow(correct2)
  incorrect_rows <- nrow(incorrect2)
  NAandWrong_rows <- nrow(NAandWrong2)
  
  # 计算所有行数之和
  all_number <- correct_rows + incorrect_rows + nonfound_rows + missing_rows + NAandWrong_rows
  
  # 计算正确数据框占总行数的百分比
  percentage_nonfound <- (nonfound_rows / all_number) * 100
  percentage_missing <- (missing_rows / all_number) * 100
  percentage_correct <- (correct_rows / all_number) * 100
  percentage_incorrect <- (incorrect_rows / all_number) * 100
  percentage_NAandWrong <- (NAandWrong_rows / all_number) * 100
  
  # 打印百分比
  print(paste("正确数据占总数的百分比:", percentage_correct, "%"))
  print(paste("行数验证:",  selected_cols_rows, "=",all_number))
  
  # 将数据写入工作表
  writeData(wb2, sheet = "correct", correct2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "incorrect", incorrect2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "nonfound", nonfound2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "missing", missing2, startCol = 1, startRow = 1)
  writeData(wb2, sheet = "NAandWrong", NAandWrong2, startCol = 1, startRow = 1)
  
  # 保存工作簿为Excel文件
  saveWorkbook(wb2, file = output_file2, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
  
  # 将信息写入文件
  cat(paste("无法匹配占总数的百分比:", percentage_nonfound, "%"), "\n", file = file_con2)
  cat(paste("数据缺失占总数的百分比:", percentage_missing, "%"), "\n", file = file_con2)
  cat(paste("正确数据占总数的百分比:", percentage_correct, "%"), "\n", file = file_con2)
  cat(paste("错误数据占总数的百分比:", percentage_incorrect, "%"), "\n", file = file_con2)
  cat(paste("编码缺失占总数的百分比:", percentage_NAandWrong, "%"), "\n", file = file_con2)
  
  cat(paste("行数验证:", selected_cols_rows, "=", all_number), "\n", file = file_con2)
  cat(paste("完成", year, "的工作"), "\n", file = file_con2)
  
  # 计算正确和不正确数据框的行数
  nonfound_rows <- nrow(nonfound3)
  missing_rows <- nrow(missing3)
  correct_rows <- nrow(correct3)
  incorrect_rows <- nrow(incorrect3)
  NAandWrong_rows <- nrow(NAandWrong3)
  
  # 计算所有行数之和
  all_number <- correct_rows + incorrect_rows + nonfound_rows + missing_rows + NAandWrong_rows
  
  # 计算正确数据框占总行数的百分比
  percentage_nonfound <- (nonfound_rows / all_number) * 100
  percentage_missing <- (missing_rows / all_number) * 100
  percentage_correct <- (correct_rows / all_number) * 100
  percentage_incorrect <- (incorrect_rows / all_number) * 100
  percentage_NAandWrong <- (NAandWrong_rows / all_number) * 100
  
  # 打印百分比
  print(paste("正确数据占总数的百分比:", percentage_correct, "%"))
  print(paste("行数验证:",  selected_cols_rows, "=",all_number))
  
  # 将数据写入工作表
  writeData(wb3, sheet = "correct", correct3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "incorrect", incorrect3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "nonfound", nonfound3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "missing", missing3, startCol = 1, startRow = 1)
  writeData(wb3, sheet = "NAandWrong", NAandWrong3, startCol = 1, startRow = 1)
  
  # 保存工作簿为Excel文件
  saveWorkbook(wb3, file = output_file3, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
  
  # 将信息写入文件
  cat(paste("无法匹配占总数的百分比:", percentage_nonfound, "%"), "\n", file = file_con3)
  cat(paste("数据缺失占总数的百分比:", percentage_missing, "%"), "\n", file = file_con3)
  cat(paste("正确数据占总数的百分比:", percentage_correct, "%"), "\n", file = file_con3)
  cat(paste("错误数据占总数的百分比:", percentage_incorrect, "%"), "\n", file = file_con3)
  cat(paste("编码缺失占总数的百分比:", percentage_NAandWrong, "%"), "\n", file = file_con3)
  
  cat(paste("行数验证:", selected_cols_rows, "=", all_number), "\n", file = file_con3)
  cat(paste("完成", year, "的工作"), "\n", file = file_con3)
}
# 关闭文件连接
close(file_con1)
close(file_con2)
close(file_con3)

# 打印提示信息
print("信息已写入文件。")
