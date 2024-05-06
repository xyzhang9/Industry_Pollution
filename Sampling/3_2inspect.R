library(openxlsx)
library(dplyr)
# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)

# 设置文件夹路径和文件名
input_folder_1 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算3/地址/address_annual/"
input_folder_2 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算3/地址/inspect/"

excel_file <- paste0(input_folder_2, "指标说明.xlsx")

# 需要处理的年份列表
years <- c("1999","2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010","2013") 
# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_1, "address_data_",year, ".csv")
  output_file1 <- paste0(output_folder, "inspect_",year,".xlsx")
  
  all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  all_data_df <- as.data.frame(all_data)
  all_data <- all_data_df
  
  # 创建一个Excel工作簿
  wb1 <- createWorkbook()
  
  # 创建一个新的工作表，并命名为当前年
  addWorksheet(wb1, "sample")
  addWorksheet(wb1, "completed")
  addWorksheet(wb1, "correct")
  addWorksheet(wb1, "incorrect")
  addWorksheet(wb1, "manual")
  
  # 创建两个空的数据框用于存放结果
  incorrect <- data.frame(matrix(ncol = 11, nrow = 0))
  colnames(incorrect) <- c("province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  correct <- data.frame(matrix(ncol = 11, nrow = 0))
  colnames(correct) <- c("province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  # 随机提取50行数据
  sample <- all_data %>% sample_n(50)
  
  # 将所有空值替换为 "NA"
  sample_na <- sample %>% replace(is.na(.), "NA")
  
  # 选择特定的8个列
  selected_cols <- sample_na %>%
    select(qymc,frdbxm, province, 方位, city, town, county, sdxm, yzbm, 行政区代码)
  
  # 挑出这8个列中有7个不是 "NA" 且不等于 "wrong province" 等的行
  filtered_rows <- selected_cols %>%
    filter(rowSums(. != "NA" & . != "wrong province"& . != "wrong city" 
                   & . != "wrong sdxm"& . != "wrong yzbm"
                   & . != "wrong 行政区代码",na.rm = TRUE) >= 8)
  
  # 计算符合条件的行数占总行数的比例
  percentage <- nrow(filtered_rows) / nrow(sample_na) * 100
  
  # 输出结果
  print(paste("相对完整的数据行数的占比为:", percentage, "%"))
  
  
  # 在这里假设有两个表格，你需要根据实际情况修改表格名称和列名
  sheet1 <- read.xlsx(excel_file, sheet = "省地县代码")
  sheet2 <- read.xlsx(excel_file, sheet = "全国")
  
  # 逐行提取sdxm、yzbm和行政区代码，并命名
  for (i in 1:nrow(filtered_rows)) {
    row <- filtered_rows[i, ]
    # 创建一个仅包含所需列的新行 row_new
    row_new <- row[, c("qymc","frdbxm", "province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码")]
    
    zb <- 0
    information <- c()
    
    # 初始化一个空字符向量来存储有效值
    valid_information <- character()
    
    for (col_name in c("province", "city", "town", "county")) {
      col_values <- row[[col_name]]
      # 选择除了 "NA", "wrong province", "wrong city" 之外的值
      valid_values <- col_values[!(col_values %in% c("NA", "wrong province", "wrong city"))]
      # 将这些值添加到 valid_information 中
      valid_information <- c(valid_information, valid_values)
    }
    
    # 将 valid_information 合并成一个长字符串，用 "|" 分隔
    information <- paste(valid_information, collapse = "|")
    
    sdxm_individual <- row$sdxm
    yzbm_individual <- row$yzbm
    行政区代码_individual <- row$行政区代码
    
    # 判断sdxm_individual是否不是字符串"NA"
    if (sdxm_individual != "NA"&&sdxm_individual != "wrong sdxm") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(substr(sheet1$代码, 1, 6) == substr(sdxm_individual, 1, 6))
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$sdxm_reference<-"NA"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$sdxm_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(paste(reference, collapse = "|"), information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          zb <- 1
        }
      }
    } else {
      row_new$sdxm_reference<-"NA"
      print("sdxm_individual 的值是 NA 或者 wrong sdxm")
    }
    
    # 判断yzbm_individual是否不是字符串"NA"
    if (yzbm_individual != "NA"&&yzbm_individual != "wrong yzbm") {
      zbb <- 1
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(substr(sheet2$D, 1, 4) == substr(yzbm_individual, 1, 4) & nchar(sheet2$D) == nchar(yzbm_individual))
      matching_rows <- which(substr(sheet2$D, 1, 3) == substr(yzbm_individual, 1, 3) & nchar(sheet2$D) == nchar(yzbm_individual))
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$yzbm_reference<-"NA"
      } else {
        for(j in matching_rows){
          reference <- sheet2$B[j]
          
          # 判断reference是否是information中某个值的子字符串
          is_duplicate <- any(grepl(paste(reference, collapse = "|"), information))
          is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
          
          if (is_duplicate) {
            zbb <- 0
            row_new$yzbm_reference<- reference
          }
        }
        if(zbb==1){
          zb <- 1
          row_new$yzbm_reference<-"NA"
        }
       }
      } else {
      row_new$yzbm_reference<-"NA"
      print("yzbm_individual 的值是 NA 或者 wrong yzbm")
    }
    
    # 判断行政区代码_individual是否不是字符串"NA"
    if (行政区代码_individual != "NA"&&行政区代码_individual != "wrong 行政区代码") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(sheet1$代码 == 行政区代码_individual)
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$行政区代码_reference<-"NA"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$行政区代码_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(paste(reference, collapse = "|"), information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          zb <- 1
        }
      }
    } else {
      row_new$行政区代码_reference<-"NA"
      print("行政区代码_individual 的值是 NA 或者 wrong 行政区代码")
    }
    if (row_new$行政区代码_reference == "NA" &&
        row_new$yzbm_reference == "NA" &&
        row_new$sdxm_reference == "NA"){
      zb <- 1
    }
    
    if (zb == 0){
      correct <- rbind(correct, row_new)
    } else{
      incorrect <- rbind(incorrect, row_new)
    } 
    
    # 获取 correct 的行数
    num_rows <- nrow(correct)
    
    # # 判断行数是否大于等于10
    # if (num_rows >= 10) {
    #   # 随机提取10行数据
    #   sample_new <- correct %>% sample_n(10)
    # } else {
    #   # 将所有 correct 赋给 sample_new
    #   sample_new <- correct
    # }
    sample_new <- filtered_rows %>% sample_n(10)
  }
  # 格式化 correct 数据框的所有列为字符型
  sample_na[] <- lapply(sample_na, as.character)
  filtered_rows[] <- lapply(filtered_rows, as.character)
  correct[] <- lapply(correct, as.character)
  incorrect[] <- lapply(incorrect, as.character)
  sample_new[] <- lapply(sample_new, as.character)
  
  # 将数据写入工作表
  writeData(wb1, sheet = "sample", sample_na, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "completed", filtered_rows, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "correct", correct, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "incorrect", incorrect, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "manual", sample_new, startCol = 1, startRow = 1)
  # 保存工作簿为Excel文件
  saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
}
  
# 需要处理的年份列表
years <- c("2011","2012","2014") 
# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_1, "address_data_",year, ".csv")
  output_file1 <- paste0(output_folder, "inspect_",year,".xlsx")
  
  all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  all_data_df <- as.data.frame(all_data)
  all_data <- all_data_df
  
  # 创建一个Excel工作簿
  wb1 <- createWorkbook()
  
  # 创建一个新的工作表，并命名为当前年
  addWorksheet(wb1, "sample")
  addWorksheet(wb1, "completed")
  addWorksheet(wb1, "correct")
  addWorksheet(wb1, "incorrect")
  addWorksheet(wb1, "manual")
  
  # 创建两个空的数据框用于存放结果
  incorrect <- data.frame(matrix(ncol = 11, nrow = 0))
  colnames(incorrect) <- c("province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  correct <- data.frame(matrix(ncol = 11, nrow = 0))
  colnames(correct) <- c("province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码", "sdxm_reference", "yzbm_reference", "行政区代码_reference")
  
  # 随机提取50行数据
  sample <- all_data %>% sample_n(50)
  
  # 将所有空值替换为 "NA"
  sample_na <- sample %>% replace(is.na(.), "NA")
  
  # 选择特定的8个列
  selected_cols <- sample_na %>%
    select(qymc,frdbxm, province, 方位, city, town, county, sdxm, yzbm, 行政区代码)
  
  # 挑出这8个列中有7个不是 "NA" 且不等于 "wrong province" 等的行
  filtered_rows <- selected_cols %>%
    filter(rowSums(. != "NA" & . != "wrong province"& . != "wrong city" 
                   & . != "wrong sdxm"& . != "wrong yzbm"
                   & . != "wrong 行政区代码",na.rm = TRUE) >= 7)
  
  # 计算符合条件的行数占总行数的比例
  percentage <- nrow(filtered_rows) / nrow(sample_na) * 100
  
  # 输出结果
  print(paste("相对完整的数据行数的占比为:", percentage, "%"))
  
  
  # 在这里假设有两个表格，你需要根据实际情况修改表格名称和列名
  sheet1 <- read.xlsx(excel_file, sheet = "省地县代码")
  sheet2 <- read.xlsx(excel_file, sheet = "全国")
  
  # 逐行提取sdxm、yzbm和行政区代码，并命名
  for (i in 1:nrow(filtered_rows)) {
    row <- filtered_rows[i, ]
    # 创建一个仅包含所需列的新行 row_new
    row_new <- row[, c("qymc","province", "方位", "city", "town", "county", "sdxm", "yzbm", "行政区代码")]
    
    zb <- 0
    information <- c()
    
    # 初始化一个空字符向量来存储有效值
    valid_information <- character()
    
    for (col_name in c("province", "city", "town", "county")) {
      col_values <- row[[col_name]]
      # 选择除了 "NA", "wrong province", "wrong city" 之外的值
      valid_values <- col_values[!(col_values %in% c("NA", "wrong province", "wrong city"))]
      # 将这些值添加到 valid_information 中
      valid_information <- c(valid_information, valid_values)
    }
    
    # 将 valid_information 合并成一个长字符串，用 "|" 分隔
    information <- paste(valid_information, collapse = "|")
    
    sdxm_individual <- row$sdxm
    yzbm_individual <- row$yzbm
    行政区代码_individual <- row$行政区代码
    
    # 判断sdxm_individual是否不是字符串"NA"
    if (sdxm_individual != "NA"&&sdxm_individual != "wrong sdxm") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(substr(sheet1$代码, 1, 6) == substr(sdxm_individual, 1, 6))
      #matching_rows <- which(substr(sheet1$代码, 1, 6) == paste(substr(sdxm_individual, 1, 4), "00", sep = ""))
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$sdxm_reference<-"NA"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$sdxm_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(paste(reference, collapse = "|"), information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          zb <- 1
        }
      }
    } else {
      row_new$sdxm_reference<-"NA"
      print("sdxm_individual 的值是 NA 或者 wrong sdxm")
    }
    
    # 判断yzbm_individual是否不是字符串"NA"
    if (yzbm_individual != "NA"&&yzbm_individual != "wrong yzbm") {
      zbb <- 1
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(substr(sheet2$D, 1, 4) == substr(yzbm_individual, 1, 4) & nchar(sheet2$D) == nchar(yzbm_individual))
      matching_rows <- which(substr(sheet2$D, 1, 3) == substr(yzbm_individual, 1, 3) & nchar(sheet2$D) == nchar(yzbm_individual))
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$yzbm_reference<-"NA"
      } else {
        for(j in matching_rows){
          reference <- sheet2$B[j]
          
          # 判断reference是否是information中某个值的子字符串
          is_duplicate <- any(grepl(paste(reference, collapse = "|"), information))
          is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
    
          if (is_duplicate) {
            zbb <- 0
            row_new$yzbm_reference<- reference
          }
        }
        if(zbb==1){
          zb <- 1
          row_new$yzbm_reference<-"NA"
        }
      }
    } else {
      row_new$yzbm_reference<-"NA"
      print("yzbm_individual 的值是 NA 或者 wrong yzbm")
    }
    
    # 判断行政区代码_individual是否不是字符串"NA"
    if (行政区代码_individual != "NA"&&行政区代码_individual != "wrong 行政区代码") {
      # 在sheet1的A列中搜索和sdxm_individual一样的行数
      matching_rows <- which(sheet1$代码 == 行政区代码_individual)
      
      # 如果没有匹配到任何行，将zb设为1，否则获取匹配的行内容的B列值
      if (length(matching_rows) == 0) {
        zb <- 1
        row_new$行政区代码_reference<-"NA"
      } else {
        reference <- sheet1$省市[matching_rows]
        row_new$行政区代码_reference <- reference
        # 判断reference是否是information中某个值的子字符串
        is_duplicate <- any(grepl(paste(reference, collapse = "|"), information))
        is_duplicate <- any(grepl(paste(substr(reference, 1, 2), collapse = "|"), information))
        
        if (!is_duplicate) {
          zb <- 1
        }
      }
    } else {
      row_new$行政区代码_reference<-"NA"
      print("行政区代码_individual 的值是 NA 或者 wrong 行政区代码")
    }
    if (row_new$行政区代码_reference == "NA" &&
        row_new$yzbm_reference == "NA" &&
        row_new$sdxm_reference == "NA"){
      zb <- 1
    }
    
    if (zb == 0){
      correct <- rbind(correct, row_new)
    } else{
      incorrect <- rbind(incorrect, row_new)
    } 
    
    # # 获取 correct 的行数
    # num_rows <- nrow(correct)
    # 
    # # 判断行数是否大于等于10
    # if (num_rows >= 10) {
    #   # 随机提取10行数据
    #   sample_new <- correct %>% sample_n(10)
    # } else {
    #   # 将所有 correct 赋给 sample_new
    #   sample_new <- correct
    # }
    
    sample_new <- filtered_rows %>% sample_n(10)
  }
  # 格式化 correct 数据框的所有列为字符型
  sample_na[] <- lapply(sample_na, as.character)
  filtered_rows[] <- lapply(filtered_rows, as.character)
  correct[] <- lapply(correct, as.character)
  incorrect[] <- lapply(incorrect, as.character)
  sample_new[] <- lapply(sample_new, as.character)
  
  # 将数据写入工作表
  writeData(wb1, sheet = "sample", sample_na, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "completed", filtered_rows, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "correct", correct, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "incorrect", incorrect, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "manual", sample_new, startCol = 1, startRow = 1)
  # 保存工作簿为Excel文件
  saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
  print(paste("完成", year, "的工作"))
}
  
  
  