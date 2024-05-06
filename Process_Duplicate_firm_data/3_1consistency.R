library(openxlsx)
library(dplyr)
# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)

# 设置文件夹路径和文件名
input_folder_1 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算1/"
input_folder_2 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算3/地址/"

excel_file <- paste0(input_folder_2, "指标说明.xlsx")

# 需要处理的年份列表
years <- c("1998","1999","2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010") 
# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_1, year, ".csv")
  output_file1 <- paste0(output_folder, "address_data_",year,".xlsx")
  output_file2 <- paste0(output_folder, "wrong_address_", year, ".xlsx")
  
  all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  all_data_df <- as.data.frame(all_data)
  all_data <- all_data_df
  
  # 创建一个Excel工作簿
  wb1 <- createWorkbook()
  wb2 <- createWorkbook()
  
  # 创建一个新的工作表，并命名为当前年
  addWorksheet(wb1, "total")
  addWorksheet(wb1, "non_duplicate")
  addWorksheet(wb1, "duplicate")
  
  # 创建一个空的数据框用于存储所有处理过的行
  address_data <- data.frame(matrix(ncol = 20, nrow = 0))
  colnames(address_data) <- colnames(all_data)[1:20]
  dudu_data <- data.frame(matrix(ncol = 20, nrow = 0))
  colnames(dudu_data) <- colnames(all_data)[1:20]
  
  # 将所有不重复的行添加到 address_data 中
  non_duplicated_rows <- all_data[!duplicated(all_data$id), 1:20]
  non_duplicated_rows[is.na(non_duplicated_rows)] <- "NA"
  address_data <- rbind(address_data, non_duplicated_rows)
  
  # 提取重复的 id
  duplicated_ids <- all_data$id[duplicated(all_data$id)]
  uq <- unique(duplicated_ids)
  
  # 遍历每个重复的 id
  for (id in unique(duplicated_ids)) {
    addWorksheet(wb2, id)
    wrong_data <- data.frame(matrix(ncol = 20, nrow = 0))
    colnames(wrong_data) <- colnames(all_data)[1:20]
    
    # 获取当前重复 id 对应的地址数据
    duplicate_rows <- all_data[all_data$id == id, 1:20]
    
    # 将缺失值填充为 "NA"
    duplicate_rows[is.na(duplicate_rows)] <- "NA"
    
    # 计算每行中NA值的数量
    na_counts <- apply(duplicate_rows, 1, function(row) sum(row == "NA"))
    
    # 找到NA值最少的行的索引
    min_na_index <- which.min(na_counts)
    
    # 将对应行赋值给 address_row
    address_row <- duplicate_rows[min_na_index, ]
    
    # 检查重复的行
    # 
    if (all(duplicate_rows$province == address_row$province) &&
        all(duplicate_rows$city == address_row$city) &&
        all(duplicate_rows$sdxm == address_row$sdxm) &&
        all(duplicate_rows$yzbm == address_row$yzbm)&&
        all(duplicate_rows$行政区代码 == address_row$行政区代码)) {
      # 如果所有重复的行都相同，则只保留第一行
      address_data <- rbind(address_data, address_row)
      dudu_data <- rbind(dudu_data, address_row)
    } else {
        # 忽略空值行
      # 找到包含 "NA" 值的行的索引
      rows_with_na <- which(rowSums(duplicate_rows[, c("province", "city", "sdxm", "yzbm", "行政区代码")] == "NA") > 0)
      
      # 如果任何一列包含 "NA" 值，则将其索引存储在 rows_with_na 中
      if (length(rows_with_na) > 0) {
        # 如果有 NA 值的行，则进行删除操作
        duplicate_rows <- duplicate_rows[-rows_with_na, ]
      }
      
      # 检查哪些列存在不同的值
          # 创建一个空向量来存储不同的列名
          diff_cols_names <- c()
          
          # 检查哪些列存在不同的值
          if (!all(duplicate_rows$province == address_row$province)) {
            diff_cols_names <- c(diff_cols_names, "province")
          }
          if (!all(duplicate_rows$city == address_row$city)) {
            diff_cols_names <- c(diff_cols_names, "city")
          }
          if (!all(duplicate_rows$sdxm == address_row$sdxm)) {
            diff_cols_names <- c(diff_cols_names, "sdxm")
          }
          if (!all(duplicate_rows$yzbm == address_row$yzbm)) {
            diff_cols_names <- c(diff_cols_names, "yzbm")
          }
          if (!all(duplicate_rows$行政区代码 == address_row$行政区代码)) {
            diff_cols_names <- c(diff_cols_names, "行政区代码")
          }
          # 现在 diff_cols_names 中存储的就是不同的列名
          
      # 处理不同的列
      for (col_name in diff_cols_names) {
        
        if (col_name == "province" || col_name == "city") {
          # 将错误数据添加到 wrong_data
          wrong_data <- rbind(wrong_data, duplicate_rows)
          
          # 将对应的 "province" 或者 "city" 列替换为 "wrong province or city"
          address_row[col_name] <- paste0("wrong ", col_name)
        } else if (col_name == "sdxm") {
          # 检查前四位是否一样
          first_sdxm <- substr(duplicate_rows$sdxm[1], 1, 4)
          diff_sdxm <- which(substr(duplicate_rows$sdxm, 1, 4) != first_sdxm)
          
          if (length(diff_sdxm) == 0) {
            # 如果前四位都一样，则保留address_row
            address_row$sdxm <- address_row$sdxm
          } else {
            if (nrow(wrong_data) == 0 ) {
              # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
              wrong_data <- rbind(wrong_data, duplicate_rows)
            }
            # 将对应的 "sdxm" 列替换为 "wrong sdxm"，其他列不变
            address_row$sdxm <- "wrong sdxm"
          }
        } else if (col_name == "yzbm") {
          # 检查前四位是否一样
          first_yzbm <- substr(duplicate_rows$yzbm[1], 1, 4)
          diff_yzbm <- which(substr(duplicate_rows$yzbm, 1, 4) != first_yzbm)
          
          if (length(diff_yzbm) == 0) {
            # 如果前四位都一样，则保留address_row中的yzbm
            address_row$yzbm <- address_row$yzbm
          } else {
            if (nrow(wrong_data) == 0) {
              # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
              wrong_data <- rbind(wrong_data, duplicate_rows)
            }
            # 将对应的 "yzbm" 列替换为 "wrong yzbm"，其他列不变
            address_row$yzbm <- "wrong yzbm"
          }
        }
        else if (col_name == "行政区代码"){
          # 检查前四位是否一样
          first_xzqdm <- substr(duplicate_rows$行政区代码[1], 1, 4)
          diff_xzqdm <- which(substr(duplicate_rows$行政区代码, 1, 4) != first_xzqdm)

          if (length(diff_xzqdm) == 0) {
            # # 进一步检查是否 "sdxm" 和 "行政区代码" 的前四位数字都相同
            # same_sdxm_xzqdm <- which(substr(duplicate_rows$sdxm, 1, 4) == substr(duplicate_rows$行政区代码, 1, 4))
            # 
            # if (length(same_sdxm_xzqdm) == nrow(duplicate_rows)) {
            #   # 如果所有行的 "sdxm" 和 "行政区代码" 的前四位都相同，则保留行政区代码加入 address_data
              address_row$行政区代码 <- address_row$行政区代码
            }else {
            if (nrow(wrong_data) == 0) {
              # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
              wrong_data <- rbind(wrong_data, duplicate_rows)
            }
            # 将对应的 "行政区代码" 列替换为 "wrong 行政区代码"，其他列不变
            address_row$行政区代码 <- "wrong 行政区代码"
          }
         }
       }
      address_data <- rbind(address_data, address_row)
      dudu_data <- rbind(dudu_data, address_row)
      }
    # 将数据写入工作表
    writeData(wb2, sheet = id, wrong_data, startCol = 1, startRow = 1)
  }
  # 输出 address_data 到 CSV 文件
  output_csv_file <- paste0(output_folder, "address_data_",year,".csv")
  write.csv(address_data, file = output_csv_file, row.names = FALSE, fileEncoding = "UTF-8")

  
  # 将数据写入工作表
  writeData(wb1, sheet = "total", address_data, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "non_duplicate", non_duplicated_rows, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "duplicate", dudu_data, startCol = 1, startRow = 1)
  # 保存工作簿为Excel文件
  saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
  # 保存工作簿为Excel文件
  saveWorkbook(wb2, file = output_file2, overwrite = TRUE)
  
  # 输出成功信息
  cat("数据已写入到 Excel 文件和 CSV 文件中\n")
}
#11到13年是行政区划代码
# 需要处理的年份列表
years <- c("2011","2012","2013") 
# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_1, year, ".csv")
  output_file1 <- paste0(output_folder, "address_data_",year,".xlsx")
  output_file2 <- paste0(output_folder, "wrong_address_", year, ".xlsx")
  
  all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  all_data_df <- as.data.frame(all_data)
  all_data <- all_data_df
  
  # 创建一个Excel工作簿
  wb1 <- createWorkbook()
  wb2 <- createWorkbook()
  
  # 创建一个新的工作表，并命名为当前年
  addWorksheet(wb1, "total")
  addWorksheet(wb1, "non_duplicate")
  addWorksheet(wb1, "duplicate")
  
  # 创建一个空的数据框用于存储所有处理过的行
  address_data <- data.frame(matrix(ncol = 20, nrow = 0))
  colnames(address_data) <- colnames(all_data)[1:20]
  dudu_data <- data.frame(matrix(ncol = 20, nrow = 0))
  colnames(dudu_data) <- colnames(all_data)[1:20]
  
  # 将所有不重复的行添加到 address_data 中
  non_duplicated_rows <- all_data[!duplicated(all_data$id), 1:20]
  non_duplicated_rows[is.na(non_duplicated_rows)] <- "NA"
  address_data <- rbind(address_data, non_duplicated_rows)
  
  # 提取重复的 id
  duplicated_ids <- all_data$id[duplicated(all_data$id)]
  uq <- unique(duplicated_ids)
  
  # 遍历每个重复的 id
  for (id in unique(duplicated_ids)) {
    addWorksheet(wb2, id)
    wrong_data <- data.frame(matrix(ncol = 20, nrow = 0))
    colnames(wrong_data) <- colnames(all_data)[1:20]
    
    # 获取当前重复 id 对应的地址数据
    duplicate_rows <- all_data[all_data$id == id, 1:20]
    
    # 将缺失值填充为 "NA"
    duplicate_rows[is.na(duplicate_rows)] <- "NA"
    
    # 计算每行中NA值的数量
    na_counts <- apply(duplicate_rows, 1, function(row) sum(row == "NA"))
    
    # 找到NA值最少的行的索引
    min_na_index <- which.min(na_counts)
    
    # 将对应行赋值给 address_row
    address_row <- duplicate_rows[min_na_index, ]
    
    # 检查重复的行
    # 
    if (all(duplicate_rows$province == address_row$province) &&
        all(duplicate_rows$city == address_row$city) &&
        all(duplicate_rows$sdxm == address_row$sdxm) &&
        all(duplicate_rows$yzbm == address_row$yzbm)&&
        all(duplicate_rows$行政区划代码 == address_row$行政区划代码)) {
      # 如果所有重复的行都相同，则只保留第一行
      address_data <- rbind(address_data, address_row)
      dudu_data <- rbind(dudu_data, address_row)
    } else {
      # 忽略空值行
      # 找到包含 "NA" 值的行的索引
      rows_with_na <- which(rowSums(duplicate_rows[, c("province", "city", "sdxm", "yzbm", "行政区代码")] == "NA") > 0)
      
      # 如果任何一列包含 "NA" 值，则将其索引存储在 rows_with_na 中
      if (length(rows_with_na) > 0) {
        # 如果有 NA 值的行，则进行删除操作
        duplicate_rows <- duplicate_rows[-rows_with_na, ]
      }
      
      # 检查哪些列存在不同的值
      # 创建一个空向量来存储不同的列名
      diff_cols_names <- c()
      
      # 检查哪些列存在不同的值
      if (!all(duplicate_rows$province == address_row$province)) {
        diff_cols_names <- c(diff_cols_names, "province")
      }
      if (!all(duplicate_rows$city == address_row$city)) {
        diff_cols_names <- c(diff_cols_names, "city")
      }
      if (!all(duplicate_rows$sdxm == address_row$sdxm)) {
        diff_cols_names <- c(diff_cols_names, "sdxm")
      }
      if (!all(duplicate_rows$yzbm == address_row$yzbm)) {
        diff_cols_names <- c(diff_cols_names, "yzbm")
      }
      if (!all(duplicate_rows$行政区划代码 == address_row$行政区划代码)) {
        diff_cols_names <- c(diff_cols_names, "行政区代码")
      }
      # 现在 diff_cols_names 中存储的就是不同的列名
      
      # 处理不同的列
      for (col_name in diff_cols_names) {
        
        if (col_name == "province" || col_name == "city") {
          # 将错误数据添加到 wrong_data
          wrong_data <- rbind(wrong_data, duplicate_rows)
          
          # 将对应的 "province" 或者 "city" 列替换为 "wrong province or city"
          address_row[col_name] <- paste0("wrong ", col_name)
        } else if (col_name == "sdxm") {
          # 检查前四位是否一样
          first_sdxm <- substr(duplicate_rows$sdxm[1], 1, 4)
          diff_sdxm <- which(substr(duplicate_rows$sdxm, 1, 4) != first_sdxm)
          
          if (length(diff_sdxm) == 0) {
            # 如果前四位都一样，则保留address_row
            address_row$sdxm <- address_row$sdxm
          } else {
            if (nrow(wrong_data) == 0 ) {
              # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
              wrong_data <- rbind(wrong_data, duplicate_rows)
            }
            # 将对应的 "sdxm" 列替换为 "wrong sdxm"，其他列不变
            address_row$sdxm <- "wrong sdxm"
          }
        } else if (col_name == "yzbm") {
          # 检查前四位是否一样
          first_yzbm <- substr(duplicate_rows$yzbm[1], 1, 4)
          diff_yzbm <- which(substr(duplicate_rows$yzbm, 1, 4) != first_yzbm)
          
          if (length(diff_yzbm) == 0) {
            # 如果前四位都一样，则保留address_row中的yzbm
            address_row$yzbm <- address_row$yzbm
          } else {
            if (nrow(wrong_data) == 0) {
              # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
              wrong_data <- rbind(wrong_data, duplicate_rows)
            }
            # 将对应的 "yzbm" 列替换为 "wrong yzbm"，其他列不变
            address_row$yzbm <- "wrong yzbm"
          }
        }
        else if (col_name == "行政区划代码"){
          # 检查前四位是否一样
          first_xzqdm <- substr(duplicate_rows$行政区划代码[1], 1, 4)
          diff_xzqdm <- which(substr(duplicate_rows$行政区划代码, 1, 4) != first_xzqdm)
          
          if (length(diff_xzqdm) == 0) {
            # # 进一步检查是否 "sdxm" 和 "行政区代码" 的前四位数字都相同
            # same_sdxm_xzqdm <- which(substr(duplicate_rows$sdxm, 1, 4) == substr(duplicate_rows$行政区代码, 1, 4))
            # 
            # if (length(same_sdxm_xzqdm) == nrow(duplicate_rows)) {
            #   # 如果所有行的 "sdxm" 和 "行政区代码" 的前四位都相同，则保留行政区代码加入 address_data
            address_row$行政区划代码 <- address_row$行政区划代码
          }else {
            if (nrow(wrong_data) == 0) {
              # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
              wrong_data <- rbind(wrong_data, duplicate_rows)
            }
            # 将对应的 "行政区代码" 列替换为 "wrong 行政区代码"，其他列不变
            address_row$行政区划代码 <- "wrong 行政区代码"
          }
        }
      }
      address_data <- rbind(address_data, address_row)
      dudu_data <- rbind(dudu_data, address_row)
    }
    # 将数据写入工作表
    writeData(wb2, sheet = id, wrong_data, startCol = 1, startRow = 1)
  }
  # 输出 address_data 到 CSV 文件
  output_csv_file <- paste0(output_folder, "address_data_",year,".csv")
  write.csv(address_data, file = output_csv_file, row.names = FALSE, fileEncoding = "UTF-8")
  
  
  # 将数据写入工作表
  writeData(wb1, sheet = "total", address_data, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "non_duplicate", non_duplicated_rows, startCol = 1, startRow = 1)
  writeData(wb1, sheet = "duplicate", dudu_data, startCol = 1, startRow = 1)
  # 保存工作簿为Excel文件
  saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
  # 保存工作簿为Excel文件
  saveWorkbook(wb2, file = output_file2, overwrite = TRUE)
  
  # 输出成功信息
  cat("数据已写入到 Excel 文件和 CSV 文件中\n")
}

#2014年单独来，因为它没有id列，采用企业名称
input_file_path <- paste0(input_folder_1, "2014", ".csv")
output_file2 <- paste0(output_folder, "wrong_address_", "2014", ".xlsx")
output_file1 <- paste0(output_folder, "address_data_","2014",".xlsx")
all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
all_data_df <- as.data.frame(all_data)
all_data <- all_data_df

# 创建一个Excel工作簿
wb1 <- createWorkbook()
wb2 <- createWorkbook()

addWorksheet(wb1, "total")
addWorksheet(wb1, "non_duplicate")
addWorksheet(wb1, "duplicate")

# 创建一个空的数据框用于存储所有处理过的行
address_data <- data.frame(matrix(ncol = 20, nrow = 0))
colnames(address_data) <- colnames(all_data)[1:20]
dudu_data <- data.frame(matrix(ncol = 20, nrow = 0))
colnames(dudu_data) <- colnames(all_data)[1:20]

# 将所有不重复的行添加到 address_data 中
non_duplicated_rows <- all_data[!duplicated(all_data$qymc), 1:20]
address_data <- rbind(address_data, non_duplicated_rows)

# 提取重复的 id
duplicated_ids <- all_data$qymc[duplicated(all_data$qymc)]
uq <- unique(duplicated_ids)

# 遍历每个重复的 id
for (id in unique(duplicated_ids)) {
  addWorksheet(wb2, id)
  wrong_data <- data.frame(matrix(ncol = 20, nrow = 0))
  colnames(wrong_data) <- colnames(all_data)[1:20]
  
  # 获取当前重复 id 对应的地址数据
  duplicate_rows <- all_data[all_data$qymc == id, 1:20]
  
  # 将缺失值填充为 "NA"
  duplicate_rows[is.na(duplicate_rows)] <- "NA"
  
  # 计算每行中NA值的数量
  na_counts <- apply(duplicate_rows, 1, function(row) sum(row == "NA"))
  
  # 找到NA值最少的行的索引
  min_na_index <- which.min(na_counts)
  
  # 将对应行赋值给 address_row
  address_row <- duplicate_rows[min_na_index, ]
  # 检查重复的行
  # 
  if (all(duplicate_rows$province == address_row$province) &&
      all(duplicate_rows$city == address_row$city) &&
      all(duplicate_rows$sdxm == address_row$sdxm) &&
      all(duplicate_rows$yzbm == address_row$yzbm) &&
      all(duplicate_rows$行政区划代码 == address_row$行政区划代码)) {
    # 如果所有重复的行都相同，则只保留第一行
    address_data <- rbind(address_data, address_row)
    dudu_data <- rbind(dudu_data, address_row)
  } else {
    # 忽略空值行
    # 找到包含 "NA" 值的行的索引
    rows_with_na <- which(rowSums(duplicate_rows[, c("province", "city", "sdxm", "yzbm", "行政区代码")] == "NA") > 0)
    
    # 如果任何一列包含 "NA" 值，则将其索引存储在 rows_with_na 中
    if (length(rows_with_na) > 0) {
      # 如果有 NA 值的行，则进行删除操作
      duplicate_rows <- duplicate_rows[-rows_with_na, ]
    }
    
    # 检查哪些列存在不同的值
    # &
    #   duplicate_rows$行政区代码 == address_row$行政区代码
      # 创建一个空向量来存储不同的列名
      diff_cols_names <- c()
      
      # 检查哪些列存在不同的值
      if (!all(duplicate_rows$province == address_row$province)) {
        diff_cols_names <- c(diff_cols_names, "province")
      }
      if (!all(duplicate_rows$city == address_row$city)) {
        diff_cols_names <- c(diff_cols_names, "city")
      }
      if (!all(duplicate_rows$sdxm == address_row$sdxm)) {
        diff_cols_names <- c(diff_cols_names, "sdxm")
      }
      if (!all(duplicate_rows$yzbm == address_row$yzbm)) {
        diff_cols_names <- c(diff_cols_names, "yzbm")
      }
      if (!all(duplicate_rows$行政区划代码 == address_row$行政区划代码)) {
        diff_cols_names <- c(diff_cols_names, "行政区代码")
      }
      # 现在 diff_cols_names 中存储的就是不同的列名
      
    # 处理不同的列
    for (col_name in diff_cols_names) {
      
      if (col_name == "province" || col_name == "city") {
        # 将错误数据添加到 wrong_data
        wrong_data <- rbind(wrong_data, duplicate_rows)
        
        # 将对应的 "province" 或者 "city" 列替换为 "wrong province or city"
        address_row[col_name] <- paste0("wrong ", col_name)
      } else if (col_name == "sdxm") {
        # 检查前四位是否一样
        first_sdxm <- substr(duplicate_rows$sdxm[1], 1, 4)
        diff_sdxm <- which(substr(duplicate_rows$sdxm, 1, 4) != first_sdxm)
        
        if (length(diff_sdxm) == 0) {
          # 如果前四位都一样，则保留address_row
          address_row$sdxm <- address_row$sdxm
        } else {
          if (nrow(wrong_data) == 0 ) {
            # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
            wrong_data <- rbind(wrong_data, duplicate_rows)
          }
          # 将对应的 "sdxm" 列替换为 "wrong sdxm"，其他列不变
          address_row$sdxm <- "wrong sdxm"
        }
      } else if (col_name == "yzbm") {
        # 检查前四位是否一样
        first_yzbm <- substr(duplicate_rows$yzbm[1], 1, 4)
        diff_yzbm <- which(substr(duplicate_rows$yzbm, 1, 4) != first_yzbm)
        
        if (length(diff_yzbm) == 0) {
          # 如果前四位都一样，则保留address_row中的yzbm
          address_row$yzbm <- address_row$yzbm
        } else {
          if (nrow(wrong_data) == 0) {
            # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
            wrong_data <- rbind(wrong_data, duplicate_rows)
          }
          # 将对应的 "yzbm" 列替换为 "wrong yzbm"，其他列不变
          address_row$yzbm <- "wrong yzbm"
        }
      }
      else if (col_name == "行政区划代码"){
        # 检查前四位是否一样
        first_xzqdm <- substr(duplicate_rows$行政区划代码[1], 1, 4)
        diff_xzqdm <- which(substr(duplicate_rows$行政区划代码, 1, 4) != first_xzqdm)
        
        if (length(diff_xzqdm) == 0) {
          # # 进一步检查是否 "sdxm" 和 "行政区代码" 的前四位数字都相同
          # same_sdxm_xzqdm <- which(substr(duplicate_rows$sdxm, 1, 4) == substr(duplicate_rows$行政区代码, 1, 4))
          # 
          # if (length(same_sdxm_xzqdm) == nrow(duplicate_rows)) {
          #   # 如果所有行的 "sdxm" 和 "行政区代码" 的前四位都相同，则保留行政区代码加入 address_data
          address_row$行政区划代码 <- address_row$行政区划代码
        }else {
          if (nrow(wrong_data) == 0) {
            # 如果没有包含当前错误数据，则将当前错误数据添加到 wrong_data
            wrong_data <- rbind(wrong_data, duplicate_rows)
          }
          # 将对应的 "行政区代码" 列替换为 "wrong 行政区代码"，其他列不变
          address_row$行政区划代码 <- "wrong 行政区代码"
        }
      }
      
     
    }
    address_data <- rbind(address_data, address_row)
    dudu_data <- rbind(dudu_data, address_row)
    }
  # 将数据写入工作表
  writeData(wb2, sheet = id, wrong_data, startCol = 1, startRow = 1)
}
writeData(wb1, sheet = "total", address_data, startCol = 1, startRow = 1)
writeData(wb1, sheet = "non_duplicate", non_duplicated_rows, startCol = 1, startRow = 1)
writeData(wb1, sheet = "duplicate", dudu_data, startCol = 1, startRow = 1)
# 保存工作簿为Excel文件
saveWorkbook(wb2, file = output_file2, overwrite = TRUE)

# 输出成功信息
cat("数据已写入到 wrong_address_.xlsx 文件中\n")

# 保存工作簿为Excel文件
saveWorkbook(wb1, file = output_file1, overwrite = TRUE)

# 输出 address_data 到 CSV 文件
output_csv_file <- paste0(output_folder, "address_data_2014.csv")
write.csv(address_data, file = output_csv_file, row.names = FALSE, fileEncoding = "UTF-8")

# 输出成功信息
cat("数据已写入到 Excel 文件和 CSV 文件中\n")

