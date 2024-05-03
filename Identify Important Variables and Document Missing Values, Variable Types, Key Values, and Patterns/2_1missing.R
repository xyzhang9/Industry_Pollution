library(openxlsx)

# 设置文件夹路径和文件名
input_folder_1 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算1/new/"
input_folder_2 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算2/"

excel_file <- paste0(input_folder_2, "指标说明.xlsx")
output_file <- paste0(output_folder, "data_organization_total_new.xlsx")


# 获取所有csv文件的文件路径
csv_files <- list.files(path = input_folder_1, pattern = "\\.csv$", full.names = TRUE)

# 读取所有csv文件并合并
all_data <- do.call(rbind, lapply(csv_files, function(file) {
  read.csv(file, fileEncoding = "UTF-8", stringsAsFactors = FALSE)
}))
all_data_df <- as.data.frame(all_data)
all_data <- all_data_df

#str(all_data)

# 计算每列中 "NA" 的百分比
na_percentages <- colMeans(is.na(all_data)) * 100

# 定义一个函数，用于获取每列出现次数最多的5个值作为 Key_value
get_key_values <- function(x) {
  key_values <- lapply(x, function(col) {
    # 去除缺失值
    non_missing_values <- col[!is.na(col)]
    
    # 统计每个值出现的次数
    value_counts <- table(non_missing_values)
    
    # 按照出现次数从大到小排序
    sorted_values <- sort(value_counts, decreasing = TRUE)
    
    # 获取前5个非缺失值，以值（出现的次数）的形式保存
    key_vals <- names(sorted_values)[1:5]
    key_vals_with_counts <- paste0(key_vals, " (", sorted_values[key_vals], ")")
    
    # 如果非缺失值的种类少于5种，补充为NA
    if (length(sorted_values) == 0) {
      key_vals_with_counts <- NA
    } else if (length(sorted_values) < 5) {
      k <- length(sorted_values)
      key_vals_with_counts <- c(key_vals_with_counts[1:k], rep(NA, 5 - length(key_vals)))
    }
    
    return(key_vals_with_counts)  # 返回前5个值（值+出现的次数）
  })
  
  return(key_values)
}


# 使用函数获取每列出现次数最多的5个非缺失值作为 Key_value
key_values <- get_key_values(all_data)

# 读取"变量"表
variable_sheet <- read.xlsx(excel_file, sheet = "变量")

# 创建一个包含列名和对应"NA"百分比的数据框
na_summary <- data.frame(
  `Variable_name_original` = names(na_percentages),
  `Variable_name_Chinese` = rep(NA, length(na_percentages)),  # 这里用原始列名作为中文列名，您可以根据实际情况修改
  `Percentage_of_missing_value` = sprintf("%.2f%%", na_percentages),
  `Variable_type` = sapply(all_data, function(x) if(is.character(x)) "Character" else "Numeric"),
  `Key_value_1` = rep(NA, length(na_percentages)),
  `Key_value_2` = rep(NA, length(na_percentages)),
  `Key_value_3` = rep(NA, length(na_percentages)),
  `Key_value_4` = rep(NA, length(na_percentages)),
  `Key_value_5` = rep(NA, length(na_percentages)),
  stringsAsFactors = FALSE
)

# 将获取的Key_value填入na_summary
for (i in seq_along(na_percentages)) {
  na_summary$Key_value_1[i] <- key_values[[i]][1]
  na_summary$Key_value_2[i] <- key_values[[i]][2]
  na_summary$Key_value_3[i] <- key_values[[i]][3]
  na_summary$Key_value_4[i] <- key_values[[i]][4]
  na_summary$Key_value_5[i] <- key_values[[i]][5]
}

# 处理每个Variable_name_original值
for (i in seq_along(na_summary$Variable_name_original)) {
  # 获取当前值
  current_value <- na_summary$Variable_name_original[i]
  
  # 在"变量"表的"A"列中查找对应的行数
  row_index <- which(variable_sheet[, 1] == current_value)
  
  # 如果找到对应的行
  if (length(row_index) > 0) {
    # 获取中文名，并存储到na_summary中
    na_summary$Variable_name_Chinese[i] <- variable_sheet[, 3][row_index]
  }else {
    # 如果找不到，将保留原始的变量名
    na_summary$Variable_name_Chinese[i] <- current_value
  }
}

# 创建一个Excel工作簿
wb <- createWorkbook()

# 在工作簿中添加一个工作表
addWorksheet(wb, "Total")

# 将数据框写入工作表中
writeData(wb, sheet = "Total", na_summary, startCol = 1, startRow = 1)

# 保存工作簿为Excel文件
saveWorkbook(wb, file = output_file, overwrite = TRUE)

# 输出成功信息
cat("数据已写入到 data_organization.xlsx 文件中\n")