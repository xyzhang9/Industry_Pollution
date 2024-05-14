library(openxlsx)
library(dplyr)
# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)

# 指定 CSV 文件路径的文件夹
input_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算4/"

excel_file <- paste0(input_folder_path, "指标说明.xlsx")
# 读取"变量"表
variable_sheet <- read.xlsx(excel_file, sheet = "变量")

output_file1 <- paste0(output_folder_path, "industry_cate.xlsx")

#创建输出1个表格
wb1 <- createWorkbook()

# 在工作簿中添加一个工作表
addWorksheet(wb1,"all")

# 挑选需要的列
selected_cols <- c(
  "hylb",	"行业类别代码",	"hymlmc",	"行业类别名称",	"hydldm",	"hydlmc",
  "hyzldm",	"hyzlmc",	"hyxldm",	"hyxlmc"
)
# 个10

# 创建一个空的数据框用于存储所有年份的数据
indus_zong <- data.frame()

# 创建一个空的数据框
present_year <- data.frame(matrix(ncol = length(selected_cols), nrow = 1))
processed_years <- data.frame(matrix(ncol = length(selected_cols), nrow = 1))
# 设置列名为 selected_cols
colnames(present_year) <- selected_cols
colnames(processed_years) <- selected_cols

# 创建空向量
selected_cols_zw <- c()


# 处理每个Variable_name_original值
for (i in 1:length(selected_cols)) {
  # 获取当前值
  current_value <- selected_cols[i] 
  
  # 在"变量"表的"A"列中查找对应的行数
  row_index <- which(variable_sheet[, 1] == current_value)
  
  # 如果找到对应的行
  if (length(row_index) > 0) {
    # 获取中文名，并存储到na_summary中
    row_index <- row_index[1]
    selected_cols_zw[i] <- variable_sheet[, 3][row_index]
  } else {
    # 如果找不到，将保留原始的变量名
    selected_cols_zw[i] <- current_value
  }
}


# 需要处理的年份列表
years <- c("1998","1999","2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010","2011","2012","2013","2014") 

# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_path, year, "_new.csv")
  
  
  # 读取原始的 1998.csv 文件，指定文件编码为 UTF-8
  data <- read.csv(input_file_path, fileEncoding = "UTF-8",na.strings = c("", "NA"))
  all_data_df <- as.data.frame(data)
  all_data <- all_data_df
  
  # 将所有空值替换为 "NA"
  all_data[is.na(all_data)] <- "NA"
  all_data[all_data == ""] <- "NA"
  
  # 获取与selected_cols的交集
  intersect_cols <- intersect(names(all_data), selected_cols)
  
  # 记录出现的年份
  # 计算交集列中每列值为"NA"的百分比
  for (col in intersect_cols) {
    if (is.na(present_year[col])) {
      present_year[col] <- year
    } else {
      present_year[col] <- paste(present_year[col], year, sep = ",")
    }
  }
  
  # 继续处理数据
  selected_data <- data.frame(matrix(ncol = length(selected_cols) , nrow = nrow(data)))
  # 设置列名为 selected_cols
  colnames(selected_data) <- selected_cols
  
  # 将原始数据中的选定列复制到新的数据框中，并将非标准的 "NA" 值替换为 NA
  for (col in selected_cols) {
    if (col %in% colnames(all_data)) {
      # 检查每个单元格，将空格替换为 NA
      selected_data[[col]] <- all_data[, col]
    } else {
      selected_data[[col]] <- "NA"
    }
  }

  # 将当前年份的数据添加到indus_zong中
  indus_zong <- rbind(indus_zong,  selected_data)
  
  # 输出成功信息
  print(paste("完成", year, "的工作"))
}
# 处理年份

for (col in selected_cols) {
  present_year_in <- present_year[col]
  
  jl_index <- 0
  lx_index <- 0
  jl <- c()
  lx <- c()
  # 提取 present_year_in 列，并转换为字符向量
  present_year_in <- as.character(present_year_in)
  # 使用逗号将字符串拆分成单独的年份
  present_year_in <- unlist(strsplit(present_year_in, ","))
  if(length(present_year_in)>1){
    start_year <- present_year_in[1]
    fin_year <- present_year_in[1]
    # 遍历 present_year_in 中的年份
    for (i in 2:length(present_year_in)) {
      
      # 初始化一个变量，用于记录上一个处理的年份
      previous_year <- as.numeric(present_year_in[i-1])
      
      # 获取当前年份的数值表示
      current_year <- as.numeric(present_year_in[i])
      
      if(current_year-previous_year == 1){
        if(start_year == previous_year&&jl_index == 1){
          # 使用逗号将字符串拆分成单独的年份
          jl <- unlist(strsplit(jl, ","))
          jl <- paste(jl[-length(jl)], collapse = ",")
        }
        fin_year <- current_year
        lx <- paste0(start_year,"~",fin_year)
        lx_index <- 1
      }else{
        start_year <- current_year
        if(fin_year == present_year_in[1]&&jl_index==0){
          jl <- fin_year
          jl <- append(jl,current_year)
        }else {
          if(lx_index == 1){
            if(length(jl)==0){
              jl <- lx
            }else{
              jl <- append(jl,lx)
            }
            jl <- append(jl,current_year)
          }else{
            jl <- append(jl,current_year)
          }
        }
        jl <- paste(jl, collapse = ",")
        start_year <- current_year
        lx_index <- 0
        jl_index <- 1
      }
    }
    if(jl_index == 0){
      processed_years[col] <- lx
    }else{
      if(lx_index == 1) {
        jl <- append(jl,lx)
        jl <- paste(jl, collapse = ",")
        processed_years[col] <- jl
      }else{
        processed_years[col] <- jl
      }
    }
  }else{
    processed_years[col] <- present_year_in
  }
  
  # 输出处理后的结果
  processed_years[col] <-  processed_years[col]
}

# 定义一个函数，用于获取每列出现次数最多的10个值作为 Key_value
get_key_values <- function(x) {
  key_values <- lapply(x, function(col) {
    # 去除值为 "NA" 的元素
    non_missing_values <- col[col != "NA"]
    
    # 统计每个值出现的次数
    value_counts <- table(non_missing_values)
    
    # 按照出现次数从大到小排序
    sorted_values <- sort(value_counts, decreasing = TRUE)
    
    # 获取前10个非缺失值，以值（出现的次数）的形式保存
    key_vals <- names(sorted_values)[1:10]
    key_vals_with_counts <- paste0(key_vals, " (", sorted_values[key_vals], ")")
    
    # 如果非缺失值的种类少于10种，补充为NA
    if (length(sorted_values) == 0) {
      key_vals_with_counts <- NA
    } else if (length(sorted_values) < 10) {
      k <- length(sorted_values)
      key_vals_with_counts <- c(key_vals_with_counts[1:k], rep(NA, 10 - length(key_vals)))
    }
    
    return(key_vals_with_counts)  # 返回前10个值（值+出现的次数）
  })
  
  return(key_values)
}

# 使用函数获取每列出现次数最多的10个非缺失值作为 Key_value
key_values <- get_key_values(indus_zong)

# 创建一个包含列名和对应"NA"百分比的数据框
summary <- data.frame(
  `Variable_name_original` = names(indus_zong),
  `Variable_name_Chinese` = selected_cols_zw,  
  `Year_of_presenting` = rep(NA, length(indus_zong)),
  `Key_value_1` = rep(NA, length(indus_zong)),
  `Key_value_2` = rep(NA, length(indus_zong)),
  `Key_value_3` = rep(NA, length(indus_zong)),
  `Key_value_4` = rep(NA, length(indus_zong)),
  `Key_value_5` = rep(NA, length(indus_zong)),
  `Key_value_6` = rep(NA, length(indus_zong)),
  `Key_value_7` = rep(NA, length(indus_zong)),
  `Key_value_8` = rep(NA, length(indus_zong)),
  `Key_value_9` = rep(NA, length(indus_zong)),
  `Key_value_10` = rep(NA, length(indus_zong)),
  stringsAsFactors = FALSE
)

# 将获取的Key_value填入na_summary
for (i in seq_along(selected_cols)) {
  summary$Year_of_presenting[i] <- processed_years[[i]]
  summary$Key_value_1[i] <- key_values[[i]][1]
  summary$Key_value_2[i] <- key_values[[i]][2]
  summary$Key_value_3[i] <- key_values[[i]][3]
  summary$Key_value_4[i] <- key_values[[i]][4]
  summary$Key_value_5[i] <- key_values[[i]][5]
  summary$Key_value_6[i] <- key_values[[i]][6]
  summary$Key_value_7[i] <- key_values[[i]][7]
  summary$Key_value_8[i] <- key_values[[i]][8]
  summary$Key_value_9[i] <- key_values[[i]][9]
  summary$Key_value_10[i] <- key_values[[i]][10]
}


# 将数据框写入工作表中
writeData(wb1, sheet = "all", summary, startCol = 1, startRow = 1)

# 保存工作簿为Excel文件
saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
cat("数据整理完成并保存到 industry_cate.xlsx 文件中\n")
