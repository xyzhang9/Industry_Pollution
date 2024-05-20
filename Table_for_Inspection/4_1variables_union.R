library(openxlsx)
library(dplyr)
# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)

# 指定 CSV 文件路径的文件夹
input_folder_path <- ""
output_folder_path <- ""

excel_file <- paste0(input_folder_path, "指标说明.xlsx")

output_file1 <- paste0(output_folder_path, "enviro_yelei_vari.xlsx")

# 读取"变量"表
variable_sheet <- read.xlsx(excel_file, sheet = "变量")

#创建输出2个表格
wb1 <- createWorkbook()

addWorksheet(wb1, "environment")
addWorksheet(wb1, "gongye")

# 建立一个空列表，用于存储每年的中文列名
Zhongwen_zong <- list()
Yelei_zong <- list()

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
 
  # 继续处理数据
  
  # 获取数据框的列名
  column_names <- names(all_data)
  
  # 通过正则表达式筛选出中文列名
  Zhongwen <- column_names[grepl("[\u4e00-\u9fa5]", column_names)]
  
  # 输出中文列名向量
  print(Zhongwen)
  
  # 将中文列名添加到列表中
  Zhongwen_zong[[year]] <- Zhongwen
  
  # 输出非中文列名
  FeiZhongwen <- column_names[!grepl("[\u4e00-\u9fa5]", column_names)]
  print(FeiZhongwen)
  
  # 处理每个Variable_name_original值
  for (i in 1:length(FeiZhongwen)) {
    # 获取当前值
    current_value <- FeiZhongwen[i] 
    
    # 在"变量"表的"A"列中查找对应的行数
    row_index <- which(variable_sheet[, 1] == current_value)
    
    # 如果找到对应的行
    if (length(row_index) > 0) {
      # 获取中文名，并存储到na_summary中
      row_index <- row_index[1]
      FeiZhongwen[i] <- variable_sheet[, 3][row_index]
    } else {
      # 如果找不到，将保留原始的变量名
      FeiZhongwen[i] <- current_value
    }
  }
  
  # 寻找同时包含"类"和"业"的元素，并存储到 Yelei[year] 中
  Yelei <- FeiZhongwen[grepl("类", FeiZhongwen)]
  Yelei_zong[[year]] <- Yelei
  
  
  print(paste("完成", year, "的工作"))
}
# 建立一个 Zhongwen_zong 来取17年的并集
Zhongwen_zong <- unique(unlist(Zhongwen_zong))
Yelei_zong <- unique(unlist(Yelei_zong))

writeData(wb1, sheet = "environment", Zhongwen_zong , startCol = 1, startRow = 1)
writeData(wb1, sheet = "gongye", Yelei_zong , startCol = 1, startRow = 1)

# 保存工作簿为Excel文件
saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
