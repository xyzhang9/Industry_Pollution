library(openxlsx)

# 设置文件夹路径和文件名
input_folder_1 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算1/"
input_folder_2 <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算3/"

excel_file <- paste0(input_folder_2, "指标说明.xlsx")

# 需要处理的年份列表
years <- c("1998","1999","2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010","2011","2012","2013") 

# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_1, year, ".csv")
  output_file <- paste0(output_folder, "ann-merge_", year, ".xlsx")
  
  # 创建一个Excel工作簿
  wb <- createWorkbook()
  
  all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  all_data_df <- as.data.frame(all_data)
  all_data <- all_data_df
  
  # 提取重复的 id
  duplicated_ids <- all_data$id[duplicated(all_data$id)]
  
  # 遍历每个重复的 id
  for (id in unique(duplicated_ids)) {
    # 获取当前重复 id 对应的所有行数据
    duplicate_rows <- all_data[all_data$id == id, ]
    
    # 创建一个新的工作表，并命名为当前重复的 id
    addWorksheet(wb, id)
    
    # 将数据写入工作表
    writeData(wb, sheet = id, duplicate_rows, startCol = 1, startRow = 1)
  }
  # 保存工作簿为Excel文件
  saveWorkbook(wb, file = output_file, overwrite = TRUE)
  
  # 输出成功信息
  cat("数据已写入到 ann-merge.xlsx 文件中\n")
}
#2014年单独来，因为它没有id列，采用企业名称
input_file_path <- paste0(input_folder_1, "2014", ".csv")
output_file <- paste0(output_folder, "ann-merge_", "2014", ".xlsx")

# 创建一个Excel工作簿
wb <- createWorkbook()

all_data <- read.csv(input_file_path, fileEncoding = "UTF-8")
all_data_df <- as.data.frame(all_data)
all_data <- all_data_df

# 提取重复的 id
duplicated_ids <- all_data$qymc[duplicated(all_data$qymc)]

# 遍历每个重复的 id
for (id in unique(duplicated_ids)) {
  # 获取当前重复 id 对应的所有行数据
  duplicate_rows <- all_data[all_data$qymc == id, ]
  
  # 创建一个新的工作表，并命名为当前重复的 id
  addWorksheet(wb, id)
  
  # 将数据写入工作表
  writeData(wb, sheet = id, duplicate_rows, startCol = 1, startRow = 1)
}
# 保存工作簿为Excel文件
saveWorkbook(wb, file = output_file, overwrite = TRUE)

# 输出成功信息
cat("数据已写入到 ann-merge.xlsx 文件中\n")
  
cat("数据整理完成并保存到 ann-merge.xlsx 文件中\n")