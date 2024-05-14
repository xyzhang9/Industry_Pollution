library(openxlsx)
library(dplyr)
# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)

# 指定 CSV 文件路径的文件夹
input_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算4/"

output_file1 <- paste0(output_folder_path, "enviromtal_vari.xlsx")

#创建输出1个表格
wb1 <- createWorkbook()

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
  # 挑选需要的列
  selected_cols <- c(
    "工业用水总量吨",	"煤炭消费总量吨",	"其中新鲜水量吨",	"重复用水量吨",
    "其中燃料煤消费量吨",	"原料煤消费量吨",	"燃料煤平均硫份",	"燃料油消费量不含车船用吨",
    "其中重油吨",	"柴油吨",	"重油平均硫份",	"洁净燃气消费量万立方米",	"废水治理设施数套",
    "废水治理设施处理能力吨日",	"工业废水处理量吨",	"工业废水排放量吨",	"氨氮去除量千克",
    "化学需氧量去除量千克",	"其中当年新增设施去除的千克",	"化学需氧量排放量千克",	
    "氨氮排放量千克",	"工业废气排放总量万标立方米",	"废气治理设施数套",	"其中脱硫设施数套",
    "废气治理设施处理能力标立方米时",	"其中脱硫设施脱硫能力千克时",	"二氧化硫去除量千克",
    "二氧化硫排放量千克",	"氮氧化物去除量千克",	"氮氧化物排放量千克",	"烟尘去除量千克",
    "烟尘排放量千克",	"工业粉尘去除量千克",	"工业粉尘排放量千克","其中新鲜用水量吨",
    "燃料油评价硫份",	"洁净燃气消费万立方米",	"废水治理设施处理能力日吨",	"其中当年新增实施去除量千克",
    "工业用水量吨",	"其中取水量吨",	"其中重复用水量吨",	"废水治理设施运行费用万元",
    "化学需氧量产生量吨",	"化学需氧量排放量吨",	"氨氮产生量吨",	"氨氮排放量吨",
    "二氧化硫产生量吨",	"二氧化硫排放量吨",	"氮氧化物产生量吨",	"氮氧化物排放量吨",
    "烟粉尘排放量吨",	"烟粉尘产生量吨","煤炭消耗量吨",	"其中燃料煤消耗量吨",	"其中燃料煤平均含硫量",
    "燃料油消耗量不含车船用吨",	"天然气消耗量万立方米",	"工业废水化学需氧量产生量吨",
    "工业废水化学需氧量排放量吨",	"工业废水氨氮产生量吨",	"工业废水氨氮排放量吨",
    "工业废气排放量万立方米",	"废气治理设施处理能力立方米时",	"其中脱硫设施脱硫能力立方米时",
    "工业废气二氧化硫产生量吨",	"工业废气二氧化硫排放量吨",	"工业废气氮氧化物产生量吨",
    "工业废气氮氧化物排放量吨",	"工业废气烟粉尘排放量吨",	"工业废气烟粉尘产生量吨"
  )
  # 个71
  
  # 创建一个长度与selected_cols相同的空向量用于存储na_percentages
  na_percentages <- rep(NA, length(selected_cols))
  
  # 获取与selected_cols的交集
  intersect_cols <- intersect(names(all_data), selected_cols)
  
  # 计算交集列中每列值为"NA"的百分比
  for (col in intersect_cols) {
    na_percentages[col] <- paste(round(mean(all_data[[col]] == "NA") * 100, 2), "%", sep = "")
  }
  
  # 获取selected_cols中存在但在all_data中不存在的列名
  nonexist_cols <- selected_cols[!selected_cols %in% intersect_cols]
  
  # 将不存在的列的na_percentages设为"nonexist"
  for (col in nonexist_cols) {
  na_percentages[col] <- "nonexist"
  }
  
  paixu <- c()  # 初始化空向量
  
  for (col in selected_cols){
    paixu <- append(paixu, na_percentages[col])
  }
  
  # 创建na_summary数据框
  na_summary <- data.frame(Variables = selected_cols, Attribute = paixu)
  
  # 在工作簿中添加一个工作表
  addWorksheet(wb1, year)
  
  # 将数据框写入工作表中
  writeData(wb1, sheet = year, na_summary, startCol = 1, startRow = 1)
  
  # 输出成功信息
  print(paste("完成", year, "的工作"))
}
# 保存工作簿为Excel文件
saveWorkbook(wb1, file = output_file1, overwrite = TRUE)
cat("数据整理完成并保存到 enviromtal_vari.xlsx 文件中\n")
