memory.limit(1000000000000)
library(openxlsx)
library(dplyr)
library(tidyr)
# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)

# 指定 CSV 文件路径的文件夹
input_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算4/"

excel_file <- paste0(input_folder_path, "指标说明.xlsx")
# 读取"变量"表
variable_sheet <- read.xlsx(excel_file, sheet = "变量")

output_file1 <- paste0(output_folder_path, "all_ind_env.csv")
output_file2 <- paste0(output_folder_path, "all_ind_env.xlsx")

#创建输出1个表格
wb1 <- createWorkbook()

# 在工作簿中添加一个工作表
addWorksheet(wb1,"all")

# 挑选需要的列
selected_cols <- c(
  "qymc","province", "city", "town", "county", "sdxm","yzbm","行政区代码","行政区划代码",
  "hylb",	"行业类别代码",	"hymlmc",	"行业类别名称",	"hydldm",	"hydlmc",
  "hyzldm",	"hyzlmc",	"hyxldm",	"hyxlmc",
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
# 个90

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
  
  # 临时存储已出现的名称及其出现次数
  name_counts <- list()
  
  # 遍历元素，添加后缀
  for (i in 1:length(selected_data$qymc)) {
    name <- selected_data$qymc[i]
    
    # 检查该名称是否已经出现过
    if (is.null(name_counts[[name]])) {
      # 如果没有出现过，将其加入列表，并添加后缀"_year"
      name_counts[[name]] <- 1
      selected_data$qymc[i] <- paste(name, year, "1" ,sep = "_")
    } else {
      # 如果已经出现过，将计数加一，并添加对应的后缀
      count <- name_counts[[name]]
      name_counts[[name]] <- count + 1
      selected_data$qymc[i] <- paste(name,year, count+1, sep = "_")
    }
  }
  
  # 显示处理后的数据
  #print(selected_data$qymc)
  
  
  
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

# 移除 qymc 列，因为你想要将其作为列名
indus_zong_no_qymc <- indus_zong[, -which(names(indus_zong) == "qymc")]

# 使用 t() 函数将数据转置
wide_table <- t(indus_zong_no_qymc)

# 设置列名为 qymc 列的值
colnames(wide_table) <- indus_zong$qymc

# 创建一个新的数据框，第一列是除了 qymc 之外的变量名
wide_table_new <- data.frame(Variable_names = names(indus_zong_no_qymc))

# 将 indus_zong_no_qymc 中的数据作为 wide_table 的第二列
wide_table <- cbind(wide_table_new, wide_table)

# 选择前10列
wide_table_subset <- wide_table[, 1:10]

Variable_name_Ch <- selected_cols_zw[-1]
Year_of_presenting <- processed_years[-1]

wide_table_subset$Year_of_presenting <- NA
wide_table_subset <- wide_table_subset %>%
  mutate(
    Variable_name_Chinese = Variable_name_Ch,
  ) %>%
  select(Variable_names,Variable_name_Chinese, Year_of_presenting, everything())
# 将 Variable_name_Chinese 和 Year_of_presenting 列移动到第二列和第三列
for (i in seq_along(Year_of_presenting)) {
  wide_table_subset$Year_of_presenting[i] <- Year_of_presenting[[i]]
}


#这里要重新弄



# 将数据框写入工作表中
writeData(wb1, sheet = "all", wide_table_subset, startCol = 1, startRow = 1)
# 保存工作簿为Excel文件
saveWorkbook(wb1, file = output_file2, overwrite = TRUE)
# 写入新的 CSV 文件
write.csv(wide_table, file = output_file1, row.names = FALSE, fileEncoding = "UTF-8")

cat("数据整理完成并保存到 all_ind_env.xlsx 文件中\n")
