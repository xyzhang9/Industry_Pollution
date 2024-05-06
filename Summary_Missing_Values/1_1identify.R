library(dplyr)

# 设置 R 选项，避免写入 CSV 文件时出现科学技术法
options(scipen = 999)

# 定义方位映射函数
get_direction <- function(province) {
  direction_map <- c(
    "内蒙古自治区"  =  "西部",
    "广西壮族自治区" = "西部",
    "重庆市" = "西部",
    "四川省" = "西部",
    "贵州省" = "西部",
    "云南省" = "西部",
    "西藏自治区" = "西部",
    "陕西省" = "西部",
    "甘肃省" = "西部",
    "青海省" = "西部",
    "宁夏回族自治区" = "西部",
    "新疆维吾尔自治区" = "西部",
    
    "北京市" = "东部",
    "天津市" = "东部",
    "河北省" = "东部",
    "上海市" = "东部",
    "江苏省" = "东部",
    "浙江省" = "东部",
    "山东省" = "东部",
    "福建省" = "东部",
    "广东省" = "东部",
    "海南省" = "东部",
    
    "山西省" = "中部",
    "安徽省" = "中部",
    "江西省" = "中部",
    "河南省" = "中部",
    "湖北省" = "中部",
    "湖南省" = "中部",
    
    "吉林省" = "东北",
    "辽宁省" = "东北",
    "黑龙江省" = "东北"
  )
  
  return(direction_map[province])
}

# 指定 CSV 文件路径的文件夹
input_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/"
output_folder_path <- "D:/照片材料/文件/应聘科研助理_炊晨阳_武汉大学_18028377912/环境/数据/计算1/"

# 需要处理的年份列表
years <- c("1998","1999","2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010","2011","2012","2013","2014") 

# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_path, year, ".csv")
  output_file_path <- paste0(output_folder_path, year, ".csv")

  # 读取原始的 1998.csv 文件，指定文件编码为 UTF-8
  data <- read.csv(input_file_path, fileEncoding = "UTF-8")
  
  # 继续处理数据
  
  # 挑选需要的列
  selected_cols <- c(
    "id", "qymc", "frdm", "frdbxm",
    "province", "city", "town", "county", "c", "jdbsc", "jwh", "sdxm", "xcm", "yzbm", "dqdm", "行政区代码","行政区划代码", "qh", "dmdm", "dz", 
    "hylb", "hymlmc", "hydldm", "hydlmc", "hyzldm", "hyzlmc", "hyxldm", "hyxlmc", "cp1", "cp2", "cp3", 
    "gyzczbbjxgd", "工业总产值现价万元", "工业总产值现价", "gyxsczxjxgd", "gyzczxjxgd", "zysr", "yysr", "zyywsr", 
    "cyrs", "cyrym", "cyryn", "yjsm", "yjsn", "bksm", "bksn", "zksm", "zksn", "gzsm", "gzsn", "czyxm", "czyxn", "gjzcm", "gjzcn", "zjzcm", "zjzcn", "djzcm", "djzcn", "gjjsm", "gjjsn", "jsm", "jsn", "gjgm", "gjgn", "zjgm", "zjgn", 
    "煤炭消费总量吨", "其中燃料煤消费量吨", "原料煤消费量吨", "料油消费量不含车船用吨", "其中重油吨", "柴油吨", "洁净燃气消费量万立方米", 
    "工业用水总量吨", "其中新鲜水量吨", "重复用水量吨", 
    "燃料煤平均硫份燃", "重油平均硫份", "工业废水排放量吨", "化学需氧量排放量千克", "氨氮排放量千克", "工业废气排放总量万标立方米", "氮氧化物排放量千克", "二氧化硫排放量千克", "烟尘排放量千克", "工业粉尘排放量千克",
    "废水治理设施数套", "废水治理设施处理能力吨日", "工业废水处理量吨", "废气治理设施数套", "废气治理设施处理能力标立方米时", "氨氮去除量千克", "化学需氧量去除量千克", "氮氧化物去除量千克", "其中脱硫设施数套", "其中脱硫设施脱硫能力千克时", "二氧化硫去除量千克", "其中当年新增设施去除的千克", "烟尘去除量千克", "工业粉尘去除量千克", 
    "sszb", "gjzbj", "jtzbj", "frzbj", "grzbj", "gatzbj", "wszbj", "btsr", "gdzchj", "gdzcyjhj", "gdzcjznpjye",
    "zjtrhj", "zzzjtr", "glzjtr", "yyzjtr"
  )
  # 个116

  # 创建一个空的数据框，用于存储最终结果
  selected_data <- data.frame(matrix(ncol = length(selected_cols) + 1, nrow = nrow(data)))
  colnames(selected_data) <- c(selected_cols, "方位")
  
  # 将原始数据中的选定列复制到新的数据框中，并将非标准的 "NA" 值替换为 NA
  for (col in c(selected_cols, "方位")) {
    if (col %in% colnames(data)) {
      # 检查每个单元格，将空格替换为 NA
      selected_data[[col]] <- lapply(data[[col]], function(x) ifelse(x %in% c("", "NA", "N/A"), NA, x))
    } else {
      selected_data[[col]] <- NA
    }
  }
  
  # 添加方位列
  #selected_data$方位 <- sapply(selected_data$province, get_direction)
  selected_data$方位 <- ifelse(is.na(selected_data$province), NA, sapply(selected_data$province, get_direction))
  
  # 检查并转换列表列
  for (col in colnames(selected_data)) {
    if (is.list(selected_data[[col]])) {
      selected_data[[col]] <- sapply(selected_data[[col]], function(x) paste(x, collapse = ","))
    }
  }
  
  # 调整列的顺序
  selected_data <- selected_data %>%
    select(id, qymc, frdm, province, 方位, everything())
  
  # 写入新的 CSV 文件
    write.csv(selected_data, file = output_file_path, row.names = FALSE, fileEncoding = "UTF-8")
  
  # 输出成功信息
  cat("已经生成新的 CSV 文件 .csv\n")
}