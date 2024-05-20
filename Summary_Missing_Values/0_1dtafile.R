# load library haven
library(haven)
library(rio)

input_folder_path <- ""
output_folder_path <- ""

# 需要处理的年份列表
years <- c("1998","1999","2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010","2011","2012","2013","2014") 

# 循环处理每个年份的文件
for (year in years) {
  input_file_path <- paste0(input_folder_path,"工业企业与污染匹配结果", year, "年.dta")
  output_file_path <- paste0(output_folder_path, year, "_new.csv")
  
 
  # Vector to store NA counts for each combination of column classes
  na_counts <- vector("list", length = 3)
  
  # Iterate over different combinations of column classes
  for (num_class in c("numeric", "integer", "double")) {
    for (char_class in c("character", "factor")) {
      for (logical_class in c("logical", "integer")) {
        col_classes <- c(num_class, char_class, logical_class)
        data <- tryCatch(import(input_file_path, col_types = col_classes),
                         error = function(e) return(NULL))
        if (!is.null(data)) {
          na_counts[[paste(num_class, char_class, logical_class)]] <- sum(is.na(data))
        }
      }
    }
  }
  
  # Find the combination with the fewest NA values
  min_na_combination <- names(na_counts)[which.min(unlist(na_counts))]
  
  # Read the data with the combination of column classes that had the fewest NA values
  data <- import(input_file_path, col_types = strsplit(min_na_combination, " ")[[1]])
  
  # # import .dat file
  # data <- read_dta(input_file_path)
  
  # Check for missing values (NA or NULL) in the data frame
  if (anyNA(data) || any(is.null(data))) {
    print("The data frame contains missing values.")
  } else {
    print("The data frame does not contain missing values.")
  }
  
  # # Convert string variables to numeric
  # string_vars <- sapply(data, is.character) # Identify character variables
  # data[string_vars] <- lapply(data[string_vars], as.numeric) # Convert character variables to numeric

  for (col in names(data)) {
    if (is.character(data[[col]])) {  # Check if column is character type
      converted_col <- suppressWarnings(as.numeric(data[[col]]))  # Try to convert to numeric
      if (!any(is.na(converted_col))) {  # Check if conversion successful
        data[[col]] <- converted_col  # Update column with numeric values
      }
    }
  }
  
  # print head and summary of data frame
  print(paste("Top 6 entries of data frame for year", year, ":"))
  head(data)
  print(paste("Summary for year", year, ":"))
  summary(data)
  
  # Write data to csv
  write.csv(data, output_file_path, row.names = FALSE, fileEncoding = "UTF-8")
  
}

