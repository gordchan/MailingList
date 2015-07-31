## A function to return all sheets in xlsx to a list with dataframe elements

# library(xlsx)
# 
# x <- "input/PMH Site Info.xlsx"

read_all_sheets <- function (x){
  
  require(xlsx)
  
  input <- x
  
  wb <- loadWorkbook(input)
  sheets <- getSheets(wb)
  sheet_no <- length(sheets)
  
  sheet_index <- c(1:sheet_no)
  
  all_sheets <- lapply(sheet_index, function (x) {read.xlsx2(file = input, sheetIndex = x)})
  
  names(all_sheets) <- names(sheets)
  
  all_sheets
  
}

# all_sheets <- read_all_sheets(input)