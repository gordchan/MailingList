## A function to return all sheets in xlsx to a list with dataframe elements

library(xlsx)

x <- "PMH Site Info.xlsx"
input <- paste ("input/", x, sep = "")

read_all_sheets <- function (x){
  
  require(xlsx)
  
  wb <- loadWorkbook(input)
  sheets <- getSheets(wb)
  sheet_no <- length(sheets)
  
  sheet_index <- c(1:sheet_no)
  
  all_sheets <- lapply(sheet_index, function (x) {read.xlsx2(file = input, sheetIndex = x)})
  
  names(all_sheets) <- sheets
  
}

sheets <- read_all_sheets(input)
