verify <- function(file = "mailing list.txt", name = "mailing_list"){
    
  ## Read through email validated by Outlook "To:" field and return whether the email has been validated
  
  require(dplyr)
  require(xlsx)
  
  input <- file.path("input", file)
  
  verified <- scan(input, what = "character", sep=";")

  verified <- sapply(verified, FUN = function(x) gsub("(^ )", "", x))
    
  ver_df <- data.frame(verified)
  
  ver_df <- ver_df %>% mutate(valid = grepl("*<*>$", verified))
  
  output <- file.path("output", "mailing list.xlsx")
  
  write.xlsx(ver_df, file = output, sheetName = name, row.names = FALSE)
  
    wb <- loadWorkbook(output)
      sheets <- getSheets(wb)
        sheet <- sheets[[1]]
  
        setColumnWidth(sheet, colIndex=1, colWidth=60)
        setColumnWidth(sheet, colIndex=2, colWidth=24)
        
        row <- getRows(sheet, rowIndex=1)
        cb <- CellBlock(sheet, startRow=1, startColumn=1, noRows=1, noColumns=2, create=FALSE)
        
        header <- Font(wb, isBold=TRUE)
        
        CB.setFont(cellBlock=cb, font=header, rowIndex=1, colIndex=1:2)
        
    saveWorkbook(wb, output)
        
}

verify(name = "KWC_Q&S_Forum_2016_OC")
