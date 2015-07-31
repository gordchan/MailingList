# mailing <- function(x = "PMH Site Info.xlsx") {
  
  ## To generate mailing list from xlsx file
  ## With multiple lines in single cells
  
  x <- "PMH Site Info.xlsx"
  email_cols <- c(5,6)

  
  library(xlsx)
  library(dplyr)
  
  source("read_all_sheets.R")
  
  ## Prepare filepath of i/o of file
  
  input <- paste ("input/", x, sep = "")
  
  col_no <- length(email_cols)
  
  sheets <- getSheets(loadWorkbook(input))
  sheet_no <- length(sheets)

  if (grepl("*PMH*", x)){
    list_body <- "_PMH"
  } else if (grepl("*KWC*", x)){
    list_body <- "_KWC"
  }
  
  ## Read source xlsx from working directory

  all_sheets <- read_all_sheets(input)
  
  ## Extract all email
  
  email <- list(letters[1:(length(email_cols)*sheet_no)])

  k <- 1
  
  for (i in 1:sheet_no){
    for (j in 1:col_no){
      
      email[k]  <- list(all_sheets[[i]][email_cols[j]])
      names(email)[k] <- paste(sheets[i], j, sep = "")
      
      k <<- k+1
    }
  }
  
  
  email_temp <- data.frame()
  
  list_index <- seq(1, sheet_no*col_no, col_no)
  
  for (i in 1: sheet_no){
    
    email_temp <- rbind(email_temp, email[list_index[i]])
    
  }
  
unlist(email_temp)
  