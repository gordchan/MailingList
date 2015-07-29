dc_mailing <- function(x = "PMH Site Info.xlsx") {
  
  ## To generate mailing list from xlsx file
  ## With multiple lines in single cells
  
  library(xlsx)
  library(dplyr)
  
  ## Prepare filepath of i/o of file
  
  input <- paste ("input/", x, sep = "")

  if (grepl("*PMH*", x)){
    list_body <- "_PMH"
  } else if (grepl("*KWC*", x)){
    list_body <- "_KWC"
  }
  
  if (list_body=="_PMH"){
    
    ## Read source xlsx from working directory
    
    source_d_dco <- read.xlsx(input, 1)
    source_c_dco <- read.xlsx(input, 2)
    
    ## Extract email addresses
    
    source_list_dco <- c(source_d_dco[5], source_c_dco[5])
    source_list_ta <- c(source_d_dco[6], source_c_dco[6])
    
    source_list_dco <- unlist(source_list_dco)
    source_list_ta <- unlist(source_list_ta)
    
    source_list_dco <- as.character.factor(source_list_dco)
    source_list_ta <- as.character.factor(source_list_ta)
    
    ## Seperate according to line break 
    
    email_list_dco <- strsplit(source_list_dco, "(\\n|  +)")
    email_list_ta <- strsplit(source_list_ta, "(\\n|  +)")
    
    ## Flatten all_address list to vector
    
    all_dco <- unlist(email_list_dco)
    all_ta <- unlist(email_list_ta)
    
    ## Remove comma in tail
    
    all_dco <- sapply(all_dco, FUN = function (x) gsub(",$", "", x))
    all_ta <- sapply(all_ta, FUN = function (x) gsub(",$", "", x))
    
    ## Convert flattened list to data frame
    
    all_dco <- as.data.frame(all_dco)
    all_ta <- as.data.frame(all_ta)
    
    ## Generate distinct address lists
    
    mailing_list_dco <- distinct(all_dco)
    mailing_list_ta <- distinct(all_ta)

    mailing_list_dco$all_dco[mailing_list_dco$all_dco == ""] <- NA
    mailing_list_ta$all_ta[mailing_list_ta$all_ta == ""] <- NA
    
    mailing_list_dco <- mailing_list_dco %>% na.omit()
    mailing_list_ta <- mailing_list_ta %>% na.omit()
    
    colnames(mailing_list_dco) <- "All_DCOs"
    colnames(mailing_list_ta) <- "All_TAs"
    
    mailing_list_dco <- mailing_list_dco %>% arrange(All_DCOs)
    mailing_list_ta <- mailing_list_ta %>% arrange(All_TAs)
    
    ## Write txt for Outlook
    
    write.table(mailing_list_dco, file=paste("output/mailing_list", list_body, "_DCO_Outlook.txt", sep = ""), quote =FALSE, eol =";", na="", row.names=FALSE, col.names=FALSE)
    write.table(mailing_list_ta, file=paste("output/mailing_list", list_body, "_TA_Outlook.txt", sep = ""), quote =FALSE, eol =";", na="", row.names=FALSE, col.names=FALSE)
    
    ## Write txt for Human
    
#     write.table(mailing_list_dco, file=paste("output/mailing_list", list_body, "_DCO.txt", sep = ""), quote =FALSE, na="", row.names=FALSE, col.names=FALSE)
#     write.table(mailing_list_ta, file=paste("output/mailing_list", list_body, "_TA", sep = ""), quote =FALSE, na="", row.names=FALSE, col.names=FALSE)
    
    ## Write xlsx
    
    xlsx_file <- paste("output/mailing_list", list_body, ".xlsx", sep = "")
    
    write.xlsx(x = mailing_list_dco, file = xlsx_file, sheetName = "PMH",  col.names=TRUE, row.names = FALSE, showNA = FALSE)
    
    wb <- loadWorkbook(xlsx_file)
    sheets <- getSheets(wb)
    sheet <- sheets[[1]]
    
    addDataFrame(x = mailing_list_ta, sheet, col.names=TRUE, row.names = FALSE, startColumn=2, showNA = FALSE)
    
    saveWorkbook(wb, file = xlsx_file)
    
    
  }else if (list_body=="_KWC"){
    
    ## Read source xlsx from working directory
    
    source_c_dco <- read.xlsx(input, 1)
    
    ## Extract email addresses
    
    source_list_dco <- source_c_dco[5]
    source_list_ta <- source_c_dco[6]
    
    source_list_dco <- unlist(source_list_dco)
    source_list_ta <- unlist(source_list_ta)
    
    source_list_dco <- as.character.factor(source_list_dco)
    source_list_ta <- as.character.factor(source_list_ta)
    
    ## Seperate according to line break 
    
    email_list_dco <- strsplit(source_list_dco, "(\\n|  +)")
    email_list_ta <- strsplit(source_list_ta, "(\\n|  +)")
    
    ## Flatten all_address list to vector
    
    all_dco <- unlist(email_list_dco)
    all_ta <- unlist(email_list_ta)
    
    ## Remove comma in tail
    
    all_dco <- sapply(all_dco, FUN = function (x) gsub(",$", "", x))
    all_ta <- sapply(all_ta, FUN = function (x) gsub(",$", "", x))
    
    ## Convert flattened list to data frame
    
    all_dco <- as.data.frame(all_dco)
    all_ta <- as.data.frame(all_ta)
    
    ## Generate distinct address lists
    
    mailing_list_dco <- distinct(all_dco)
    mailing_list_ta <- distinct(all_ta)
    
    mailing_list_dco$all_dco[mailing_list_dco$all_dco == ""] <- NA
    mailing_list_ta$all_ta[mailing_list_ta$all_ta == ""] <- NA
    
    mailing_list_dco <- mailing_list_dco %>% na.omit()
    mailing_list_ta <- mailing_list_ta %>% na.omit()
    
    colnames(mailing_list_dco) <- "All_DCOs"
    colnames(mailing_list_ta) <- "All_TAs"
    
    mailing_list_dco <- mailing_list_dco %>% arrange(All_DCOs)
    mailing_list_ta <- mailing_list_ta %>% arrange(All_TAs)
    
    ## Write txt for Outlook
    
    write.table(mailing_list_dco, file=paste("output/mailing_list", list_body, "_DCO_Outlook.txt", sep = ""), quote =FALSE, eol =";", na="", row.names=FALSE, col.names=FALSE)
    write.table(mailing_list_ta, file=paste("output/mailing_list", list_body, "_TA_Outlook.txt", sep = ""), quote =FALSE, eol =";", na="", row.names=FALSE, col.names=FALSE)
    
    ## Write txt for Human
    
#     write.table(mailing_list_dco, file=paste("output/mailing_list", list_body, "_DCO.txt", sep = ""), quote =FALSE, na="", row.names=FALSE, col.names=FALSE)
#     write.table(mailing_list_ta, file=paste("output/mailing_list", list_body, "_TA.txt", sep = ""), quote =FALSE, na="", row.names=FALSE, col.names=FALSE)
    
    ## Write xlsx
    
    xlsx_file <- paste("output/mailing_list", list_body, ".xlsx", sep = "")
    
    write.xlsx(x = mailing_list_dco, file = xlsx_file, sheetName = "KWC",  col.names=TRUE, row.names = FALSE, showNA = FALSE)
    
    wb <- loadWorkbook(xlsx_file)
    sheets <- getSheets(wb)
    sheet <- sheets[[1]]
    
    addDataFrame(x = mailing_list_ta, sheet, col.names=TRUE, row.names = FALSE, startColumn=2, showNA = FALSE)
    
    saveWorkbook(wb, file = xlsx_file)
    
  }
    
  }

