verify <- function(x){

  # x = "mailing list.txt"
    
  ## Read through email validated by Outlook "To:" field and return whether the email has been validated
  
  library(dplyr)
  library(xlsx)
  
  input <- paste("input/", x, sep = "")
  
  verified <- scan(input, what = "character", sep=";")

  verified <- sapply(verified, FUN = function(x) gsub("(^ )", "", x))
    
  ver_df <- data.frame(verified)
  
  ver_df <- ver_df %>% mutate(valid = grepl("*<*>$", verified))
  
  write.xlsx(ver_df, file ="output/verified_address.xlsx", sheetName="verified_email", row.names = FALSE)
  
}

verify("mailing list.txt")
