mailing <- function(file = "mailing list.xlsx", name){
  
  ## Read through email validated by Outlook "To:" field and return whether the email has been validated
  
  require(dplyr)
  require(xlsx)
  
  input <- file.path("input", file)
  
  mailing <- read.xlsx(input, 1)
  
  ver_df <- mailing %>% filter(valid == TRUE) %>% select(verified)
  
  write.table(ver_df, file ="output/mailing list.txt", quote = FALSE, sep = ";", eol = ";", row.names = FALSE, col.names = FALSE)
  
}

mailing()