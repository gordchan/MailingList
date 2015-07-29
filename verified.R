verify <- function(x){
  
  ## Read through email validated by Outlook "To:" field and return whether the email has been validated
  
  library(dplyr)
  
  input <- paste("input/", x, sep = "")
  
  verified <- scan(input, what = "character", sep=";")
  
  ver_df <- data.frame(verified)
  
  ver_df <- ver_df %>% mutate(valid = grepl("*<*>$", verified))
  
  write.csv(ver_df, file ="output/verified_address.csv")
  
}



