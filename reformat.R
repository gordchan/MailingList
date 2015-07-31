## Aid to reformat the DCO and TA list

## Return number of DCO & TA for each dept/comm

x <- "input/PMH Site Info.xlsx"

library(xlsx)
library(dplyr)
library(tidyr)

source("read_all_sheets.R")

all_sheets <- read_all_sheets(x)

## Add index/id number

dept <- all_sheets$Department
  i_dept <- (1:nrow(dept))
    dept <- cbind(dept, i_dept)

comm <- all_sheets$Committee
  i_comm <- (1:nrow(comm))
    comm <- cbind(comm, i_comm)

skip_i <- seq(1, nrow(comm), 2)
    
#     rm(i_dept)
#     rm(i_comm)
    
## For Department List
  ## DCO
    dept_dco_source <- as.character(dept$DCO)
    dept_dco_source <- sapply(dept_dco_source, FUN = function(x){if (x=="") "No DCO" else x})
    dept_dco <- strsplit(dept_dco_source, "(\\n|  +)")
    
    rm(dept_dco_source)
    
    DCO_number <- sapply(dept_dco, FUN = function(x)length(x))
    dept <- cbind(dept, DCO_number)
    
    names(dept_dco) <- dept$Department
    
  ## TA
    dept_ta_source <- as.character(dept$TA)
    dept_ta_source <- sapply(dept_ta_source, FUN = function(x){if (x=="") "No TA" else x})
    dept_ta <- strsplit(dept_ta_source, "(\\n|  +)")
    
    rm(dept_ta_source)
    
    TA_number <- sapply(dept_ta, FUN = function(x)length(x))
    dept <- cbind(dept, TA_number)
    
    names(dept_ta) <- dept$Department
    
   ## APProver
    dept_app_source <- as.character(dept$Approver)
    dept_app_source <- sapply(dept_app_source, FUN = function(x){if (x=="") "No Approver" else x})
    dept_app <- strsplit(dept_app_source, "(\\n|  +)")
    
    rm(dept_app_source)
    
    APP_number <- sapply(dept_app, FUN = function(x)length(x))
    dept <- cbind(dept, APP_number)
    
    names(dept_app) <- dept$Department
    
    Dept_numbers <- select(dept, Department, DCO_number, TA_number, APP_number)
    
## For Committee List
    ## DCO
    comm_dco_source <- as.character(comm$DCO)
    comm_dco_source <- sapply(comm_dco_source, FUN = function(x){if (x=="") "No DCO" else x})
    comm_dco <- strsplit(comm_dco_source, "(\\n|  +)")
    
    rm(comm_dco_source)
    
    DCO_number <- sapply(comm_dco, FUN = function(x)length(x))
    comm <- cbind(comm, DCO_number)
    
    names(comm_dco) <- comm$Committee
    
    ## TA
    comm_ta_source <- as.character(comm$TA)
    comm_ta_source <- sapply(comm_ta_source, FUN = function(x){if (x=="") "No TA" else x})
    comm_ta <- strsplit(comm_ta_source, "(\\n|  +)")
    
    rm(comm_ta_source)
    
    TA_number <- sapply(comm_ta, FUN = function(x)length(x))
    comm <- cbind(comm, TA_number)
    
    names(comm_ta) <- comm$Committee
    
    ## APProver
    comm_app_source <- as.character(comm$Approver)
    comm_app_source <- sapply(comm_app_source, FUN = function(x){if (x=="") "No Approver" else x})
    comm_app <- strsplit(comm_app_source, "(\\n|  +)")
    
    rm(comm_app_source)
    
    APP_number <- sapply(comm_app, FUN = function(x)length(x))
    comm <- cbind(comm, APP_number)
    
    names(comm_app) <- comm$Department
    
    Comm_numbers <- select(comm, Committee, DCO_number, TA_number, APP_number)
    

    ## Remove duplicates
    
    comm_dco <- comm_dco[skip_i]
    comm_ta <- comm_ta[skip_i]
    comm_app <- comm_app[skip_i]
    
    rm(DCO_number)
    rm(TA_number)
    rm(APP_number)
    
    
## Indexing for Dept/Comm number list
    
    Dept_numbers <- cbind(i_dept, Dept_numbers)
    Comm_numbers <- cbind(i_comm, Comm_numbers)
    
    Comm_numbers <- filter(Comm_numbers, i_comm %in% skip_i)
    
    
## Generate email list for each dept/comm
    
    ## Dept
    new_dept_dco <- data.frame(character(), character())
    new_dept_ta <- data.frame(character(), character())
    new_dept_app <- data.frame(character(), character())
    
    for (i in 1:nrow(Dept_numbers)){
      
      c1 <- names(dept_dco)[i]
      names(dept_dco)[i] <- "Document Control Officer"
      names(dept_ta)[i] <- "Technical Assistant"
      names(dept_app)[i] <- "Approver"
      
      c2 <- dept_dco[i]
      c3 <- dept_ta[i]
      c4 <- dept_app[i]
      
      
      new_dept_dco_temp <- data.frame(Department = c1, DCO = c2)
      new_dept_ta_temp <- data.frame(Department = c1, TA= c3)
      new_dept_app_temp <- data.frame(Department = c1, Approver = c4) 

      
      new_dept_dco <- rbind(new_dept_dco, new_dept_dco_temp)
      new_dept_ta <- rbind(new_dept_ta, new_dept_ta_temp)
      new_dept_app <- rbind(new_dept_app, new_dept_app_temp)
      
    }
    
    rm(new_dept_dco_temp)
    rm(new_dept_ta_temp)
    rm(new_dept_app_temp)
    
    ## Comm
    new_comm_dco <- data.frame(character(), character())
    new_comm_ta <- data.frame(character(), character())
    new_comm_app <- data.frame(character(), character())
    
    for (i in 1:nrow(Comm_numbers)){
      
      c1 <- names(comm_dco)[i]
      names(comm_dco)[i] <- "Document Control Officer"
      names(comm_ta)[i] <- "Technical Assistant"
      names(comm_app)[i] <- "Approver"
      
      c2 <- comm_dco[i]
      c3 <- comm_ta[i]
      c4 <- comm_app[i]
      
#       names(comm_dco)[i] <- "DCO"
#       names(comm_ta)[i] <- "TA"
      
      new_comm_dco_temp <- data.frame(Committee = c1, DCO = c2)
      new_comm_ta_temp <- data.frame(Committee = c1, TA= c3)
      new_comm_app_temp <- data.frame(Committee = c1, TA= c4)
      
      
      new_comm_dco <- rbind(new_comm_dco, new_comm_dco_temp)
      new_comm_ta <- rbind(new_comm_ta, new_comm_ta_temp)
      new_comm_app <- rbind(new_comm_app, new_comm_app_temp)
      
    }
    
    rm(new_comm_dco_temp)
    rm(new_comm_ta_temp)
    rm(new_comm_app_temp)
    
## Reformat into tidy data
    
    tidy_dept_dco <- gather(new_dept_dco, "Role", "email", 2)
    tidy_dept_ta <- gather(new_dept_ta, "Role", "email", 2)
    tidy_dept_app <- gather(new_dept_app, "Role", "email", 2)
    
    tidy_dept <- bind_rows(tidy_dept_dco, tidy_dept_ta, tidy_dept_app)
    
    tidy_comm_dco <- gather(new_comm_dco, "Role", "email", 2)
    tidy_comm_ta <- gather(new_comm_ta, "Role", "email", 2)
    tidy_comm_app <- gather(new_comm_app, "Role", "email", 2)
    
    tidy_comm <- bind_rows(tidy_comm_dco, tidy_comm_ta, tidy_comm_app)
    
    rm(tidy_dept_dco)
    rm(tidy_dept_ta)
    rm(tidy_dept_app)
    
    rm(tidy_comm_dco)
    rm(tidy_comm_ta)
    rm(tidy_comm_app)
    
    ## Sort tidy list by Dept/Comm
    
    tidy_dept <- tidy_dept %>% arrange(Department, Role)
    
    tidy_comm <- tidy_comm %>% arrange(Committee, Role)
    
    ## Output tidy list to xlsx file
    
    write.xlsx(tidy_dept, file = "output/PMH Document Control users.xlsx", sheetName = names(all_sheets)[1], append = FALSE)
    
    write.xlsx(tidy_comm, file = "output/PMH Document Control users.xlsx", sheetName = names(all_sheets)[2], append = TRUE)
    
    ## Formatting output xlsx file
    
    