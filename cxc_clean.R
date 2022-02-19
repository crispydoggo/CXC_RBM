

# CLEAR LAC
# Spring 2022
# CXC RBM Policy

# ==========================
# Packages 
# ==========================
library(haven)
library(readr)
library(tidyverse)
library(XLConnect) #Writes Excel Sheets for easy reading
library(readxl)

# ==========================
# Data
# ==========================

CARICOM <-  read_delim("cxc.txt", delim = "$")

# Correct some features of data

CARICOM <- CARICOM %>% 
  select(bullet_id = `bullet_id `, comment = ` comment` )%>% 
  mutate( bullet_id = as.integer(bullet_id))


# ==========================
# Functions 
# ==========================


sort_rbm <- function(CARICOM){
  # Sorts the comments by bullet id and exports them into an excel file
  # ----------------------------------
  
  book <- loadWorkbook("RBM_com.xlsx")
  RBM <- tibble()
  
  for( i in 1:max(CARICOM$bullet_id)){
    RBM <- CARICOM %>% 
    filter( bullet_id == i)
    createSheet(book, name = paste("bullet_", i, sep = ""))
    writeWorksheet(book, RBM, sheet = paste("bullet_", i, sep = ""))
  }
  saveWorkbook(book, file = "CXC_clean.xlsx")
}

add_rbm <- function(file, n){
  
  RBM_info <- read_excel(file, sheet = 1)
  RBM_info <- arrange(RBM_info, sub_id, order)
  for( i in 2:n){
    
    aux <- read_excel(file, sheet = i)
    aux <- arrange(aux, sub_id, order)
    RBM_info <- bind_rows(RBM_info, aux)
    
  }
  
  RBM_info <- RBM_info %>% 
    select( bullet_id, sub_id, comment,order) %>% 
    mutate(bullet_id = as.integer(bullet_id), sub_id = as.integer(sub_id)) 
  
  return(RBM_info)
  
}

# ==========================
# Data Cleaning
# ==========================


# Arranges the raw Comment section 
sort_rbm(CARICOM)

# Process the ordered Comment Section
RBM_info <- add_rbm("CXC_clean.xlsx",20)


# ==========================
# Export Data
# ==========================


book <- loadWorkbook("RBM_com.xlsx")
writeWorksheet(book, RBM_info, sheet = 1)
saveWorkbook(book, file = "RBM_CXC.xlsx")


# ==============================
# Import Clean Data and Rearrange
# ===============================






