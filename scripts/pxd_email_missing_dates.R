#*************************************************
#name: pq_email_missing_dates.R
#purpose: check procxed output for dates
#that are missing DD-MM-YYYY format and email
#the contacts listed from statsgov mailbox

#written by: Graeme 16/12/19
#*************************************************

#***************************************
#install packages####
#***************************************
#RDCOMClient won't install on linux server
#install.packages("RDCOMClient", repos = "http://www.omegahat.net/R")
#install.packages("tidyverse")
#install.packages("tidylog)
#install.packages("lubridate)
#install.packages("readxl")
#install.packages("glue")
#install.packages("janitor")

#***************************************
#load packages####
#***************************************

#load packages
library(RDCOMClient)
library(tidyverse)
library(tidylog)
library(lubridate)
library(readxl)
library(glue)
library(janitor)

#***************************************
#define variables and load data####
#***************************************

#define filepath for procxed output
path <- "path/to/procxed/output"
file_to_use <- "ISD Forthcoming Publications Nov-19.xlsx"

#only check dates before this 
#e.g. end of previous month
date_limit <- "2020-04-01"

#get contact info from file
emails <- read_csv("data/email_addresses.csv")
stats_gov <- emails %>% 
              filter(contact == "stats_gov") %>% 
              select(info) %>% 
              paste0()

#generate email intro
email_intro <- glue(
  "Hi all,<br><br>",
  "We're about to submit ProcXed for the month - would you be able to add finalised",
  "publication dates (DD-MM-YYYY) for the publications listed below? Although these",
  "dates are 'finalised' they can still be changed up to 6 weeks before.<br><br>",
  "Thanks,<br><br>",
  "StatsGov<br><br>")

#***************************************
#process data####
#***************************************

#read pubs
pubs <- read_xlsx(glue("{path}{file_to_use}"))
pubs <- pubs %>% clean_names()

#get email addresses from contact info
pubs <- pubs %>% 
        mutate(cd = paste0(str_extract_all(contact_details, ".+@.+")),
              cd = str_replace_all(cd, "e-mail: ", ""),
              cd = str_replace_all(cd, pattern = "c\\(", replacement = ""),
              cd = str_replace_all(cd, pattern = '\"', ""),
              cd = str_replace_all(cd, pattern = '\\)', ""),
              cd = str_replace_all(cd, pattern = ",", ";"))

#generate temp date columns
pubs <- pubs %>% mutate(temp_date = if_else(is.na(dmy(publication_date)), 
                                    dmy(paste0("01-", publication_date)),
                                    dmy(publication_date)))

#only keep months of interest
#typically 5 months in advance
#e.g. by end of Dec, finalize May
pubs <- pubs %>% filter(temp_date < date_limit)

#which dates are missing?
#if date doesn't parse as dmy then flag
pubs <- pubs %>% mutate(flagged = if_else(
                            is.na(dmy(publication_date)),
                            TRUE,
                            FALSE))

#keep only flagged pubs
flagged <- pubs %>% filter(flagged == TRUE)

#***************************************
#send one email to everyone####
#***************************************

flagged <- flagged %>% 
          mutate(msg = glue(
          "<span style='font-size:13px; font-family:arial'>",
          "<strong>{publication_series}</strong><br>",
          "<strong>Date to edit: {publication_date}</strong><br>",
          "{contact_details}<br>",
          "{synopsis}<br>",
          "<br></span>"))

#collapse message
msg_body <- glue_collapse(flagged$msg)

#collapse contact info
send_to <- glue_collapse(flagged$cd, sep = "; ")
                    
#create an email
OutApp <- COMCreate("Outlook.Application")
outMail <- OutApp$CreateItem(0)

#configure email parameters
outMail[["To"]] = send_to
outMail[["SentOnBehalfOfName"]] = stats_gov
outMail[["subject"]] <- "ProcXed - dates missing"
outMail[["htmlbody"]] <- paste0(email_intro, msg_body)

#send it
outMail$Send()

#***************************************
#send one email per contact group####
#***************************************

#for any pubs that have the same authors
#send one email to each set
#group by contact info
grp_email <- flagged %>% 
            group_by(cd) %>%
            summarize(pooled = glue_collapse(msg))

#for each group:
for (i in 1:nrow(grp_email)) {
  
  #extract group by row
  x <- slice(grp_email, i)
  
  #extract info for emails
  send_to <- str_squish(x$cd)
  subject <- "ProcXed - dates missing"
  msg_body <- paste0(email_intro, "<br><br>", x$pooled)
  
  #create an email
  OutApp <- COMCreate("Outlook.Application")
  outMail <- OutApp$CreateItem(0)
  
  #configure email parameters
  outMail[["To"]] <- send_to
  
  outMail[["SentOnBehalfOfName"]] <- stats_gov
  outMail[["subject"]] <- subject
  outMail[["htmlbody"]] <- msg_body
  
  #send it
  outMail$Send()
}
