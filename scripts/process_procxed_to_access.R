##########################################################
# Process ProcXed to Access
# Graeme Gowans
# October 2019
# Latest update author - Graeme
# Latest update date - March 2020
# reflect simpler process for beta website
# Extraction
# R Studio Desktop
# 3.5.1
# Description of content
# processes output from ProcXed monthly report to format 
# required for Access to update ISD Forthcoming page
# Approximate run time < 5 seconds
##########################################################

### 1 - Housekeeping ----

# Load packages
library(magrittr)
library(dplyr)
library(readr)
library(tidyr)
library(stringr)
library(janitor)
library(readxl)
library(lubridate)
library(tidylog)
library(fuzzyjoin)
library(glue)

# use date in filename
filename <- "Feb-20"

#use one of these to define path
location <- file.path("F:", "PHI", "Publications", "Governance", "Pre Announcement", "Outputs", "2020")
location <- "F:/PHI/Publications/Governance/Pre Announcement/Outputs/2020/"

# Set limit for showing full date
date_limit <- "2020-04-01"


### 2 check new and changed dates ----

#this comes from a different ProcXed report
#load publications with changed dates
#clean col names, make new columns to parse dates for checking
date_changed <- read_xlsx(glue("{location}/ISD Dates Changed ",
                               "{filename}.xlsx"), col_types = "text") %>% 
                clean_names() %>% 
                select(-producer_organisation, -statistics_type, -url_address) %>%   
                mutate(prev = if_else(is.na(ymd(previous_publication_date)), 
                                      dmy(paste0("01-", previous_publication_date)),
                                      ymd(previous_publication_date))) %>% 
                mutate(curr = if_else(is.na(ymd(current_publication_date)), 
                                      dmy(paste0("01-", current_publication_date)),
                                      ymd(current_publication_date)))

#has date within date limit changed?
date_changed %>% filter(prev < date_limit)

#is new date in the past?
date_changed %>% filter(curr < now())

#new pubs within date limit?
date_changed %>% filter(previous_publication_date == "NEWLY ADDED" &
                          current_publication_date < date_limit)

#any issues here might need followed up

### 3 Read in Data ----

# Read the data - these are files from ProcXed report
# Rename column names for Access database
new_pubs <- read_xlsx(glue("{location}/ISD Forthcoming Publications ",
                           "{filename}.xlsx"), col_types = "text") %>%
            rename(DatePublished = `Publication Date`, 
                   Title = `Publication Series`, 
                   Synopsis = `Synopsis`) %>% 
            select(DatePublished, Title, Synopsis)

### 4 Formatting Dates ----

# Change all dates into dmy format
# If date doesn't parse as dmy (e.g. those in MM-YYY format) then add "01-" to the 
# start of the date and parse again, otherwise just parse
new_pubs %<>% 
  mutate(DatePublished = if_else(is.na(dmy(DatePublished)), 
                                    dmy(paste0("01-", DatePublished)),
                                    dmy(DatePublished)))

# If DatePublished is before date_limit then show full date
# If not then show DatePublished as the first of the month
new_pubs %<>% 
  mutate(DatePublished = if_else(DatePublished < date_limit, 
                                    DatePublished, 
                                    floor_date(DatePublished, unit = "month")))

# might not be needed anymore?
# If publication date is after date_limit then add 1 to NotSet
# If publication date is before date_limit then leave blank
new_pubs %<>% 
  mutate(NotSet = if_else(DatePublished < date_limit, "", "1"))

#check all dates are Tuesday, show those that aren't
#will need changed once HPS are included
new_pubs %>% 
  mutate(day_of_week = wday(DatePublished, label = TRUE)) %>% 
  filter(DatePublished < date_limit & day_of_week != "Tue")

######### remember to rerun script to add new topics ##############

# Process the title to get rid of the edition (e.g. October 2020 release)
# If only 1 comma, then take text before comma
# If more than 1 comma, then flag title as needing manual clean up

new_pubs %<>% 
      mutate(comma_count = str_count(Title, ",")) %>%
      separate(Title, sep = ",", into = c("series", "edition"), remove = FALSE) %>%
      mutate(Title = if_else(comma_count == 1, series, Title),
              title_flag = if_else(comma_count > 1, "TRUE", "")) %>% 
      select(-c(comma_count, series, edition))

### 7 General Formatting ----

# Add rescheduled_to and revised columns as NA
# these will need to be manually adjusted
# Add html tags for synopsis on website
# Change date format to dd/mm/yy
# Rearrange columns

new_pubs %<>% 
  mutate(RescheduledTo = NA, 
         Revised = NA, 
         Synopsis = glue("<p>{Synopsis}</p>"), 
         DatePublished = format(DatePublished, "%d/%m/%y")) %>% 
  select(title_flag, DatePublished, Title, Synopsis, 
         RescheduledTo, Revised, NotSet)


### 8 Add Rescheduled Publications ----

# Remember to add reason for delay in Synopsis
# There should be an entry already for the rescheduled date
# ProcXed can run reports to show publications where dates 
# have changed to highlight these

rescheduled_pubs <- tibble(
          DatePublished = "DD/MM/YYYY",
          Title = "title_here",
          HealthTopic = "topic_here",
          Synopsis = "<p>synopsis_here</p>",
          ContactName = paste("analyst_1 tel. 123 456", 
                               "analyst_2 tel. 123 456"),
          RescheduledTo = "DD/MM/YYYY",
          Revised = NA, 
          NotSet = NA)

# Bind resched pubs
new_pubs <- bind_rows(new_pubs, rescheduled_pubs)


### 9 Remove Duplicate ED Weekly Publications ----

# Make a key of DatePublished, Title and Synopsis
# Keep only the distinct ones
# Should be only weekly ED ones from later months

# This shows what will be removed, can check only ED weekly
dups <- new_pubs %>%
        mutate(dupkey = paste0(DatePublished, Title, Synopsis)) %>% 
        mutate(is_dup = duplicated(dupkey)) %>%
        filter(is_dup == TRUE)

# Remove these rows
new_pubs %<>%
  mutate(dupkey = paste0(DatePublished,Title, Synopsis)) %>%
  distinct(dupkey, .keep_all = TRUE) %>%
  select(-dupkey)

### 10 Save as .csv ----

# Export file as .csv
# Use write.csv as write_csv sometimes gives weird encoding in access

write.csv(new_pubs, 
          glue("{location}/ISD_{filename}_processed_xl.csv"), 
          na = "",
          row.names = FALSE)

### END OF SCRIPT ###
