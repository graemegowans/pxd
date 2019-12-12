##########################################################
# Process ProcXed to Access
# Graeme Gowans
# October 2019
# Latest update author - Calum Purdie
# Latest update date - November 2019
# Remove duplicate ED weekly pubs
# Remove edition from pub titles
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

# Define filepaths
filename <- "Dec-19"
location <- file.path("//path", "to", "procxed", "output")

# Set limit for showing full date
date_limit <- "2020-02-01"

# Get lookup table
lookup <- read_csv("data/ISD_lookup_topics.csv")


### 2 Read in Data ----

# Read the data - these are files from ProcXed report
# Rename column names for Access database
new_pubs <- read_xlsx(glue("{location}/ISD Forthcoming Publications ",
                           "{filename}.xlsx")) %>%
            rename(DatePublished = `Publication Date`, 
                   Title = `Publication Series`, 
                   Synopsis = `Synopsis`, 
                   ContactName = `Contact Details`) %>% 
            select(DatePublished, Title, Synopsis, ContactName)

### 3 Formatting Dates ----

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

# If publication date is after date_limit then add 1 to NotSet
# If publication date is before date_limit then leave blank
new_pubs %<>% 
  mutate(NotSet = if_else(DatePublished < date_limit, "", "1"))

### 4 Remove Emails ----

# Remove the "\r\n" prefix from telephone numbers
# Remove email addresses
# Remove remaining "\r\n" text

new_pubs %<>% 
  mutate(ContactName = gsub("\r\ntel.", " tel.", ContactName)) %>% 
  mutate(ContactName = gsub("\r\ne-mail: (?:.*?)\r\n", "", ContactName)) %>% 
  mutate(ContactName = gsub("\r\n", " ", ContactName))

### 5 Health Topic Lookup ----

# Remove brackets from (CAMHS) for matching
new_pubs %<>% mutate(Title = str_replace_all(Title, "\\(CAMHS\\)", "CAMHS"))

# Fuzzy join by detecting string name
new_pubs %<>% fuzzy_left_join(y = lookup, 
                              by = c("Title" = "TitleCode"), 
                              match_fun = str_detect)

# Replace brackets for CAMHS
new_pubs %<>% mutate(Title = str_replace_all(Title, "CAMHS", "\\(CAMHS\\)"))

# Count numbers
count(new_pubs, HealthTopic) %>% print(n = Inf)

######### MANUALLY FIX MISSING HealthTopic VALUES ##############

new_pubs %>% filter(is.na(HealthTopic)) %>% select(Title)

# Manually add this in to the lookup file and rerun

# Process the title to get rid of the edition (e.g. October 2020 release)
# If only 1 comma, then take text before comma
# If more than 1 comma, then flag title as needing manual clean up

new_pubs %<>% 
      mutate(comma_count = str_count(Title, ",")) %>%
      separate(Title, sep = ",", into = c("series", "edition"), remove = FALSE) %>%
      mutate(Title = if_else(comma_count == 1, series, Title),
              title_flag = if_else(comma_count > 1, "TRUE", "")) %>% 
      select(-c(comma_count, series, edition))

### 6 General Formatting ----

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
  select(title_flag, DatePublished, Title, HealthTopic, Synopsis, ContactName, 
         RescheduledTo, Revised, NotSet)


### 7 Add Rescheduled Publications ----

# Remember to add reason for delay in Synopsis
# There should be an entry already for the rescheduled date
# ProcXed can run reports to show publications where dates 
# have changed to highlight these

rescheduled_pubs <- 
  tibble(DatePublished = "DD/MM/YYYY",
          Title = "title_here",
          HealthTopic = "topic_here",
          Synopsis = "<p>synopsis_here</p>",
          ContactName = paste0("analyst_1 tel. 123 456", 
                               "analyst_2 tel. 123 456"),
          RescheduledTo = "DD/MM/YYYY",
          Revised = NA, 
          NotSet = NA)

# Bind resched pubs
new_pubs <- bind_rows(new_pubs, rescheduled_pubs)


### 8 Remove Duplicate ED Weekly Publications ----

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

### 9 Save as .csv ----

# Export file as .csv
# Use write.csv as write_csv sometimes gives weird encoding in access

write.csv(new_pubs, 
          glue("{location}/ISD_{filename}_processed_xl.csv"), 
          na = "",
          row.names = FALSE)

### END OF SCRIPT ###
