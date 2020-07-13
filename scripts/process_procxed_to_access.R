##########################################################
# Process ProcXed to Access
# Graeme Gowans
# October 2019
# Latest update author - Graeme
# Latest update date - July 2020
# updates for new website

# R Studio Desktop
# 3.6.1
# Description of content
# processes output from ProcXed monthly report to format 
# required for forthcoming page on PHS site
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
library(glue)

# use date in filename
filename <- "Aug_20"

#use one of these to define path
#location <- file.path("F:", "PHI", "Publications", "Governance", "Pre Announcement", "Outputs", "2020", "2020-05-May")
location <- "F:/PHI/Publications/Governance/Pre Announcement/Outputs/2020/2020-08-Aug"

# Set limit for showing full date
date_limit <- "2020-10-01"

### 2 check new and changed dates ----

#this comes from a different ProcXed report
#load publications with changed dates
date_changed <- read_xlsx(glue("{location}/ISD_Dates_Changed_",
                               "{filename}.xlsx"), col_types = "text") %>% 
                clean_names() %>% 
                select(-producer_organisation, -statistics_type, -url_address) %>% 
                mutate(prev = if_else(is.na(ymd(previous_publication_date)), 
                                      dmy(paste0("01-",previous_publication_date)),
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
#these are still called ISD for now - need separate report
new_pubs <- read_xlsx(glue("{location}/ISD_Forthcoming_Publications_",
                           "{filename}.xlsx"), col_types = "text") %>%
                      rename(DatePublished = `Publication Date`, 
                            Title = `Publication Series`, 
                            Synopsis = `Synopsis`, 
                            ContactName = `Contact Details`) %>% 
                      select(DatePublished, Title, Synopsis)

#check for encoding problems
new_pubs %>% mutate(encd = Encoding(Synopsis)) %>% filter(encd != "unknown")
new_pubs %>% mutate(encd = Encoding(Title)) %>% filter(encd != "unknown")

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

# If publication date is after date_limit then add 1 to NotSet
# If publication date is before date_limit then leave blank
new_pubs %<>% 
  mutate(NotSet = if_else(DatePublished < date_limit, "", "1"))

#check all dates are Tuesday, show those that aren't
new_pubs %>% 
  mutate(day_of_week = wday(DatePublished, label = TRUE)) %>% 
  filter(DatePublished < date_limit & day_of_week != "Tue")


### 5 Formatting titles ----

#replace known problems
#HPV vaccination is only one
new_pubs <- new_pubs %>% 
              mutate(Title = str_replace(Title,
                     pattern = "HPV Vaccination Statistics for Men who have Sex with Men, Scotland, HPV in MSM - October 2020",
                     replacement = "HPV Vaccination Statistics for Men who have Sex with Men, Scotland"))

# Process the title to get rid of the edition (e.g. October 2020 release)
# If only 1 comma, then take text before comma
# If more than 1 comma, then flag title as needing manual clean up
new_pubs <- new_pubs %>% 
              mutate(comma_count = str_count(Title, ",")) %>%
              separate(Title, sep = ",", into = c("series", "edition"), remove = FALSE) %>%
              mutate(Title = if_else(comma_count == 1, series, Title),
                      title_flag = if_else(comma_count > 1, "TRUE", "")) %>% 
              select(-c(comma_count, series, edition))

#any problems? if only above can delete
new_pubs %>% filter(title_flag == "TRUE")

if(nrow(filter(new_pubs, title_flag == "TRUE")) == 0) {
new_pubs <- select(new_pubs, -title_flag)
}

### 6 Remove Duplicate ED Weekly Publications ----

# Make a key of DatePublished, Title and Synopsis
# Keep only the distinct ones
# Should be only weekly ED ones from later months

#generate key
new_pubs <- new_pubs %>%
              mutate(dupkey = glue("{DatePublished}_{Title}_{Synopsis}"))

#show duplicates
new_pubs %>% filter(duplicated(dupkey))

# Remove these rows
new_pubs %<>%
  distinct(dupkey, .keep_all = TRUE) %>%
  select(-dupkey)

### 7 general formatting ----

#check formatting - should all be "unknown" so this should return any issues
#if so, go change in excel or procxed
new_pubs %>% mutate(encd = Encoding(Synopsis)) %>% filter(encd != "unknown")
new_pubs %>% mutate(encd = Encoding(Title)) %>% filter(encd != "unknown")  

#add tag to synopsis and change date
new_pubs <- new_pubs %>% 
  mutate(Synopsis = glue("<p>{Synopsis}</p>"), 
         DatePublished = format(DatePublished, "%d/%m/%y"))

### 8 Save as .csv ----

# Export file as .csv
# Use write.csv as write_csv sometimes gives weird encoding in access

date_to_use <- format(today(), "%Y_%m_%d")

write.csv(new_pubs,
          file = glue("{location}/ISD_{filename}_processed_beta_{date_to_use}.csv"), 
          na = "",
          row.names = FALSE)

### END OF SCRIPT ###
