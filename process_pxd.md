process\_pxd\_output
================

These instructions are for taking the monthly ProcXed output and processing to web format.

prepare files for processing
============================

-   finalize ProcXed as usual, submit
-   run ProcXed report "Web version"
-   export as Excel
-   delete first row and delete theme column
-   save at: `F:/PHI/Publications/Governance/Pre Announcement/Outputs/2020/`
-   use filename format, changing month/year: `ISD Forthcoming Publications Feb-20.xlsx`
-   run ProcXed report - "Changes to Publication Date and Newly Added"
-   export as Excel, delete first row
-   save at: `F:/PHI/Publications/Governance/Pre Announcement/Outputs/2020/`
-   use filename format, changing month/year: `ISD Dates Changed Feb-20.xlsx`

set up R
========

-   install R and RStudio
-   download the project from Github: <https://github.com/graemegowans/pxd>
-   click on Green `Clone or Download` &gt; `Download Zip` &gt; choose location &gt; extract
-   you could also clone it if you are already using Git
-   this would let you make changes to the master script
-   navigate to saved project location
-   open project (`pxd.Rproj`) in RStudio
-   in bottom right panel, click on `scripts` then `process_procxed_to_access.R`
-   this will open in the main panel
-   lines of code can be run by highlighting and pressing `Ctrl+Enter`

process file
============

install and load packages:
--------------------------

If this is the first time you have run this, you might need to install some packages. Copy the code below and paste into the Console - lower left panel in RStudio. You will only need to do this once.

``` r
# install packages
install.packages(magrittr)
install.packages(dplyr)
install.packages(readr)
install.packages(tidyr)
install.packages(stringr)
install.packages(janitor)
install.packages(readxl)
install.packages(lubridate)
install.packages(tidylog)
install.packages(fuzzyjoin)
install.packages(glue)
```

Every time you open a new R session, you have to load the packages using `library()` function:

``` r
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
```

define files/dates
------------------

You need to change the file paths to match the file you just created:

``` r
#define file name and path
filename <- "Feb-20"
location <- file.path("//path", "to", "procxed", "output")
```

Usually location will be something like:

``` r
location <- file.path("//F:", "PHI", "Publications", "Governance", "Pre Announcement", "Outputs", "2020")
```

Or you can define in full:

``` r
location <- "F:/PHI/Publications/Governance/Pre Announcement/Outputs/2020/"
```

The `<-` symbol means assign - whatever is on the right side is assigned to the object on the left side - in this case an object called `location` which stores the path to the file. You can check by typing the name into the console to see what is stored in a variable:

``` r
location
```

    ## [1] "F:/PHI/Publications/Governance/Pre Announcement/Outputs/2020/"

Add a date for when to show full or provisional dates on the webpage - e.g. for showing all dates up to the end of March:

``` r
# Set limit for showing full date
date_limit <- "2020-04-01"
```

load health topics
------------------

We have a running list of health topics for matching to publications, this is in the project folder (`data`, should have downloaded with project) - load this into RStudio as:

``` r
# Get lookup table
lookup <- read_csv("data/ISD_lookup_topics.csv",
                   col_types = "cc")
```

Check that it looks ok by typing the name of the object in the console (`lookup`) or clicking the name of the object in the upper right `environment` panel

``` r
lookup
```

    ## # A tibble: 82 x 2
    ##    TitleCode                                        HealthTopic            
    ##    <chr>                                            <chr>                  
    ##  1 Emergency Department Activity & Waiting Times    Waiting-Times          
    ##  2 NHS Performs - Weekly Update of Emergency Depar~ Emergency-Care         
    ##  3 Delayed Discharges in NHSScotland                Health-and-Social-Comm~
    ##  4 NHS Waiting Times                                Waiting-Times          
    ##  5 Allied Health Professionals - Musculoskeletal W~ Waiting-Times          
    ##  6 Prescribing Statistics                           Prescribing-and-Medici~
    ##  7 Cancelled Planned Operations                     Hospital-Care          
    ##  8 Cancer Waiting Times                             Waiting-Times          
    ##  9 Chronic Pain Waiting Times                       Waiting-Times          
    ## 10 IVF Waiting Times                                Waiting-Times          
    ## # ... with 72 more rows

check new and changed dates
---------------------------

Can check if there are any issues with changed dates. Load (`read_xlsx()`) the output from the ProcXed report (new and changed dates) using the location and filename prepaerd earlier. Prepare new date columns for checking:

``` r
#load publications with changed dates
date_changed <- read_xlsx(glue("{location}/ISD Dates Changed ",
                               "{filename}.xlsx"), col_types = "text") %>% 
                clean_names() %>% 
                select(-producer_organisation, -statistics_type, -url_address)  
```

    ## select: dropped 3 variables (producer_organisation, statistics_type, url_address)

``` r
#prepare date columns for checking  
date_changed <- date_changed %>% 
                mutate(prev = if_else(is.na(ymd(previous_publication_date)), 
                                      dmy(paste0("01-", previous_publication_date)),
                                      ymd(previous_publication_date))) %>% 
                mutate(curr = if_else(is.na(ymd(current_publication_date)), 
                                      dmy(paste0("01-", current_publication_date)),
                                      ymd(current_publication_date)))
```

    ## mutate: new variable 'prev' with 10 unique values and 38% NA

    ## mutate: new variable 'curr' with 19 unique values and 0% NA

Can check changes to publications as:

``` r
#has date within date limit changed?
date_changed %>% filter(prev < date_limit)
```

    ## filter: removed 28 rows (82%), 6 rows remaining

    ## # A tibble: 6 x 7
    ##   series_name edition_name previous_public~ current_publica~ synopsis
    ##   <chr>       <chr>        <chr>            <chr>            <chr>   
    ## 1 Cancer Sur~ April 2020 ~ 2020-03-31       2020-04-28       Cancer ~
    ## 2 Child Heal~ February 20~ 2019-02-11       2020-02-11       This pu~
    ## 3 Resource A~ January 202~ 2020-01-21       2020-01-28       NHSScot~
    ## 4 Allied Hea~ March 2020 ~ Mar-20           2020-03-17       Waiting~
    ## 5 Cancelled ~ March 2020 ~ Mar-20           2020-03-03       Statist~
    ## 6 Chronic Pa~ March 2020 ~ Mar-20           2020-03-10       Update ~
    ## # ... with 2 more variables: prev <date>, curr <date>

``` r
#is new date in the past?
date_changed %>% filter(curr < now())
```

    ## filter: removed 33 rows (97%), one row remaining

    ## # A tibble: 1 x 7
    ##   series_name edition_name previous_public~ current_publica~ synopsis
    ##   <chr>       <chr>        <chr>            <chr>            <chr>   
    ## 1 Resource A~ January 202~ 2020-01-21       2020-01-28       NHSScot~
    ## # ... with 2 more variables: prev <date>, curr <date>

``` r
#new pubs within date limit?
date_changed %>% filter(previous_publication_date == "NEWLY ADDED" &
                          current_publication_date < date_limit)
```

    ## filter: removed 30 rows (88%), 4 rows remaining

    ## # A tibble: 4 x 7
    ##   series_name edition_name previous_public~ current_publica~ synopsis
    ##   <chr>       <chr>        <chr>            <chr>            <chr>   
    ## 1 ScotPHO we~ ScotPHO web~ NEWLY ADDED      2020-03-31       The Sco~
    ## 2 Scottish A~ March 2020   NEWLY ADDED      2020-03-31       Scottis~
    ## 3 Child and ~ March 2020 ~ NEWLY ADDED      2020-03-03       Describ~
    ## 4 Workforce ~ March 2020 ~ NEWLY ADDED      2020-03-03       Describ~
    ## # ... with 2 more variables: prev <date>, curr <date>

Any problems here might need followed up on.

read forthcoming publications
-----------------------------

Load (`read_xlsx()`) the list of forthcoming publications saved from ProcXed. This uses the location and filename that we stored earlier. It also renames columns to match the Access files and only keeps (`select()`) the listed columns:

``` r
# Read the data - these are files from ProcXed report
# Rename column names for Access database
new_pubs <- read_xlsx(glue("{location}/ISD Forthcoming Publications ",
                           "{filename}.xlsx")) %>%
            rename(DatePublished = `Publication Date`, 
                   Title = `Publication Series`, 
                   Synopsis = `Synopsis`, 
                   ContactName = `Contact Details`) %>% 
            select(DatePublished, Title, Synopsis, ContactName)
```

    ## select: dropped 3 variables (Frequency, Reason for Change, Publication Type)

Check that it looks ok by typing `new_pubs` into the console or clicking on it in upper right panel.

``` r
new_pubs
```

    ## # A tibble: 178 x 4
    ##    DatePublished Title             Synopsis           ContactName          
    ##    <chr>         <chr>             <chr>              <chr>                
    ##  1 28-Jan-20     Heart Disease St~ Annual update of ~ "Colin Houston\r\nte~
    ##  2 28-Jan-20     Scottish Atlas o~ Scottish Atlas of~ "Graham McGowan\r\nt~
    ##  3 28-Jan-20     Stroke Statistic~ Annual update of ~ "Colin Houston\r\nte~
    ##  4 28-Jan-20     Resource Allocat~ NHSScotland 2022/~ "Edmund Anderson\r\n~
    ##  5 28-Jan-20     NHS Performs - W~ This release of N~ "Kathy McGregor\r\nt~
    ##  6 04-Feb-20     Delayed Discharg~ The publication p~ "Alice Byers\r\ntel.~
    ##  7 04-Feb-20     Scottish Bowel S~ Bowel screening k~ "Garry Hecht\r\ntel.~
    ##  8 04-Feb-20     Cancelled Planne~ Statistics on NHS~ "Greig Stanners\r\nt~
    ##  9 04-Feb-20     Emergency Depart~ Attendances trend~ "Kathy McGregor\r\nt~
    ## 10 04-Feb-20     NHS Performs - W~ This release of N~ "Catriona Haddow\r\n~
    ## # ... with 168 more rows

format publication dates
------------------------

The publication dates are formatted to deal with finalized and provisional dates. This uses the `mutate` function to change the `DatePublished` column and some `lubridate` functions to check date formats. If the date is not in the format Day-Month-Year (e.g. Jan-20 is not), then paste a "01-" at the start and try again. All dates should now be in the format (YYYY-MM-DD), but those that started with provisional dates (MM-YYYY) will all be dated the first of that month.

Example:

``` r
#generate test dates
final <- "27-Jan-2020"
prov <- "June-20"

#check if they are in the right format (can be parsed as DMY)
#this can
dmy(final)
```

    ## [1] "2020-01-27"

``` r
#this one can't
dmy(prov)
```

    ## Warning: All formats failed to parse. No formats found.

    ## [1] NA

``` r
#add "01-" to prov date and try again
prov2 <- paste0("01-", prov)
dmy(prov2)
```

    ## [1] "2020-06-01"

Process the DatePublished column and show output:

``` r
# Change all dates into dmy format
# If date doesn't parse as dmy (e.g. those in MM-YYY format) then add "01-" to the 
# start of the date and parse again, otherwise just parse
new_pubs %<>% 
  mutate(DatePublished = if_else(is.na(dmy(DatePublished)), 
                                    dmy(paste0("01-", DatePublished)),
                                    dmy(DatePublished)))
```

    ## Warning: 46 failed to parse.

    ## Warning: 68 failed to parse.

    ## Warning: 46 failed to parse.

    ## mutate: converted 'DatePublished' from character to double (0 new NA)

``` r
new_pubs
```

    ## # A tibble: 178 x 4
    ##    DatePublished Title             Synopsis           ContactName          
    ##    <date>        <chr>             <chr>              <chr>                
    ##  1 2020-01-28    Heart Disease St~ Annual update of ~ "Colin Houston\r\nte~
    ##  2 2020-01-28    Scottish Atlas o~ Scottish Atlas of~ "Graham McGowan\r\nt~
    ##  3 2020-01-28    Stroke Statistic~ Annual update of ~ "Colin Houston\r\nte~
    ##  4 2020-01-28    Resource Allocat~ NHSScotland 2022/~ "Edmund Anderson\r\n~
    ##  5 2020-01-28    NHS Performs - W~ This release of N~ "Kathy McGregor\r\nt~
    ##  6 2020-02-04    Delayed Discharg~ The publication p~ "Alice Byers\r\ntel.~
    ##  7 2020-02-04    Scottish Bowel S~ Bowel screening k~ "Garry Hecht\r\ntel.~
    ##  8 2020-02-04    Cancelled Planne~ Statistics on NHS~ "Greig Stanners\r\nt~
    ##  9 2020-02-04    Emergency Depart~ Attendances trend~ "Kathy McGregor\r\nt~
    ## 10 2020-02-04    NHS Performs - W~ This release of N~ "Catriona Haddow\r\n~
    ## # ... with 168 more rows

Then check if the `DatePublished` is before the date limit - if so, then keep it as it is. If not, then display it as the first of that month (`floor_date()` rounds to the nearest month):

``` r
#example of rounding
test_dates <- dmy(c("27-Jan-2020", 
                    "28-Feb-2020",
                    "02-Mar-2020"))
floor_date(test_dates, unit = "month")
```

    ## [1] "2020-01-01" "2020-02-01" "2020-03-01"

Round the publication dates as needed:

``` r
# If DatePublished is before date_limit then show full date
# If not then show DatePublished as the first of the month
new_pubs %<>% 
  mutate(DatePublished = if_else(DatePublished < date_limit, 
                                    DatePublished, 
                                    floor_date(DatePublished, unit = "month")))
```

    ## mutate: changed 72 values (40%) of 'DatePublished' (0 new NA)

If the date is after the date limit then add "1" to a new column called `NotSet` - needed for website.

``` r
# If publication date is after date_limit then add 1 to NotSet
# If publication date is before date_limit then leave blank
new_pubs %<>% 
  mutate(NotSet = if_else(DatePublished < date_limit, "", "1"))
```

    ## mutate: new variable 'NotSet' with 2 unique values and 0% NA

Check if any of the publication dates up to the date limit are not Tuesday. These are probably typos and might need changed in ProcXed.

``` r
#check all dates are Tuesday, show those that aren't
new_pubs %>% 
  mutate(day_of_week = wday(DatePublished, label = TRUE)) %>% 
  filter(DatePublished < date_limit & day_of_week != "Tue")
```

    ## mutate: new variable 'day_of_week' with 7 unique values and 0% NA

    ## filter: removed all rows (100%)

    ## # A tibble: 0 x 6
    ## # ... with 6 variables: DatePublished <date>, Title <chr>, Synopsis <chr>,
    ## #   ContactName <chr>, NotSet <chr>, day_of_week <ord>

clean up contact details
------------------------

The website doesn't show email addresses so these are removed:

``` r
# Remove the "\r\n" prefix from telephone numbers
# Remove email addresses
# Remove remaining "\r\n" text

new_pubs %<>% 
  mutate(ContactName = gsub("\r\ntel.", " tel.", ContactName)) %>% 
  mutate(ContactName = gsub("\r\ne-mail: (?:.*?)\r\n", "", ContactName)) %>% 
  mutate(ContactName = gsub("\r\n", " ", ContactName))
```

    ## mutate: changed 178 values (100%) of 'ContactName' (0 new NA)
    ## mutate: changed 178 values (100%) of 'ContactName' (0 new NA)

    ## mutate: changed 121 values (68%) of 'ContactName' (0 new NA)

add health topics
-----------------

The `lookup` object is used to merge health topics to publication title. Uses a `fuzzy_left_join()` with `str_detect` to account for cases where there are dates in the title and therefore not a perfect match.

``` r
# Remove brackets from (CAMHS) for matching
new_pubs %<>% mutate(Title = str_replace_all(Title, "\\(CAMHS\\)", "CAMHS"))
```

    ## mutate: changed 5 values (3%) of 'Title' (0 new NA)

``` r
# Fuzzy join by detecting string name
new_pubs %<>% fuzzy_left_join(y = lookup, 
                              by = c("Title" = "TitleCode"), 
                              match_fun = str_detect)

# Replace brackets for CAMHS
new_pubs %<>% mutate(Title = str_replace_all(Title, "CAMHS", "\\(CAMHS\\)"))
```

    ## mutate: changed 5 values (3%) of 'Title' (0 new NA)

The number of pubs in each area are counted - if there are any NA then these need added to the lookup file.

``` r
# count numbers per topic
count(new_pubs, HealthTopic) %>% print(n = Inf)
```

    ## count: now 19 rows and 2 columns, ungrouped

    ## # A tibble: 19 x 2
    ##    HealthTopic                          n
    ##    <chr>                            <int>
    ##  1 Cancer                              12
    ##  2 Child-Health                         9
    ##  3 Drug-and-Alcohol-Misuse              3
    ##  4 Emergency-Care                      38
    ##  5 Finance                              1
    ##  6 General-Practice                     2
    ##  7 Health-and-Social-Community-Care    17
    ##  8 Heart-Disease                        3
    ##  9 Hospital-Care                       11
    ## 10 Maternity-and-Births                 3
    ## 11 Mental-Health                        3
    ## 12 Prescribing-and-Medicines           11
    ## 13 Primary-Care                         2
    ## 14 Public-Health                        7
    ## 15 Quality-Indicators                  12
    ## 16 Sexual-Health                        3
    ## 17 Stroke                               2
    ## 18 Waiting-Times                       36
    ## 19 Workforce                            3

If any are in `NA` group, you can see what they are:

``` r
######### MANUALLY FIX MISSING HealthTopic VALUES ##############
new_pubs %>% filter(is.na(HealthTopic)) %>% select(Title)
```

These can be added within R and a new version of the lookup file saved. Use the series only, not the edition (the part with the date), so that it doesn't need added every release:

``` r
# add this in to the lookup file and rerun from loading lookup
#new topics to add
topic_to_add <- tibble(TitleCode = "series to add",
                       HealthTopic = "Health-Topic")

#add rows
lookup <- bind_rows(lookup, topic_to_add)

#overwrite lookup table for next time
write_csv(lookup, "data/ISD_lookup_topics.csv")
```

You then need to rerun from section 3 (loading forthcoming pubs) to match these new health topics.

clean up titles
---------------

Titles usually also have edition (e.g. October 2020 release) and these should be removed. This counts the number of commas in the title - if there is one, then the title is split here and only the first part kept. If there is more than 1, then the title is returned unchanged and a `title_flag` column added to highlight manual cleanup is needed.

``` r
# Process the title to get rid of the edition (e.g. October 2020 release)
# If only 1 comma, then take text before comma
# If more than 1 comma, then flag title as needing manual clean up

new_pubs %<>% 
      mutate(comma_count = str_count(Title, ",")) %>%
      separate(Title, sep = ",", into = c("series", "edition"), remove = FALSE) %>%
      mutate(Title = if_else(comma_count == 1, series, Title),
              title_flag = if_else(comma_count > 1, "TRUE", "")) %>% 
      select(-c(comma_count, series, edition))
```

    ## mutate: new variable 'comma_count' with 2 unique values and 0% NA

    ## Warning: Expected 2 pieces. Additional pieces discarded in 1 rows [153].

    ## mutate: changed 177 values (99%) of 'Title' (0 new NA)

    ##         new variable 'title_flag' with 2 unique values and 0% NA

    ## select: dropped 3 variables (series, edition, comma_count)

general formatting
------------------

This adds extra columns for revised and rescheduled publications, adds html tags (
<p>
&
</p>
) to the synopsis, reformats the date and rearranges columns.

``` r
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
```

    ## mutate: converted 'DatePublished' from double to character (0 new NA)

    ##         changed 178 values (100%) of 'Synopsis' (0 new NA)

    ##         new variable 'RescheduledTo' with one unique value and 100% NA

    ##         new variable 'Revised' with one unique value and 100% NA

    ## select: dropped one variable (TitleCode)

adding rescheduled publications
-------------------------------

Publications that have been rescheduled need two entries in the list to ensure that they appear at the original and new dates. The new date will already be in the ProcXed output, but the original date needs added back in. The this is added to the new\_pubs object using `bind_rows()`. Remember to add a reason for delay in the synopsis.

``` r
# Remember to add reason for delay in Synopsis
# There should be an entry already for the rescheduled date
# ProcXed can run reports to show publications where dates 
# have changed to highlight these

rescheduled_pubs1 <- 
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
```

remove duplicates
-----------------

Weekly publications will have multiple entries per month but only want to display one of these on the website. To indentify these, make a key from Date, Title and Synopsis then extract the duplicated ones to check:

``` r
# Make a key of DatePublished, Title and Synopsis
# Keep only the distinct ones
# Should be only weekly ED ones from later months

# This shows what will be removed, can check only ED weekly
dups <- new_pubs %>%
        mutate(dupkey = paste0(DatePublished, Title, Synopsis)) %>% 
        mutate(is_dup = duplicated(dupkey)) %>%
        filter(is_dup == TRUE)
```

    ## mutate: new variable 'dupkey' with 158 unique values and 0% NA

    ## mutate: new variable 'is_dup' with 2 unique values and 0% NA

    ## filter: removed 158 rows (89%), 20 rows remaining

``` r
dups
```

    ## # A tibble: 20 x 11
    ##    title_flag DatePublished Title HealthTopic Synopsis ContactName
    ##    <chr>      <chr>         <chr> <chr>       <glue>   <chr>      
    ##  1 ""         01/04/20      NHS ~ Emergency-~ <p>This~ Catriona H~
    ##  2 ""         01/04/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ##  3 ""         01/04/20      NHS ~ Emergency-~ <p>This~ Catriona H~
    ##  4 ""         01/05/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ##  5 ""         01/05/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ##  6 ""         01/05/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ##  7 ""         01/06/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ##  8 ""         01/06/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ##  9 ""         01/06/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 10 ""         01/06/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 11 ""         01/07/20      NHS ~ Emergency-~ <p>This~ Catriona H~
    ## 12 ""         01/07/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 13 ""         01/07/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 14 ""         01/08/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 15 ""         01/08/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 16 ""         01/08/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 17 ""         01/09/20      NHS ~ Emergency-~ <p>This~ Catriona H~
    ## 18 ""         01/09/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 19 ""         01/09/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## 20 ""         01/09/20      NHS ~ Emergency-~ <p>This~ Kathy McGr~
    ## # ... with 5 more variables: RescheduledTo <lgl>, Revised <lgl>,
    ## #   NotSet <chr>, dupkey <chr>, is_dup <lgl>

These can then be removed from `new_pubs`:

``` r
# Remove these rows
new_pubs %<>%
  mutate(dupkey = paste0(DatePublished,Title, Synopsis)) %>%
  distinct(dupkey, .keep_all = TRUE) %>%
  select(-dupkey)
```

    ## mutate: new variable 'dupkey' with 158 unique values and 0% NA

    ## distinct: removed 20 rows (11%), 158 rows remaining

    ## select: dropped one variable (dupkey)

save output
-----------

### csv file

The `new_pubs` object is then saved as a csv file to the same location as the ProcXed output. Use `write.csv()` as `write_csv()` was giving weird encoding issues in Access.

``` r
# Export file as .csv
# Use write.csv as write_csv sometimes gives weird encoding in access
write.csv(new_pubs, 
          glue("{location}/ISD_{filename}_processed_xl.csv"), 
          na = "",
          row.names = FALSE)
### END OF SCRIPT ###
```

### R script

You can save a dated version of the R script to this location as a record of changes that were made to e.g. rescheduled publications and health topics.

Moving to Access format
=======================

Access can't import .csv files. After processing and exporting report, open csv file and:

-   change any flagged titles as needed manually (to get rid of edition)
-   delete the `title_flag` column
-   check/fix any encoding issues - sometimes problems with symbols e.g. -, £, &, ", '
-   save as .xls (Excel 97-03) for loading to Access
-   you can delete the csv file now

Go to: `F:\PHI\Publications\Governance\Pre Announcement\Updating ISD Forthcoming Webpages`

-   rename 'ForthcomingPubs.mdb' to 'ForthcomingPubs\_\[archive\_date\].mdb' (e.g. ForthcomingPubs\_20190127.mdb)
-   open Access
-   click New (`Ctrl+ N`)
-   make 'Blank database' (right hand side of screen)
-   call it ForthcomingPubs and save to `Updating ISD Forthcoming Webpages` folder
-   then: New &gt; Import Table &gt; select the .xls file made above (change file type to Excel if it doesn't appear) &gt; select "has headers" &gt; add to new Table &gt; check it looks ok &gt; add key automatically &gt; choose name = "tblPubs"
-   check rescheduled pubs are added and that all dates are formatted ok
-   save and close
-   upload using FileZilla FTP
-   check website has updated
