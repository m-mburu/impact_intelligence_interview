README
================

### Load R packages to use

``` r
library(tidyverse)
library(data.table)
library(readxl)
library(xlsx)
library(knitr)
```

### Read data into R

``` r
topic_an  <- read_excel("data/Impact Intelligence - Project Contractor Sample.xlsx", sheet = 4)
topic_metrics  <- read_excel("data/Impact Intelligence - Project Contractor Sample.xlsx", sheet = 3)
```

### Save the names of the two sheets

``` r
# convert to DT
setDT(topic_an)

nms_topic <- names(topic_an)
nms_topic_metrics <- names(topic_metrics)
```

### Function to clean names

-   ie it removes white space and other characters to make it easier to
    use a programming language like R or python etc.

``` r
nms_clean <- function(data_set){
  nms_old <- names(data_set)
  nms_new <- nms_old %>% tolower() 
  nms_new <- gsub("\\s", "_", nms_new)
  nms_new <- gsub("\\.$", "", nms_new)
  nms_new <- gsub("\\(|\\)|%", "", nms_new)
  nms_new <- gsub("/", "", nms_new)
  
  setDT(data_set)
  setnames(data_set, nms_old, nms_new)
  data_set
}
```

### Clean names

``` r
topic_an <- nms_clean(topic_an)
topic_metrics <- nms_clean(topic_metrics)
nms_topic_metrics_clean <- names(topic_metrics)
```

### Some other minor cleaning

``` r
## filter out topics where they are irrelevant
topic_an <- topic_an[!grepl("irrelevant", topic_name)]

## extract topic id so that it's easier to merge
topic_an[, topic_id := str_extract(topic_label, "^\\d{1,3}[^\\.]")]

## delete topic label so that you only have one topic label in the data

topic_an[, topic_label := NULL]

#merge the two data sets
topic_an_ans <- merge(topic_an, topic_metrics, by = "topic_id")

## calculate Share of Voice on all Topics
topic_an_ans[, share_voice := round(segment_count/sum(segment_count), 4)]
## order with share voice and select top 5 descending
setorder(topic_an_ans, -share_voice)
## select top 5
topic_an_ans <- topic_an_ans[1:5,]
## Rename to the original names

head(topic_an_ans) %>%
  kable()
```

| topic_id | topic_name | topic_label                                                                                                                                | segment_count | average_sentiment | average_engagement | overrepresented_keywords                                                                                                                                                                                                                              | share_voice |
|:---------|:-----------|:-------------------------------------------------------------------------------------------------------------------------------------------|--------------:|------------------:|-------------------:|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------:|
| 48       | TopicD     | 48\. weight, calorie, workout, meal plan, healthy, lifestyle, mindset, weight loss, lose weight, coach                                     |            57 |         0.7657323 |           5.068750 | \[‘meal’, ‘eat’, ‘diet’, ‘body’, ‘weight’, ‘fat’, ‘calorie’, ‘lose’, ‘step’, ‘day’, ‘nutrition’, ‘time’, ‘gain’, ‘workout’, ‘gym’, ‘healthy’, ‘health’, ‘meal plan’, ‘goal’, ‘fitness’\]                                                              |      0.1633 |
| 114      | TopicK     | 114\. change, environmental impact, impact, climate change, epi, shropshire, environmental, predict, cop26, climatecrisis                  |            50 |         0.4889069 |           7.218750 | \[‘consumer’, ‘change’, ‘choice’, ‘ask’, ‘environmental impact’, ‘action’, ‘brand’, ‘behaviour’, ‘future’, ‘cop26’, ‘impact’, ‘influence’, ‘habit’, ‘shropshire’, ‘epi’, ‘act’, ‘educate’, ‘population’, ‘climate change’, ‘sustainability’\]         |      0.1433 |
| 107      | TopicL     | 107\. reusable bag, reusable, plastic bag, tote bag, sustainablelive, bsci, tx287, go0367, totebag, carryout                               |            40 |         0.6687746 |          17.743697 | \[‘bag’, ‘reusable bag’, ‘reusable’, ‘shopping’, ‘durable’, ‘tote’, ‘plastic bag’, ‘use’, ‘carry’, ‘grocery store’, ‘cotton’, ‘tote bag’, ‘sustainablelive’, ‘universal’, ‘trolley’, ‘canvas’, ‘shop’, ‘grocery’, ‘trip’, ‘single’\]                  |      0.1146 |
| 93       | TopicB     | 93\. recycling, recycle, soft plastic, recycling bin, recyclable, flexible plastic, cardboard, redcycle, plastic, plasticcycle             |            40 |         0.4521038 |           7.054688 | \[‘recycle’, ‘recycling’, ‘battery’, ‘bin’, ‘plastic’, ‘soft plastic’, ‘film’, ‘recyclable’, ‘accept’, ‘recycling bin’, ‘flexible plastic’, ‘dispose’, ‘collection’, ‘cardboard’, ‘authority’, ‘bag’, ‘paper’, ‘redcycle’, ‘facility’, ‘drop’\]       |      0.1146 |
| 15       | TopicE     | 15\. palm oil, palm, rainforest, deforestation, orangutans, mspo, certified sustainable palm oil, sustainablepalmoil, greenhouse gas, rspo |            32 |         0.0972222 |          15.833333 | \[‘palm oil’, ‘oil’, ‘palm’, ‘product’, ‘rainforest’, ‘destroy’, ‘vegetable’, ‘deforestation’, ‘use’, ‘habitat’, ‘produce’, ‘crop’, ‘sustainable’, ‘mspo’, ‘orangutans’, ‘tree’, ‘kembali’, ‘certified sustainable palm oil’, ‘nutella’, ‘adelaide’\] |      0.0917 |

### Rename back columns to original names

``` r
nms_final_dt <- names(topic_an_ans)

setnames(topic_an_ans,
         nms_topic_metrics_clean, 
         nms_topic_metrics, skip_absent = T)

nms_to_rnm <- c("topic_name", "share_voice",
                "Overrepresented Keywords")

nms_to_rnm_to <- c("Topic Name", 
                   "Share of Voice on all Topics",
                   "Keywords")

setnames(topic_an_ans,
         nms_to_rnm, 
         nms_to_rnm_to, skip_absent = F)
```

### Select required columns

``` r
nms_required <- c("Topic Name", "Share of Voice on all Topics",
                  "Average Sentiment", "Average Engagement", "Keywords")

output_df <- topic_an_ans[, ..nms_required]
```

### Delete output sheet

-   it will enable in writing to it

``` r
wb_path <- "data/Impact Intelligence - Project Contractor Sample.xlsx"
wb <- loadWorkbook(wb_path)
removeSheet(wb, sheetName = "Output")
saveWorkbook(wb, wb_path)
```

### Append the sheet to your excel workbook

``` r
## Then append the sheet to your excel workbook
write.xlsx(as.data.frame(output_df), 
           file = "data/Impact Intelligence - Project Contractor Sample.xlsx", 
           sheetName="Output",
           append=TRUE,
           row.names = F)
```
