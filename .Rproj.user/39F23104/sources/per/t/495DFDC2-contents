
library(tidyverse)
library(data.table)
library(readxl)
library(xlsx)
topic_an  <- read_excel("data/Impact Intelligence - Project Contractor Sample.xlsx", sheet = 4)
topic_metrics  <- read_excel("data/Impact Intelligence - Project Contractor Sample.xlsx", sheet = 3)

setDT(topic_an)

nms_topic <- names(topic_an)
nms_topic_metrics <- names(topic_metrics)

# function to clean name so that it's easier to use for analysis

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
## Clean the names of the data set
topic_an <- nms_clean(topic_an)
topic_metrics <- nms_clean(topic_metrics)
nms_topic_metrics_clean <- names(topic_metrics)

## filter out topics where they are irrelevant
topic_an <- topic_an[!grepl("irrelevant", topic_name)]

## extract topic id so that it's easier to merge
topic_an[, topic_id := str_extract(topic_label, "^\\d{1,3}[^\\.]")]

## delete topic label so that you only have one topic label in the data
topic_an[, topic_label := NULL]

topic_an_ans <- merge(topic_an, topic_metrics, by = "topic_id")

topic_an_ans[, share_voice := round(segment_count/sum(segment_count), 4)]
## order with share voice and select top 5
setorder(topic_an_ans, -share_voice)
topic_an_ans <- topic_an_ans[1:5,]
## Rename to the original names
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

nms_required <- c("Topic Name", "Share of Voice on all Topics",
                  "Average Sentiment", "Average Engagement", "Keywords")

output_df <- topic_an_ans[, ..nms_required]

##make sure you delete the sheet name output first
## Then append the sheet to your excel workbook
write.xlsx(as.data.frame(output_df), 
           file = "data/Impact Intelligence - Project Contractor Sample.xlsx", 
           sheetName="Output",
           append=TRUE,
           row.names = F)
