# getting the data out of the excel

library(tidyverse)
library(readxl)

#Fill in the path to your excel spreadsheet
path.to.excel <- './excel_layouts/example.xlsx'

layout <- function(path) {

sheet_names <- readxl::excel_sheets(path)

info.table <- list()
for (sheet_name in sheet_names) {
  df <- readxl::read_xlsx(path, sheet = sheet_name, col_names = FALSE)
  df <- df |>
    mutate(mock = 'MOCK') |>
    pivot_longer(cols = -mock,
                 values_to = sheet_name,
                 names_to ='Number') |>
    select(sheet_name)
  info.table[[sheet_name]] <- df
}
layout.table <- do.call(cbind, info.table) 
layout.table <<- filter_all(layout.table, all_vars(. != 'NA'))
}

layout(path = path.to.excel)
layout.table
