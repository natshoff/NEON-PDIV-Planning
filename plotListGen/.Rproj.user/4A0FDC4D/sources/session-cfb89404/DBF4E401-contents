# This script creates plot level lists for each D08 site.
#
# It uses two datasets: (1) NEON portal DIV records and
# (2) the d08 master species list (d08_plants.csv) from the PDIV SSL page.


# Libraries
library(tidyverse)
library(neonUtilities)
library(openxlsx)

# Read in master species list from SSL
d08_com <- read.csv("Data/d08_plants.csv", header = T) %>% 
  select(taxonID, family, genus, specificEpithet, infraspecificEpithet,
         taxonRank, vernacularName, nativeStatusCode) %>% 
  mutate(scientificName = paste(genus, specificEpithet, infraspecificEpithet))


########
# LENO #
########

# Load PDIV data from D08 - LENO
# Don't pull earlier than 2017 - can't rely on it
leno_div <- loadByProduct(dpID = "DP1.10058.001", site = "LENO", 
                          startdate = "2017-01", include.provisional = T,
                          check.size = F)

# Extract 10m and 100m data
leno_10m100m <- leno_div[["div_10m2Data100m2Data"]] %>% 
  select(plotID, subplotID, eventID, taxonID, family)
# Extract 1m data and filter for only plants
leno_1m <- leno_div[["div_1m2Data"]] %>% 
  filter(divDataType == "plantSpecies") %>% 
  select(plotID, subplotID, eventID, taxonID, family)
# Merge to master file by plot
leno_full <- left_join(leno_10m100m,leno_1m) %>% 
  # add common name
  left_join(d08_com) %>% 
  # remove duplicates across years
  distinct(plotID, taxonID, .keep_all = T) %>% 
  select(plotID, taxonID, scientificName, vernacularName, family, nativeStatusCode)

# list of all plot names
leno_plots <- leno_full %>% 
  distinct(plotID) %>% 
  as.list() %>% 
  unlist()

# list of unique species at LENO
# Merge to master file by plot
leno_spec <- leno_full %>% 
  # remove duplicates across years
  distinct(taxonID, .keep_all = T) %>% 
  select(taxonID, scientificName, vernacularName, family, nativeStatusCode)
# write to csv
write.csv(leno_spec, "leno_spec.csv", row.names = F)


## generate per plot excel files
for (i in 1:length(leno_plots)) {
  out <- leno_full %>%
    filter(plotID == leno_plots[i]) %>%
    select(taxonID, scientificName, vernacularName, family, nativeStatusCode) %>%
    arrange(taxonID)
  
  wb <- createWorkbook(leno_plots[i])
  addWorksheet(wb, leno_plots[i])
  writeData(wb, sheet = 1, out)
  
  # Get the number of rows in the current output
  num_rows <- nrow(out)
  num_cols <- ncol(out)  # Get the number of columns
  
  # Bold for header
  hs <- createStyle(textDecoration = "bold")
  addStyle(wb, sheet = 1, hs, rows = 1, cols = 1:num_cols, gridExpand = FALSE)
  
  # Make header row light grey
  header_style <- createStyle(fgFill = "#D3D3D3")  # Light grey fill
  addStyle(wb, sheet = 1, header_style, rows = 1, cols = 1:num_cols, gridExpand = TRUE)
  
  # Border styles
  bs1 <- createStyle(border = c("top", "bottom"),
                     borderStyle = openxlsx_getOp("borderStyle", "medium"))
  bs2 <- createStyle(border = c("left", "right"),
                     borderStyle = openxlsx_getOp("borderStyle", "thin"))
  addStyle(wb, sheet = 1, bs1, rows = 2:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet = 1, bs2, rows = 1:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  
  # Italic style for scientific name (column 2, check if it's character)
  ts <- createStyle(textDecoration = "italic", wrapText = TRUE)
  scientific_name_col <- which(names(out) == "scientificName")  # Find the column for scientificName
  addStyle(wb, sheet = 1, ts, rows = 2:num_rows + 1, cols = scientific_name_col, 
           gridExpand = TRUE, stack = TRUE)
  
  # Apply conditional formatting for text (e.g., for vernacularName or family)
  for (col in 1:4) {
    is_text_col <- sapply(out[[col]], is.character)
    if (any(is_text_col)) {
      ts <- createStyle(textDecoration = "italic", wrapText = TRUE)
      addStyle(wb, sheet = 1, ts, rows = which(is_text_col) + 1, cols = col, 
               gridExpand = TRUE, stack = TRUE)
    }
  }
  
  # Conditional formatting for nativeStatusCode column (column 5)
  native_status_col <- which(names(out) == "nativeStatusCode")  # Find the column for nativeStatusCode
  
  # Light yellow fill for 'I'
  yellow_fill <- createStyle(fgFill = "#FFFF99")  # Light yellow
  # Light red fill for 'STATE' or 'FEDERAL'
  red_fill <- createStyle(fgFill = "#FFCCCB")  # Light red
  
  # Apply conditional formatting
  for (row in 2:num_rows + 1) {
    # Ensure no NA values before checking the conditions
    if (!is.na(out[row - 1, native_status_col])) {
      if (out[row - 1, native_status_col] == "I") {
        addStyle(wb, sheet = 1, yellow_fill, rows = row, cols = native_status_col, gridExpand = TRUE)
      } else if (out[row - 1, native_status_col] %in% c("STATE", "FEDERAL")) {
        addStyle(wb, sheet = 1, red_fill, rows = row, cols = native_status_col, gridExpand = TRUE)
      }
    }
  }
  
  # Add a thick border around the entire table (around all data rows and columns)
  thick_border <- createStyle(border = c("top", "bottom", "left", "right"),
                              borderStyle = openxlsx_getOp("borderStyle", "thick"))
  addStyle(wb, sheet = 1, thick_border, rows = 1:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  
  # Set column widths
  setColWidths(wb, sheet = 1, cols = c(1, 2, 3, 4), widths = c("auto", 15, "auto", "auto"))
  
  # Export files (CHANGE OUT FOLDER HERE)
  saveWorkbook(wb, paste0("Data/LENO/", leno_plots[i], ".xlsx"), overwrite = TRUE)
}

# Determine plots with the most species
leno_count <- leno_full %>% 
  group_by(plotID) %>% 
  count() %>% 
  rename(speciesCount = n)
write.csv(leno_count, "Data/LENO/leno_count.csv",
          row.names = F)


########
# TALL #
########

# Load PDIV data from D08 - TALL
# Don't pull earlier than 2017 - can't rely on it
tall_div <- loadByProduct(dpID = "DP1.10058.001", site = "TALL", 
                          startdate = "2017-01", include.provisional = T,
                          check.size = F)

# Extract 10m and 100m data
tall_10m100m <- tall_div[["div_10m2Data100m2Data"]] %>% 
  select(plotID, subplotID, eventID, taxonID, family)
# Extract 1m data and filter for only plants
tall_1m <- tall_div[["div_1m2Data"]] %>% 
  filter(divDataType == "plantSpecies") %>% 
  select(plotID, subplotID, eventID, taxonID, family)
# Merge to master file by plot
tall_full <- left_join(tall_10m100m,tall_1m) %>% 
  # add common name
  left_join(d08_com) %>% 
  # remove duplicates across years
  distinct(plotID, taxonID, .keep_all = T) %>% 
  select(plotID, taxonID, scientificName, vernacularName, family, nativeStatusCode)

# list of all plot names
tall_plots <- tall_full %>% 
  distinct(plotID) %>% 
  as.list() %>% 
  unlist()

# list of unique species at TALL
# Merge to master file by plot
tall_spec <- tall_full %>% 
  # remove duplicates across years
  distinct(taxonID, .keep_all = T) %>% 
  select(taxonID, scientificName, vernacularName, family, nativeStatusCode)
# write to csv
write.csv(tall_spec, "tall_spec.csv", row.names = F)

## generate per plot excel files
for (i in 1:length(tall_plots)) {
  out <- tall_full %>%
    filter(plotID == tall_plots[i]) %>%
    select(taxonID, scientificName, vernacularName, family, nativeStatusCode) %>%
    arrange(taxonID)
  
  wb <- createWorkbook(tall_plots[i])
  addWorksheet(wb, tall_plots[i])
  writeData(wb, sheet = 1, out)
  
  # Get the number of rows in the current output
  num_rows <- nrow(out)
  num_cols <- ncol(out)  # Get the number of columns
  
  # Bold for header
  hs <- createStyle(textDecoration = "bold")
  addStyle(wb, sheet = 1, hs, rows = 1, cols = 1:num_cols, gridExpand = FALSE)
  
  # Make header row light grey
  header_style <- createStyle(fgFill = "#D3D3D3")  # Light grey fill
  addStyle(wb, sheet = 1, header_style, rows = 1, cols = 1:num_cols, gridExpand = TRUE)
  
  # Border styles
  bs1 <- createStyle(border = c("top", "bottom"),
                     borderStyle = openxlsx_getOp("borderStyle", "medium"))
  bs2 <- createStyle(border = c("left", "right"),
                     borderStyle = openxlsx_getOp("borderStyle", "thin"))
  addStyle(wb, sheet = 1, bs1, rows = 2:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet = 1, bs2, rows = 1:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  
  # Italic style for scientific name (column 2, check if it's character)
  ts <- createStyle(textDecoration = "italic", wrapText = TRUE)
  scientific_name_col <- which(names(out) == "scientificName")  # Find the column for scientificName
  addStyle(wb, sheet = 1, ts, rows = 2:num_rows + 1, cols = scientific_name_col, 
           gridExpand = TRUE, stack = TRUE)
  
  # Apply conditional formatting for text (e.g., for vernacularName or family)
  for (col in 1:4) {
    is_text_col <- sapply(out[[col]], is.character)
    if (any(is_text_col)) {
      ts <- createStyle(textDecoration = "italic", wrapText = TRUE)
      addStyle(wb, sheet = 1, ts, rows = which(is_text_col) + 1, cols = col, 
               gridExpand = TRUE, stack = TRUE)
    }
  }
  
  # Conditional formatting for nativeStatusCode column (column 5)
  native_status_col <- which(names(out) == "nativeStatusCode")  # Find the column for nativeStatusCode
  
  # Light yellow fill for 'I'
  yellow_fill <- createStyle(fgFill = "#FFFF99")  # Light yellow
  # Light red fill for 'STATE' or 'FEDERAL'
  red_fill <- createStyle(fgFill = "#FFCCCB")  # Light red
  
  # Apply conditional formatting
  for (row in 2:num_rows + 1) {
    # Ensure no NA values before checking the conditions
    if (!is.na(out[row - 1, native_status_col])) {
      if (out[row - 1, native_status_col] == "I") {
        addStyle(wb, sheet = 1, yellow_fill, rows = row, cols = native_status_col, gridExpand = TRUE)
      } else if (out[row - 1, native_status_col] %in% c("STATE", "FEDERAL")) {
        addStyle(wb, sheet = 1, red_fill, rows = row, cols = native_status_col, gridExpand = TRUE)
      }
    }
  }
  
  # Add a thick border around the entire table (around all data rows and columns)
  thick_border <- createStyle(border = c("top", "bottom", "left", "right"),
                              borderStyle = openxlsx_getOp("borderStyle", "thick"))
  addStyle(wb, sheet = 1, thick_border, rows = 1:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  
  # Set column widths
  setColWidths(wb, sheet = 1, cols = c(1, 2, 3, 4), widths = c("auto", 15, "auto", "auto"))
  
  # Export files (CHANGE OUT FOLDER HERE)
  saveWorkbook(wb, paste0("Data/TALL/", tall_plots[i], ".xlsx"), overwrite = TRUE)
}

# Determine plots with the most species
tall_count <- tall_full %>% 
  group_by(plotID) %>% 
  count() %>% 
  rename(speciesCount = n)
write.csv(tall_count, "Data/TALL/tall_count.csv",
          row.names = F)


########
# DELA #
########

# Load PDIV data from D08 - DELA
# Don't pull earlier than 2017 - can't rely on it
dela_div <- loadByProduct(dpID = "DP1.10058.001", site = "DELA", 
                          startdate = "2017-01", include.provisional = T,
                          check.size = F)

# Extract 10m and 100m data
dela_10m100m <- dela_div[["div_10m2Data100m2Data"]] %>% 
  select(plotID, subplotID, eventID, taxonID, family)

# Extract 1m data and filter for only plants
dela_1m <- dela_div[["div_1m2Data"]] %>% 
  filter(divDataType == "plantSpecies") %>% 
  select(plotID, subplotID, eventID, taxonID, family)

# Merge to master file by plot
dela_full <- left_join(dela_10m100m,dela_1m) %>% 
  # add common name
  left_join(d08_com) %>% 
  # remove duplicates across years
  distinct(plotID, taxonID, .keep_all = T) %>% 
  select(plotID, taxonID, scientificName, vernacularName, family, nativeStatusCode)

# list of all plot names
dela_plots <- dela_full %>% 
  distinct(plotID) %>% 
  as.list() %>% 
  unlist()

# list of unique species at DELA
# Merge to master file by plot
dela_spec <- dela_full %>% 
  # remove duplicates across years
  distinct(taxonID, .keep_all = T) %>% 
  select(taxonID, scientificName, vernacularName, family, nativeStatusCode)
# write to csv
write.csv(dela_spec, "dela_spec.csv", row.names = F)



## generate per plot excel files
for (i in 1:length(dela_plots)) {
  out <- dela_full %>%
    filter(plotID == dela_plots[i]) %>%
    select(taxonID, scientificName, vernacularName, family, nativeStatusCode) %>%
    arrange(taxonID)
  
  wb <- createWorkbook(dela_plots[i])
  addWorksheet(wb, dela_plots[i])
  writeData(wb, sheet = 1, out)
  
  # Get the number of rows in the current output
  num_rows <- nrow(out)
  num_cols <- ncol(out)  # Get the number of columns
  
  # Bold for header
  hs <- createStyle(textDecoration = "bold")
  addStyle(wb, sheet = 1, hs, rows = 1, cols = 1:num_cols, gridExpand = FALSE)
  
  # Make header row light grey
  header_style <- createStyle(fgFill = "#D3D3D3")  # Light grey fill
  addStyle(wb, sheet = 1, header_style, rows = 1, cols = 1:num_cols, gridExpand = TRUE)
  
  # Border styles
  bs1 <- createStyle(border = c("top", "bottom"),
                     borderStyle = openxlsx_getOp("borderStyle", "medium"))
  bs2 <- createStyle(border = c("left", "right"),
                     borderStyle = openxlsx_getOp("borderStyle", "thin"))
  addStyle(wb, sheet = 1, bs1, rows = 2:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, sheet = 1, bs2, rows = 1:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  
  # Italic style for scientific name (column 2, check if it's character)
  ts <- createStyle(textDecoration = "italic", wrapText = TRUE)
  scientific_name_col <- which(names(out) == "scientificName")  # Find the column for scientificName
  addStyle(wb, sheet = 1, ts, rows = 2:num_rows + 1, cols = scientific_name_col, 
           gridExpand = TRUE, stack = TRUE)
  
  # Apply conditional formatting for text (e.g., for vernacularName or family)
  for (col in 1:4) {
    is_text_col <- sapply(out[[col]], is.character)
    if (any(is_text_col)) {
      ts <- createStyle(textDecoration = "italic", wrapText = TRUE)
      addStyle(wb, sheet = 1, ts, rows = which(is_text_col) + 1, cols = col, 
               gridExpand = TRUE, stack = TRUE)
    }
  }
  
  # Conditional formatting for nativeStatusCode column (column 5)
  native_status_col <- which(names(out) == "nativeStatusCode")  # Find the column for nativeStatusCode
  
  # Light yellow fill for 'I'
  yellow_fill <- createStyle(fgFill = "#FFFF99")  # Light yellow
  # Light red fill for 'STATE' or 'FEDERAL'
  red_fill <- createStyle(fgFill = "#FFCCCB")  # Light red
  
  # Apply conditional formatting
  for (row in 2:num_rows + 1) {
    # Ensure no NA values before checking the conditions
    if (!is.na(out[row - 1, native_status_col])) {
      if (out[row - 1, native_status_col] == "I") {
        addStyle(wb, sheet = 1, yellow_fill, rows = row, cols = native_status_col, gridExpand = TRUE)
      } else if (out[row - 1, native_status_col] %in% c("STATE", "FEDERAL")) {
        addStyle(wb, sheet = 1, red_fill, rows = row, cols = native_status_col, gridExpand = TRUE)
      }
    }
  }
  
  # Add a thick border around the entire table (around all data rows and columns)
  thick_border <- createStyle(border = c("top", "bottom", "left", "right"),
                              borderStyle = openxlsx_getOp("borderStyle", "thick"))
  addStyle(wb, sheet = 1, thick_border, rows = 1:num_rows + 1, cols = 1:num_cols, 
           gridExpand = TRUE, stack = TRUE)
  
  # Set column widths
  setColWidths(wb, sheet = 1, cols = c(1, 2, 3, 4), widths = c("auto", 15, "auto", "auto"))
  
  # Export files (CHANGE OUT FOLDER HERE)
  saveWorkbook(wb, paste0("Data/DELA/", dela_plots[i], ".xlsx"), overwrite = TRUE)
}

# Determine plots with the most species
dela_count <- dela_full %>% 
  group_by(plotID) %>% 
  count() %>% 
  rename(speciesCount = n)
write.csv(dela_count, "Data/DELA/dela_count.csv",
          row.names = F)