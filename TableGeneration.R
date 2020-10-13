### Project to extract BQA data, arrange and analyse in different levels and export analysis to tables and charts for QA visit data packs

oldw <- getOption("warn")
options(warn = -1)

renv::restore()
library('ggplot2')
library('purrr')
library('dplyr')
library('tidyr')
library('openxlsx')

DataExtractAll <- function (filename) {
  # Extracts breast screening pathology data from the specified filename so long as it is a BQA file extracted from NBSS as a CSV file
  #
  # Args:
  #   filename: A string of the path to the file containing the data
  #
  # Returns:
  #   A dataframe containing the relevant rows of the BQA data
  datasource <- read.csv(filename, skip = 3, stringsAsFactors = F)
  datasource <- datasource[-nrow(datasource), 1:23] #remove last row (states end of data), and extra columns with calculated fields
  datasource <- separate(datasource, col = "Primary.Sort.Value", into = PrimarySortIDHeadings, sep = "\\*") #expands the Primary.Sort.Value field 
}

BQACalcFunSum <- function(MeasureName, BoxNumbers) {
  # Creates a dataframe containing summed data from the BQA box numbers input with a Measure column containing the measure name text
  #
  # Args:
  #   MeasureName: The name of the BQA measure to be selected
  #   BoxNumbers: The BQA box numbers to include in the final value
  #
  # Returns:
  #   A dataframe containing the number of cases that match the inputs
  temp <- data.frame()
  for (i in BoxNumbers) {
    data <- pivot_longer(BoxList[[i]], col = 2:ncol(BoxList[[i]]), names_to = "variable")
    temp <- temp %>% 
      bind_rows(data) %>%
      mutate(value = replace_na(value, 0)) %>%
      group_by(Filename, variable) %>%
      summarise(value = sum(value), .groups = "drop")
  }
  temp <- temp %>%
    mutate(Measure = MeasureName) %>%
    pivot_wider(id_cols = c(Measure, Filename), names_from = variable) 
  temp
}

TableGenerator <- function (dataset, filters, colID, ...) {
  # Generates a dataframe filtered by one variable and grouped by user specified variables
  #
  # Args:
  #   dataset: The source dataset
  #   filters: A string specifying the filters to apply to dataset
  #   colID: The column ID's to use in spread
  #   ...: Strings to use to group the filtered data
  #
  # Returns:
  #   A dataframe containing filtered, grouped and summarised data based on the arguments supplied
  groupcols <- enquos(...)
  t <- dataset %>% 
    filter(!!enquo(filters)) %>%
    group_by(!!!groupcols) %>% 
    summarise(value = sum(value, na.rm = T), .groups = "drop") %>%
    spread(!!enquo(colID), value)
  t
}

winDialog(type = "ok", "Please choose the BQA files for analysis")
filessrc <- choose.files() 
filesdir <- dirname(filessrc[1])

### Variable vectors and lists
BCategory <- c("Total", "WBN.B1", "WBN.B2", "WBN.B3.ns", "WBN.B3.wa", "WBN.B3.na", "WBN.B3", "WBN.B4", "WBN.B5c", "WBN.B5b", "WBN.B5a", "WBN.B5")
BoxIDVector <- c("1", "5", "6", "7", "8", "9", "28", "32", "34", "36", "37", "41", "42", "43", "44", "46", "50", "51", "53", "54", "52")
BoxCatIDFilters <- c('WBN.B5', 'WBN.B4', 'WBN.B3', 'WBN.B2', 'WBN.B1', 'Total', 'WBN.B5', 'WBN.B4', 'WBN.B2', 'Total', 'WBN.B5', 'WBN.B4',
                   'WBN.B3', 'WBN.B2', 'WBN.B1', 'WBN.B5', 'WBN.B4', 'WBN.B3', 'WBN.B1', 'Total', 'WBN.B2')
BoxRowIDFilters <- c(10, 10, 10, 10, 10, 10, 40, 40, 40, 40, 60, 60, 60, 60, 60, 9999, 9999, 9999, 9999, 9999, 9999)
subtractNum <- list("28", c("32", "41"), "7")
subtractNumNames <- c("PPV (B5)", "PPV (B4)", "Negative Predictive Value")
calcnames <- c("Absolute Sensitivity", "Specificity (biopsy cases only)", "Specificity (Full)", "Complete Sensitivity", "PPV (B5)", "PPV (B4)", 
               "PPV (B3)", "Negative Predictive Value", "False Negative Rate", "True False Positive Rate", "Miss Rate")
calcNumSumBoxes <- list(c("1", "37"), "34", c("34", "43"), c("1", "5", "6", "37"), "46", "50", "6", "52", "7", "28", c("7", "8"))
calcDenomSumBoxes <- list(c("9", "37"), "36", c("36", "42", "43", "44"), c("9", "37"), "46", "50", "51", "52", c("9", "37"), c("9", "37"),
                          c("9", "37"))
#chartnames <- c("BcatPlot", "B5Plot", "B3Plot")
chartrownums <- list(c(1:5), c(1:3), c(1:3))
chartrows <- list(c("Number of B5", "Number of B4", "Number of B3", "Number of B2", "Number of B1"), 
                  c("Number of B5a", "Number of B5b", "Number of B5c"), 
                  c("Number of B3 with atypia", "Number of B3 without atypia", "Number of B3 with atypia not specified"))
chartdatastring <- c("B Category", "B5 sub-category", "B3 sub-category")
oldfilters <- c("Clients", "Tests")
#newfilters <- c("Path_pseudo_code", "Laboratory_name", "Loc_method", "Radiological_appearance")
TCpicker <- c("C", "T")
imgNum <- 1
imgList <- NULL

### extracts the filenames from filessrc
TableNames <- basename(filessrc)
TableNames <- gsub(pattern = "\\..[^\\.]*$", "", TableNames)

# Uses openxlsx package to create a workbook and set fontsize and font
wb <- createWorkbook()
modifyBaseFont(wb, fontSize = 12, fontName = "Arial")

## Options for C and T data only:
group <- "NBSS pathologist code"
PrimarySortIDHeadings <- "Primary_Sort_Value"

# Extract the data from the file(s)
tidydataset <- lapply(filessrc, DataExtractAll)
tidydataset <- mapply(cbind, tidydataset, "Filename"=TableNames, SIMPLIFY = F)
tidydataset <- tidydataset %>%
  bind_rows() %>%
  select(-Total.Cases.Screened, -Total.Assessed, -Total.WBN.Performed, -Total.VAE.Performed) %>%
  pivot_longer(cols = all_of(BCategory), names_to = "BCategory") %>%
  filter(Table.Identifier..A..B.or.C.=="D")
tidydataset$value[is.na(tidydataset$value)] <- 0
tidydataset$Tests.or.Clients..T.or.C. <- gsub(TRUE, "T", tidydataset$Tests.or.Clients..T.or.C.) # convert TRUE values into T string

# Check for compile for BSIS files ============================================================================

comp4BSIScheck <- tidydataset %>%
  select(Primary_Sort_Value) %>%
  unique()

if (grepl("*", comp4BSIScheck, fixed=T)) {
  rm(list = ls())
  stop(winDialog(type = "ok", "Files selected include compile for BSIS output. Please run the code again and select only normal BQA output"))
}

  
# Pathologist pseudonymisation code===============================================================================

pathList <- tidydataset %>%
  mutate(Primary_Sort_Value = case_when(
    Primary.Sort.ID == "P" ~ Primary_Sort_Value,
    Primary.Sort.ID == "TOT" ~ "TOT"
  )) %>%
  select(Primary_Sort_Value, value) %>%
  group_by(Primary_Sort_Value) %>%
  summarise(value = max(value), .groups = "drop") %>%
  arrange(desc(value), Primary_Sort_Value) %>%
  mutate(rank = 1:nrow(.)-1,
         pathCode = case_when(
           Primary_Sort_Value == "TOT" ~ "Total",
           Primary_Sort_Value == "zzzz" ~ "Unknown Pathologist",
           rank < 10 ~ paste0("PATH0", rank),
           rank >= 10 ~ paste0("PATH", rank)
         )) %>%
  select(-value, -rank)

winDialog(type = "ok", "Please review the list of pathologists and update pathCode as desired [see work instructions for details.\n*** DO NOT ALTER 'Primary_Sort_Value' COLUMN ***")
pathList <- edit(pathList)

# adds a sheet to the workbook with the pathologist NBSS codes linked to the assigned pseudonyms
addWorksheet(wb, "Pathologist Codes")
writeData(wb, 
          sheet = "Pathologist Codes", 
          pathList, 
          rowNames = FALSE) 

# ===================================================================================================================

for (TC in 1:length(unique(tidydataset$Tests.or.Clients..T.or.C.))) { # should handle cases where only C has been run (there will only be 1 loop)
  
  if (TCpicker[TC] == "C") { # if the data is clients, sets row names up for clients
    BCatDesc <- c("Total number of clients", "Number of B1", "Number of B2", "Number of B3 with atypia not specified",
                  "Number of B3 with atypia", "Number of B3 without atypia", "Number of B3", "Number of B4", "Number of B5c", "Number of B5b",
                  "Number of B5a", "Number of B5")
    BCatOrder <- c("Total number of clients", "Number of B5", "Number of B4", "Number of B3", "Number of B2", "Number of B1", "Number of B5a",
                   "Number of B5b", "Number of B5c", "Number of B3 with atypia", "Number of B3 without atypia",
                   "Number of B3 with atypia not specified")
  } else { # if the data is tests, sets row names up for tests
    BCatDesc <- c("Total number of tests", "Number of B1", "Number of B2", "Number of B3 with atypia not specified",
                  "Number of B3 with atypia", "Number of B3 without atypia", "Number of B3", "Number of B4", "Number of B5c", "Number of B5b",
                  "Number of B5a", "Number of B5")
    BCatOrder <- c("Total number of tests", "Number of B5", "Number of B4", "Number of B3", "Number of B2", "Number of B1", "Number of B5a",
                   "Number of B5b", "Number of B5c", "Number of B3 with atypia", "Number of B3 without atypia",
                   "Number of B3 with atypia not specified")
  }
  
  BCatlookup <- data.frame(BCategory, BCatDesc)
  
  tidydatasetuse <- tidydataset %>%
    filter(Tests.or.Clients..T.or.C. == TCpicker[TC]) %>% 
    inner_join(BCatlookup) %>%
    left_join(pathList) %>%
    select(-Primary_Sort_Value) %>%
    rename(Primary_Sort_Value = pathCode) %>%
    mutate(Primary_Sort_Value = replace_na(.$Primary_Sort_Value, "Total")) %>%
    group_by(across(-value)) %>%
    summarise(value = sum(value), .groups = "drop") 
  
  BQAtablescombined <- TableGenerator(tidydatasetuse, Row.Identifier==9999, Primary_Sort_Value, Filename, Primary_Sort_Value, BCatDesc)
  BQAtablescombined <- left_join(data.frame("BCatDesc" = BCatOrder), BQAtablescombined)
  BQAtablescombined$BCatDesc <- as.character(BQAtablescombined$BCatDesc)
  
  # This section produces a list of data frames that contain the values for the BQA box numbers identified by BoxIDVector
  BoxList <- list()
  for (k in 1:length(BoxIDVector)) {
    BoxList[[BoxIDVector[k]]] <- TableGenerator(tidydatasetuse, Row.Identifier==BoxRowIDFilters[k] & BCategory==BoxCatIDFilters[k],
                                                Primary_Sort_Value, Filename, Primary_Sort_Value)
  }
  
  tidynumframe <- data.frame()
  tidydenomframe <- data.frame()
  for (i in seq_along(calcnames)) {
    tidynumframe <- bind_rows(tidynumframe, BQACalcFunSum(calcnames[i], calcNumSumBoxes[[i]]))
    tidydenomframe <- bind_rows(tidydenomframe, BQACalcFunSum(calcnames[i], calcDenomSumBoxes[[i]]))
  }
  
  # For numerators: PPV (B5) need to subtract "28", for PPV (B4) need to subtract c("32","41"), for Negative Predictive Value need to 
  # subtract "7" from relevant sum boxes.
  for (i in seq_along(subtractNumNames)) {
    subframe <- BQACalcFunSum(subtractNumNames[i], subtractNum[[i]])
    common <- intersect(names(tidynumframe), names(subframe)[-c(1:2)])
    tidynumframe[tidynumframe$Measure == subtractNumNames[i], names(tidynumframe) %in% common] <-
      tidynumframe[tidynumframe$Measure == subtractNumNames[i], 
                   names(tidynumframe) %in% common] - replace(subframe[common], is.na(subframe[common]), 0)
  }
  
  # For denominators: PPV (B4) need to subtract 41 from relevant sumbox
  subframe <- BQACalcFunSum("PPV (B4)", "41")
  common <- intersect(names(tidydenomframe), names(subframe)[-c(1:2)])
  tidydenomframe[tidydenomframe$Measure == "PPV (B4)", names(tidydenomframe) %in% common] <-
    tidydenomframe[tidydenomframe$Measure == "PPV (B4)", 
                   names(tidydenomframe) %in% common] - replace(subframe[common], is.na(subframe[common]), 0)
  
  names(tidynumframe)[names(tidynumframe) == "Measure"] <- "BCatDesc"
  names(tidydenomframe)[names(tidydenomframe) == "Measure"] <- "BCatDesc"
  
  # creates lists of the numerator and denominator values separated by filename
  NumList <- list()
  for (k in TableNames) {
    NumList[[k]] <- as_tibble(tidynumframe[tidynumframe$Filename == k, c(1, 3:ncol(tidynumframe))])
  }
  DenomList <- list()
  for (k in TableNames) {
    DenomList[[k]] <- as_tibble(tidydenomframe[tidydenomframe$Filename == k, c(1, 3:ncol(tidydenomframe))])
  }
  
  # clears NA values coerced by bind_rows by replacing with 0
  tidynumframe[is.na(tidynumframe)] <- 0
  tidydenomframe[is.na(tidydenomframe)] <- 0
  # calculates the proportions for the calculated statistics and removes NaN values created by 0/0
  tidycalcframe <- cbind(tidynumframe[,1:2], tidynumframe[, 3:ncol(tidynumframe)] / tidydenomframe[, 3:ncol(tidydenomframe)])
  tidycalcframe[is.na(tidycalcframe)] <- 0
  names(tidycalcframe)[names(tidycalcframe) == "Measure"] <- "BCatDesc"
  tidydenomframemelted <- pivot_longer(tidydenomframe, 3:ncol(tidydenomframe), names_to = "variable", values_to = "denomvalue")
  tidycalcframeperc <-  tidycalcframe %>%
    pivot_longer(3:ncol(tidydenomframe), names_to = "variable") %>%
    mutate(value = paste0(format(value *100, digits = 1),"%")) %>%
    inner_join(tidydenomframemelted) %>%
    mutate(value = case_when(
      denomvalue == 0 ~ "No Cases",
      denomvalue != 0 ~ value
    )) %>%
    select(-denomvalue) %>%
    pivot_wider(values_from = value, names_from = variable, id_cols = c(Filename, BCatDesc))

  CalcList <- list()
  for (k in TableNames) {
    CalcList[[k]] <- as_tibble(tidycalcframeperc[tidycalcframeperc$Filename == k, c(2:ncol(tidycalcframeperc))])
  }
  
  # creates a list of dataframes based on the filename of the data source
  TablesList <- list()
  for (k in TableNames) {
    TablesList[[k]] <- as_tibble(BQAtablescombined[BQAtablescombined$Filename==k, c(1, 3:ncol(BQAtablescombined))])
  }
  
  # Moves the BCatDesc column to the start of the dataframes
  TablesList <- lapply(TablesList, function(df) {select(df, BCatDesc, everything())})
  # sets up list containing the ordered data
  for (i in seq_along(TablesList)) {
    TablesList[[i]] <- TablesList[[i]][, c(1, 1 + order(-as.data.frame(TablesList[[i]][1, 2:ncol(TablesList[[i]])])))]
  }
  
  TotalList <- list()
  for (k in TableNames) {
    TotalList[[k]] <- bind_rows(mutate_all(TablesList[[k]], as.character), CalcList[[k]]) %>% 
     select(where(~!any(is.na(.)))) # removes columns where any NA values exist (removes non-existant pathologists)
  }
  
  for (i in seq_along(TotalList)) {
    DenomList[[i]] <- DenomList[[i]][names(TotalList[[i]])]
    NumList[[i]] <- NumList[[i]][names(TotalList[[i]])]
  }
  
  # creates a list of dataframes per filename containing the proportion of the different B5 categories in place of the numbers of cases
  chartframeList <- list()
  for (h in seq_along(TotalList)) {
    chartframe <- TotalList[[h]] 
    for (i in c("Number of B5a", "Number of B5b", "Number of B5c")){
      chartframe[chartframe$BCatDesc == i, 2:ncol(chartframe)] <-
        as.list(as.character(as.numeric(chartframe[chartframe$BCatDesc == i, 2:ncol(chartframe)]) / 
                               as.numeric(chartframe[chartframe$BCatDesc == "Number of B5", 2:ncol(chartframe)])))
    }
    for (i in c("Number of B3 with atypia", "Number of B3 without atypia", "Number of B3 with atypia not specified")){
      chartframe[chartframe$BCatDesc==i, 2:ncol(chartframe)] <-
        as.list(as.character(as.numeric(chartframe[chartframe$BCatDesc == i,2:ncol(chartframe)]) / 
                               as.numeric(chartframe[chartframe$BCatDesc == "Number of B3", 2:ncol(chartframe)])))
    }
    for (i in c("Number of B5", "Number of B4", "Number of B3", "Number of B2", "Number of B1")) {
      chartframe[chartframe$BCatDesc == i, 2:ncol(chartframe)] <-
        as.list(as.character(as.numeric(chartframe[chartframe$BCatDesc == i,2:ncol(chartframe)]) / 
                               as.numeric(chartframe[chartframe$BCatDesc == BCatDesc[1],2:ncol(chartframe)])))
    }
    chartframeList[[names(TotalList[h])]] <- chartframe
  }
  
  
  # =========================================================================================================================================
  
  # pastes the data in the different lists and the charts into a worksheet in the opened workbook for each filename in TableNames
  for (k in seq_along(TableNames)) {
    
    # Because the sheet name can only be 31 characters max, the filename is truncated to use the last 20 characters in the TableName entry
    # this may be an issue if the filename is not informative enough.
    
    if (nchar(TableNames[k]) > 23) {
      sheetName <- trimws(substr(TableNames[k], 1, 22))
    } else {
      sheetName <- TableNames[k]
    }
    addWorksheet(wb, paste(sheetName, oldfilters[TC]))
    writeData(wb,
              sheet = paste(sheetName, oldfilters[TC]),
              x = TotalList[[k]],
              rowNames = FALSE,
              withFilter = FALSE)
    #creates data set for charts
    chart_data <- chartframeList[[TableNames[k]]]
    chart_data <- chart_data %>%
      pivot_longer(cols = 2:ncol(.)) %>%
      mutate (value = as.numeric(value)) 

    # removes pathologists who have <=5 in total row
    chartdatatotals <- chart_data %>%
      filter(BCatDesc == BCatDesc[1],
             value > 5)
    
    chart_data <- chart_data %>%
      filter(name %in% chartdatatotals$name) %>% # filters chart data to exclude pathologists without at least 5 outcomes
      mutate(BCatDesc = gsub("Number of ", "", .$BCatDesc),
             name = gsub("V1", "Total", .$name)) # removes 'Number of ' tag on description column 
    
    chart_1_data <- chart_data %>%
      filter(BCatDesc %in% c("B1","B2","B3","B4","B5")) %>%
      group_by(name) %>%
      mutate(pos = cumsum(value) - value/2)
    
    ggplot(chart_1_data, aes(y=value, x = reorder(name, -value), fill = BCatDesc)) +
      geom_bar(stat="identity", position = "stack") +
      #geom_text(aes(y = pos, label = format(value*100, digits = 1)), colour = "white") +
      theme_minimal() +
      scale_fill_manual(name = "Category",
                        values = c("B5" = "#822433",
                                   "B4" = "#00B092",
                                   "B3" = "#002776",
                                   "B2" = "#A4AEB5",
                                   "B1" = "#E9994A"))+
      theme(axis.text.x = element_text(angle = 45, hjust = 0.95, size = 12),
            axis.text.y = element_text(size = 12),
            panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(size=.1, color="dark grey"),
            plot.title = element_text(face = "bold", colour = "black", size = 16,hjust = 0.5),
            axis.title.x = element_text(face = "bold", colour = "black", size = 14),
            axis.title.y = element_text(face = "bold", colour = "black", size = 14),
            legend.title = element_blank(), legend.position = "top", legend.spacing.x = unit(0.5, "cm"), 
            legend.text = element_text(size = 12)) + 
      labs(title = paste0("BQA ",oldfilters[TC]," data by B category,\n shown by pathologist with >5 total cses")) +
      scale_y_continuous(name = "Percentage of total", breaks = c(1.00,0.80,0.60,0.40,0.20,0.00),
                         labels = c("100%", "80%", "60%", "40%", "20%", "0%"))
    ggsave(paste0("tmpPlot-",imgNum,".png"), width = 7, height = 5, units = "in")
    insertImage(wb,
                sheet = paste(sheetName, oldfilters[TC]),
                paste0("tmpPlot-",imgNum,".png"),
                width = 7, 
                height = 5,
                startCol = "A",
                startRow = 26)
    imgList <- c(paste0("tmpPlot-",imgNum,".png"), imgList)
    imgNum <- imgNum + 1

    
    chart_2_data <- chart_data %>%
      filter(BCatDesc %in% c("B5a","B5b","B5c")) %>%
      group_by(name) %>%
      mutate(pos = cumsum(value) - value/2)
    
    ggplot(chart_2_data, aes(y=value, x = name, fill = BCatDesc)) +
      geom_bar(stat="identity", position = "stack") +
      #geom_text(aes(y = pos, label = format(value*100, digits = 1)), colour = "white") +
      theme_minimal() +
      scale_fill_manual(name = "Category",
                        values = c("B5a" = "#822433",
                                   "B5b" = "#00B092",
                                   "B5c" = "#002776")) +
      theme(axis.text.x = element_text(angle = 45, hjust = 0.95, size = 12),
            axis.text.y = element_text(size = 12),
            panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(size=.1, color="dark grey"),
            plot.title = element_text(face = "bold", colour = "black", size = 16,hjust = 0.5),
            axis.title.x = element_text(face = "bold", colour = "black", size = 14),
            axis.title.y = element_text(face = "bold", colour = "black", size = 14),
            legend.title = element_blank(), legend.position = "top", legend.spacing.x = unit(0.5, "cm"), 
            legend.text = element_text(size = 12)) + 
      labs(title = paste0("BQA ",oldfilters[TC]," data by B5 sub-category category,\n shown by pathologist with >5 total cses")) +
      scale_y_continuous(name = "Percentage of total", breaks = c(1.00,0.80,0.60,0.40,0.20,0.00),
                         labels = c("100%", "80%", "60%", "40%", "20%", "0%"))
    ggsave(paste0("tmpPlot-",imgNum,".png"), width = 7, height = 5, units = "in")
    insertImage(wb,
                sheet = paste(sheetName, oldfilters[TC]),
                paste0("tmpPlot-",imgNum,".png"),
                width = 7, 
                height = 5,
                startCol = "A",
                startRow = 51)
    imgList <- c(paste0("tmpPlot-",imgNum,".png"), imgList)
    imgNum <- imgNum + 1
    
    chart_3_data <- chart_data %>%
      filter(BCatDesc %in% c(c("B3 with atypia", "B3 without atypia", "B3 with atypia not specified"))) %>%
      group_by(name) %>%
      mutate(pos = cumsum(value) - value/2)
    
    ggplot(chart_3_data, aes(y=value, x = name, fill = BCatDesc)) +
      geom_bar(stat="identity", position = "stack") +
      #geom_text(aes(y = pos, label = format(value*100, digits = 1)), colour = "white") +
      theme_minimal() +
      scale_fill_manual(name = "Category",
                        values = c("B3 with atypia" = "#822433",
                                   "B3 without atypia" = "#00B092",
                                   "B3 with atypia not specified" = "#002776")) +
      theme(axis.text.x = element_text(angle = 45, hjust = 0.95, size = 12),
            axis.text.y = element_text(size = 12),
            panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(size=.1, color="dark grey"),
            plot.title = element_text(face = "bold", colour = "black", size = 16,hjust = 0.5),
            axis.title.x = element_text(face = "bold", colour = "black", size = 14),
            axis.title.y = element_text(face = "bold", colour = "black", size = 14),
            legend.title = element_blank(), legend.position = "top", legend.spacing.x = unit(0.5, "cm"), 
            legend.text = element_text(size = 12)) + 
      labs(title = paste0("BQA ",oldfilters[TC]," data by B3 sub-category category,\n shown by pathologists with >5 total cses")) +
      scale_y_continuous(name = "Percentage of total", breaks = c(1.00,0.80,0.60,0.40,0.20,0.00),
                         labels = c("100%", "80%", "60%", "40%", "20%", "0%"))
    ggsave(paste0("tmpPlot-",imgNum,".png"), width = 7, height = 5, units = "in")
    insertImage(wb,
                sheet = paste(sheetName, oldfilters[TC]),
                paste0("tmpPlot-",imgNum,".png"),
                width = 7, 
                height = 5,
                startCol = "A",
                startRow = 76)
    imgList <- c(paste0("tmpPlot-",imgNum,".png"), imgList)
    imgNum <- imgNum + 1
    
    writeData(wb, 
              sheet = paste(sheetName, oldfilters[TC]), 
              "Numerators", 
              startCol = 1, 
              startRow = 101, 
              rowNames = FALSE)
    writeData(wb, 
              sheet = paste(sheetName, oldfilters[TC]), 
              NumList[[k]], 
              startCol = 1, 
              startRow = 102, 
              rowNames = FALSE)
    writeData(wb, 
              sheet = paste(sheetName, oldfilters[TC]), 
              "Denominators", 
              startCol = 1, 
              startRow = 115, 
              rowNames = FALSE)
    writeData(wb, 
              sheet = paste(sheetName, oldfilters[TC]), 
              DenomList[[k]], 
              startCol = 1, 
              startRow = 116, 
              rowNames = FALSE)    
    }
}

saveWorkbook(wb, paste0(filesdir,"/SQAS BQA report generated_", Sys.Date(), ".xlsx"), overwrite = T)

### removes unneeded temporary information from environment
unlink(imgList)

options(warn = oldw)

rm(list = ls())
