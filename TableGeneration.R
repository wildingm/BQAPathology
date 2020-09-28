### Project to extract BQA data, arrange and analyse in different levels and export analysis to tables and charts for QA visit data packs

oldw <- getOption("warn")
options(warn = -1)

if (exists("selector")) {
  rm(selector)
}

#source("Pathologist lookup.R")

if (!require('tools')) {
  install.packages('tools')
  library('tools')
}
if (!require('ggplot2')) {
  install.packages('ggplot2')
  library('ggplot2')
}
if (!require('scales')) {
  install.packages('scales')
  library('scales')
}
if (!require('purrr')) {
  install.packages('purrr')
  library('purrr')
}
if (!require('dplyr')) {
  install.packages('dplyr')
  library('dplyr')
}
if (!require('tidyr')) {
  install.packages('tidyr')
  library('tidyr')
}
if (!require('reshape2')) {
  install.packages('reshape2')
  library('reshape2')
}
if (!require('rlang')) {
  install.packages('rlang')
  library('rlang')
}
# if (!require('excel.link')) {
#   install.packages('excel.link')
#   library('excel.link')
# }
if (!require('openxlsx')) {
  install.packages('openxlsx')
  library('openxlsx')
}

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

BQABarPlot <- function(dataframe, rows, data, title, group) {
  # Generates a stacked bar chart using the ggplot2 package
  #
  # Args:
  #   dataframe: A dataframe containing the BQA data to be used in the chart
  #   rows: The rows to extract from the dataframe to use for the chart
  #   data: The column name to use for the id.var in melting the dataset
  #   title: A string containing B category information to use for the chart title
  #   group: A string containing information about the data split to use for the chart title
  #
  # Returns:
  #   A plot object that generates a stacked bar chart based on the Args inputted
  chart_data <- melt(dataframe[rows, ], id.var=data)
  chart_data$BCatDesc <- gsub("Number of ", "", chart_data$BCatDesc)
  BQAplot <- ggplot(chart_data, aes(y=value, x = variable, fill = BCatDesc)) +
    geom_bar(stat="identity", position = "stack") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 0.95, size = 12),
          axis.text.y = element_text(size = 12),
          panel.grid.major.x = element_blank(),
          panel.grid.major.y = element_line(size=.1, color="dark grey"),
          plot.title = element_text(face = "bold", colour = "black", size = 16,hjust = 0.5),
          axis.title.x = element_text(face = "bold", colour = "black", size = 14),
          axis.title.y = element_text(face = "bold", colour = "black", size = 14),
          legend.title = element_blank(), legend.position = "top", legend.spacing.x = unit(0.5, "cm"), 
          legend.text = element_text(size = 12)) + scale_fill_brewer(palette="Set1") +
    labs(title = paste("Proportion of", tolower(oldfilters[TC]), "broken down by", title, "\n data displayed by", group), x = group) +
    scale_y_continuous(name = "Percentage of total", breaks = c(1.00,0.80,0.60,0.40,0.20,0.00),
                       labels = c("100%", "80%", "60%", "40%", "20%", "0%"))  
  BQAplot
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
    data <- melt(BoxList[[i]])
    temp <- temp %>% bind_rows(data)
  }
  temp[is.na(temp)] <- 0 # changes NA values to 0
  temp <- cbind(Measure = MeasureName, dcast(temp, Filename~variable, sum))
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
    summarise(value = sum(value, na.rm = T)) %>%
    spread(!!enquo(colID), value)
  t
}

winDialog(type = "ok", "Please choose the BQA files for analysis")
filessrc <- choose.files() 

### Variable vectors and lists
BCatDesc <- c("Total Number of tests", "Number of B1", "Number of B2", "Number of B3 with atypia not specified",
  "Number of B3 with atypia", "Number of B3 without atypia", "Number of B3", "Number of B4", "Number of B5c", "Number of B5b",
  "Number of B5a", "Number of B5")
BCategory <- c("Total", "WBN.B1", "WBN.B2", "WBN.B3.ns", "WBN.B3.wa", "WBN.B3.na", "WBN.B3", "WBN.B4", "WBN.B5c",
               "WBN.B5b", "WBN.B5a", "WBN.B5")
BCatlookup <- data.frame(BCategory, BCatDesc)
BCatOrder <- c("Total Number of tests", "Number of B5", "Number of B4", "Number of B3", "Number of B2", "Number of B1", "Number of B5a",
             "Number of B5b", "Number of B5c", "Number of B3 with atypia", "Number of B3 without atypia",
             "Number of B3 with atypia not specified")
BoxIDVector <- c("1", "5", "6", "7", "8", "9", "28", "32", "34", "36", "37", "41", "42", "43", "44", "46", "50", "51", "53", "54", "52")
BoxCatIDFilters <- c('WBN.B5', 'WBN.B4', 'WBN.B3', 'WBN.B2', 'WBN.B1', 'Total', 'WBN.B5', 'WBN.B4', 'WBN.B2', 'Total', 'WBN.B5', 'WBN.B4',
                   'WBN.B3', 'WBN.B2', 'WBN.B1', 'WBN.B5', 'WBN.B4', 'WBN.B3', 'WBN.B1', 'Total', 'WBN.B2')
BoxRowIDFilters <- c(10, 10, 10, 10, 10, 10, 40, 40, 40, 40, 60, 60, 60, 60, 60, 9999, 9999, 9999, 9999, 9999, 9999)
subtractNum <- list("28", c("32", "41"), "7")
subtractNumNames <- c("PPV (B5)", "PPV (B4)", "Negative Predictive Value")
calcnames <- c("Absolute Sensitivity", "Specificity (biopsy cases only)",
             "Specificity (Full)", "Complete Sensitivity", "PPV (B5)", "PPV (B4)", "PPV (B3)", "Negative Predictive Value",
             "False Negative Rate", "True False Positive Rate", "Miss Rate")
calcNumSumBoxes <- list(c("1", "37"), "34", c("34", "43"), c("1", "5", "6", "37"), "46", "50", "6", "52", "7", "28", c("7", "8"))
calcDenomSumBoxes <- list(c("9", "37"), "36", c("36", "42", "43", "44"), c("9", "37"), "46", "50", "51", "52", c("9", "37"), c("9", "37"),
                          c("9", "37"))
chartnames <- c("BcatPlot", "B5Plot", "B3Plot")
chartrownums <- list(c(1:5), c(1:3), c(1:3))
chartrows <- list(c("Number of B5", "Number of B4", "Number of B3", "Number of B2", "Number of B1"), 
                  c("Number of B5a", "Number of B5b", "Number of B5c"), 
                  c("Number of B3 with atypia", "Number of B3 without atypia", "Number of B3 with atypia not specified"))
chartdatastring <- c("B Category", "B5 sub-category", "B3 sub-category")
oldfilters <- c("Tests", "Clients")
newfilters <- c("Path_pseudo_code", "Laboratory_name", "Loc_method", "Radiological_appearance")
TCpicker <- c("T", "C")

### extracts the filenames from filessrc
TableNames <- file_path_sans_ext(basename(filessrc))

#xl.workbook.add() excel.link function that no longer works :(



    ### Review all code to remove reference to compile for BSIS options ###

### This will simplify code and make it easier to update

# Reads the first file in the filenames and identifies if it is a BSIS extract or not, assigns information for downstream processing 
# depending on if ithe file is BSIS style or not

#checker <- read.csv(filessrc[1], skip = 3,stringsAsFactors = F)
#checker <- checker[2, 2]
#if (!grepl("*", checker, fixed=T)) { #if there is no * in the primary sort code this identifies if the data is not BSIS

## Options for C and T data only:
  group <- "NBSS pathologist code"
  PrimarySortIDHeadings <- "Primary_Sort_Value"

  # } else {
#   group <- newfilters
#   PrimarySortIDHeadings <- c("Clinical_team", "Location_code", "Location_name", "RA_local_code", "RA_local_name", "RA_national_code",
#                              "Laboratory_code", "Laboratory_name", "Path_local_code", "Path_local_name", "Path_national_code",
#                              "Loc_method", "Radiological_appearance")
#   
# Extracts the data from the files selected and extracts the pathologist data based on GMC code
#   allinfo <- lapply(filessrc, DataExtractAll)  %>%
#     bind_rows()
#   pathlook <- Pathologists
#   pathtable <- data.frame("NBSS_national_code" = unique(allinfo$Path_national_code), "NBSS_local_name" = NA, 
#                           "National_P_code_(if_known)" = NA, "National_name_(if_known)" = NA, stringsAsFactors = FALSE)
#   for (i in 1:nrow(pathtable)) {
#     pathtable$NBSS_local_name[i] <- allinfo$Path_local_name[match(pathtable$NBSS_national_code[i], allinfo$Path_national_code)]
#     pathtable$National_P_code_.if_known.[i] <- pathlook$P.Code[match(pathtable$NBSS_national_code[i], pathlook$Pathologist.GMC.number)]
#     pathtable$National_name_.if_known.[i] <- pathlook$Pathologist.Full.Name[match(pathtable$NBSS_national_code[i], 
#                                                                                   pathlook$Pathologist.GMC.number)]
#   }
#   rm(allinfo, pathlook)
#   xl.sheet.add("Pathologist details")
#   xl.write(pathtable, xl.get.excel()[["ActiveSheet"]]$Cells(1, 1), row.names = FALSE)
# }

# Extract the data from the file(s)
tidydataset <- lapply(filessrc, DataExtractAll)
tidydataset <- mapply(cbind, tidydataset, "Filename"=TableNames, SIMPLIFY = F)
tidydataset <- tidydataset %>%
  bind_rows() %>%
  select(-Total.Cases.Screened, -Total.Assessed, -Total.WBN.Performed, -Total.VAE.Performed) %>%
  
  ### consider pivot longer for better readability
  melt(id.vars = names(.)[c(1:which(colnames(.)=="Row.Identifier"), which(colnames(.)=="Filename"))],
       variable.name = "BCategory") %>%
  filter(Table.Identifier..A..B.or.C.=="D")
tidydataset$Path_pseudo_code <- unlist(lapply(tidydataset$Path_national_code, FindPathCode)) #### move into loop? or deal with checker first and then apply this if * present
tidydataset <- inner_join(tidydataset, BCatlookup)
tidydataset$value[is.na(tidydataset$value)] <- 0
tidydataset$Tests.or.Clients..T.or.C. <- gsub(TRUE, "T", tidydataset$Tests.or.Clients..T.or.C.) # convert TRUE values into T string

# double up the loop to go through either tests/clients or the different categories.

for (TC in 1:length(unique(tidydataset$Tests.or.Clients..T.or.C.))) {
  tidydatasetuse <- tidydataset[tidydataset$Tests.or.Clients..T.or.C. == TCpicker[TC],]
  
  for (CC in 1:length(group)) { ### MUCH OF THIS UNNECESSARY NOW
    if (group[CC] == "Radiological_appearance") {
      BQAtablescombined <- TableGenerator(tidydatasetuse, Row.Identifier==9999, Radiological_appearance,
                                          Filename, Radiological_appearance, BCatDesc)
    } else if (group[CC] == "Path_pseudo_code") {
      BQAtablescombined <- TableGenerator(tidydatasetuse, Row.Identifier==9999, Path_pseudo_code, Filename, Path_pseudo_code, BCatDesc)
    } else if (group[CC] == "Laboratory_name") {
      BQAtablescombined <- TableGenerator(tidydatasetuse, Row.Identifier==9999, Laboratory_name, Filename, Laboratory_name, BCatDesc)
    } else if (group[CC] == "Loc_method") {
      BQAtablescombined <- TableGenerator(tidydatasetuse, Row.Identifier==9999, Loc_method, Filename, Loc_method, BCatDesc)
    } else {
      BQAtablescombined <- TableGenerator(tidydatasetuse, Row.Identifier==9999, Primary_Sort_Value, Filename, Primary_Sort_Value,BCatDesc)
    }
    BQAtablescombined <- left_join(data.frame("BCatDesc" = BCatOrder), BQAtablescombined)
    BQAtablescombined$BCatDesc <- as.character(BQAtablescombined$BCatDesc)
    
    # This chunk produces a list of data frames that contain the values for the BQA box numbers identified by BoxIDVector
    BoxList <- list()
    for (k in 1:length(BoxIDVector)) {
      if (group[CC] == "Radiological_appearance") {
        BoxList[[BoxIDVector[k]]] <- TableGenerator(tidydatasetuse, Row.Identifier == BoxRowIDFilters[k] & BCategory == BoxCatIDFilters[k],
                                                    Radiological_appearance, Filename, Radiological_appearance)
      } else if (group[CC] == "Path_pseudo_code") {
        BoxList[[BoxIDVector[k]]] <- TableGenerator(tidydatasetuse, Row.Identifier == BoxRowIDFilters[k] & BCategory==BoxCatIDFilters[k],
                                                    Path_pseudo_code, Filename, Path_pseudo_code)
      } else if (group[CC] == "Laboratory_name") {
        BoxList[[BoxIDVector[k]]] <- TableGenerator(tidydatasetuse, Row.Identifier==BoxRowIDFilters[k] & BCategory==BoxCatIDFilters[k],
                                                    Laboratory_name, Filename, Laboratory_name)
      } else if (group[CC] == "Loc_method") {
        BoxList[[BoxIDVector[k]]] <- TableGenerator(tidydatasetuse, Row.Identifier==BoxRowIDFilters[k] & BCategory==BoxCatIDFilters[k],
                                                    Loc_method, Filename, Loc_method)
      } else {
        BoxList[[BoxIDVector[k]]] <- TableGenerator(tidydatasetuse, Row.Identifier==BoxRowIDFilters[k] & BCategory==BoxCatIDFilters[k],
                                                    Primary_Sort_Value, Filename, Primary_Sort_Value)
      }
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
    tidydenomframemelted <- melt(tidydenomframe, id.vars = c("BCatDesc", "Filename"))
    names(tidydenomframemelted)[names(tidydenomframemelted) == "value"] <- "denomvalue"
    tidycalcframemelted <- melt(tidycalcframe, id.vars = c("BCatDesc", "Filename"))
    tidycalcframemelted$value <- percent_format(accuracy = .01)(tidycalcframemelted$value)
    tidycalcframemelted <- inner_join(tidycalcframemelted, tidydenomframemelted)
    tidycalcframemelted$value[tidycalcframemelted$denomvalue == 0] <- "No Cases"
    tidycalcframemelted <- select(tidycalcframemelted, -denomvalue)
    tidycalcframeperc <- dcast(tidycalcframemelted, Filename + BCatDesc ~ variable, alue.var="value")
    
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
      TotalList[[k]] <- bind_rows(mutate_all(TablesList[[k]], as.character), CalcList[[k]])
    }
    
    for (i in seq_along(TotalList)) {
      DenomList[[i]] <- DenomList[[i]][names(TotalList[[i]])]
      NumList[[i]] <- NumList[[i]][names(TotalList[[i]])]
    }
    
    chartframeplot <- BQABarPlot(TotalList[[1]], 2:6, "BCatDesc", "B category", group)
    
    # creates a dataframe containing the proportion of the different B5 categories in place of the numbers of cases
    chartframe <- TotalList[[1]]
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
                               as.numeric(chartframe[chartframe$BCatDesc == "Total Number of tests",2:ncol(chartframe)])))
    }

# ==========================================================================================================================================    
    
    # creates the list to put the plots
    PlotList <- list()
    
    # for each of the plots as identified by chartnames
    for (i in TableNames) {
      tempPlotList <- list()
      for (j in seq_along(chartnames)) {
        chart_data <- chartframe[chartframe$BCatDesc %in% chartrows[[j]], ] # filter chartframe to only include the rows required
        chart_data[is.na(chart_data)] <- 0 # change any remaining NA values to 0
        chart_data <- chart_data %>%
          mutate_at(c(2:ncol(chart_data)), funs(as.numeric(as.character(.))))
        
        [2:ncol(chart_data)] <- as.numeric(unlist(chart_data[2:ncol(chart_data)])) # change the values from char to numeric
        # at this point there should be a dataframe in wide format ready to plug into the chart generation formula
        tempPlotList[[chartnames[j]]] <- BQABarPlot(chart_data, chartrownums[[j]], 'BCatDesc', chartdatastring[j], "pathologist")
      }
      PlotList[[i]] <- tempPlotList
    } 

# =========================================================================================================================================

    # pastes the data in the different lists and the charts into a worksheet in the opened workbook for each filename in TableNames
    for (k in seq_along(TableNames)) {
      
      # Because the sheet name can only be 31 characters max, the filename is truncated to use the last 20 characters in the TableName entry
      # this may be an issue if the filename is not informative enough.
      
      if (nchar(TableNames[k]) > 20) {
        sheetName <- trimws(substr(TableNames[k], nchar(TableNames[k]) - 19, nchar(TableNames[k])), which = "left")
      } else {
        sheetName <- TableNames[k]
      }
      xl.sheet.add(paste(sheetName, oldfilters[TC], substr(group[CC], 1, 3)))
      xl.write(TotalList[[k]], xl.get.excel()[["ActiveSheet"]]$Cells(1, 1), row.names = FALSE)
      print(PlotList[[k]]$BcatPlot)
      xl[a26] = current.graphics(width=1000)
      print(PlotList[[k]]$B5Plot)
      xl[a51] = current.graphics(width=1000)
      print(PlotList[[k]]$B3Plot)
      xl[a76] = current.graphics(width=1000)
      xl.write("Numerators", xl.get.excel()[["ActiveSheet"]]$Cells(101,1), row.names = FALSE)
      xl.write(NumList[[k]], xl.get.excel()[["ActiveSheet"]]$Cells(102,1), row.names = FALSE)
      xl.write("Denominators", xl.get.excel()[["ActiveSheet"]]$Cells(115,1), row.names = FALSE)
      xl.write(DenomList[[k]], xl.get.excel()[["ActiveSheet"]]$Cells(116,1), row.names = FALSE)
    }
  }
}

xl.sheet.delete("Sheet1")
xl.workbook.save(paste("SQAS BQA report generated", Sys.Date()))

### removes unneeded temporary information from enviroment
# rm(subframe, common, subtractNumNames, subtractNum, calcDenomSumBoxes, calcNumSumBoxes, BCatDesc, BCategory, BoxCatIDFilters,
#    BoxIDVector, chartframe, chartrownums, chartrows, chart_data, chartframeplot, oldfilters, newfilters, checker, allinfo, pathlook,
#    pathtable, PCodesUnused, Pathologists, sheetName, TableNames, group, PrimarySortIDHeadings, filessrc, BCatlookup, BoxList, 
#    BQAtablescombined, CalcList, DenomList, NumList, PlotList, TablesList, tempPlotList, tidycalcframe, tidycalcframemelted,
#    tidycalcframeperc, tidydataset, tidydatasetuse, tidydenomframe, tidydenomframemelted, tidynumframe, TotalList,
#    i, j, k, TC, TCpicker, chartdatastring, chartnames, CC, calcnames, BoxRowIDFilters, BCatOrder, xl)

options(warn = oldw)

rm(oldw)
