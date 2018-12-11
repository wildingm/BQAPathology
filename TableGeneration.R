### Project to extract BQA data, arrange and analyse in different levels and export analysis to tables and charts for QA visit data packs

oldw<-getOption("warn")
options(warn = -1)

if (exists("selector")) {rm(selector)}

source("Pathologist lookup.R")

if (!require('tools')) {
  install.packages('tools')
  library('tools')
}
if (!require('ggplot2')) {
  install.packages('ggplot2')
  library('ggplot2')
}

percent<-function(x,digits=2,format="f") {
  paste0(formatC(100*x,format=format,digits=digits),"%")
} #function to change a numerical value format to a percentage

BoxSelect<-function(df,rowID,colID) {
  Output<-as.numeric(df[df$Row.Identifier==rowID,colID])
  Output<-replace(Output,is.na(Output),0)
  if (is.numeric(Output) == F) {
    Output<-0
  }
  Output
} #selects the relevant box from the BQA file when passed the data frame 
#containing the data, the row ID (10, 20, 30, 40, 50, 60 or 9999) and the relevant column ID

DataExtract <- function (filename) {
  datasource<-read.csv(filename, skip = 3,stringsAsFactors = F)
  datasource<-datasource[-nrow(datasource),1:23] #remove last row (states end of data), and extra columns with calculated fields
  if(!selector=="Local_NBSS_Code") {
    datasource<-separate(datasource,col = "Primary.Sort.Value",into = PrimarySortIDHeadings, sep = "\\*") #expands the Primary.Sort.Value field 
    datasourceselected<-datasource[,c(1,match(selector,names(datasource)),15:35)]
    colnames(datasourceselected)[2]<-"Primary.Sort.Value"
  } else {
    datasourceselected<-datasource
  }
  datasourceselected$Tests.or.Clients..T.or.C.<-gsub(TRUE,"T",datasourceselected$Tests.or.Clients..T.or.C.)
  datasourceselected
}

DataExtractAll <- function (filename) {
  datasource<-read.csv(filename, skip = 3,stringsAsFactors = F)
  datasource<-datasource[-nrow(datasource),1:23] #remove last row (states end of data), and extra columns with calculated fields
  datasource<-separate(datasource,col = "Primary.Sort.Value",into = PrimarySortIDHeadings, sep = "\\*") #expands the Primary.Sort.Value field 
}

BQABarPlot<-function(rows,data,title,group) {
  chart_data<-melt(chartpropframe[rows,],id.var=data)
  chart_data$BQA_Measure<-gsub("Number of ","",chart_data$BQA_Measure)
  BQAplot<-ggplot(chart_data, aes(y=value, x = variable,fill = BQA_Measure)) +
    geom_bar(stat="identity",position = "stack") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 0.95, size = 12),
          axis.text.y = element_text(size = 12),
          panel.grid.major.x = element_blank(),
          panel.grid.major.y = element_line( size=.1, color="dark grey"),
          plot.title = element_text(face = "bold", colour = "black", size = 16,hjust = 0.5),
          axis.title.x = element_text(face = "bold", colour = "black", size = 14),
          axis.title.y = element_text(face = "bold", colour = "black", size = 14),
          legend.title = element_blank(), legend.position = "top", legend.spacing.x = unit(0.5, "cm"), 
          legend.text = element_text(size = 12)) +
    scale_fill_brewer(palette="Set1") +
    labs(title = paste("Proportion of tests broken down by",title,"\n data displayed by", group), x = group) +
    scale_y_continuous(name = "Percentage of total",breaks = c(1.00,0.80,0.60,0.40,0.20,0.00),
                       labels = c("100%","80%","60%","40%","20%","0%"))  #+ 
    #guides(fill = guide_legend(nrow = 2))
  BQAplot
}

winDialog(type = "ok", "Please choose the BQA files for analysis")
filessrc<-choose.files() 


#create bits for use later
TablesList<-list()

#define some names for use in the code
numerators<-c("BoxSelect(bqafilter,10,8),BoxSelect(bqafilter,60,8)","BoxSelect(bqafilter,40,17)","BoxSelect(bqafilter,40,17),
              BoxSelect(bqafilter,60,17)","BoxSelect(bqafilter,10,8),BoxSelect(bqafilter,10,12),BoxSelect(bqafilter,10,13),
              BoxSelect(bqafilter,60,8)","BoxSelect(bqafilter,9999,8)-BoxSelect(bqafilter,40,8)",
              "BoxSelect(bqafilter,9999,12)-BoxSelect(bqafilter,40,12)-BoxSelect(bqafilter,60,12)","BoxSelect(bqafilter,10,13)",
              "BoxSelect(bqafilter,9999,17)-BoxSelect(bqafilter,10,17)","BoxSelect(bqafilter,10,17)","BoxSelect(bqafilter,40,8)",
              "BoxSelect(bqafilter,10,17),BoxSelect(bqafilter,10,18)")
denominators<-c("BoxSelect(bqafilter,10,19),BoxSelect(bqafilter,60,8)","BoxSelect(bqafilter,40,19)","BoxSelect(bqafilter,40,19),
                BoxSelect(bqafilter,60,13),BoxSelect(bqafilter,60,17),BoxSelect(bqafilter,60,18)","BoxSelect(bqafilter,10,19),
                BoxSelect(bqafilter,60,8)","BoxSelect(bqafilter,9999,8)","BoxSelect(bqafilter,9999,12)-BoxSelect(bqafilter,60,12)",
                "BoxSelect(bqafilter,9999,13)","BoxSelect(bqafilter,9999,17)","BoxSelect(bqafilter,10,19),BoxSelect(bqafilter,60,8)",
                "BoxSelect(bqafilter,10,19),BoxSelect(bqafilter,60,8)","BoxSelect(bqafilter,60,8),BoxSelect(bqafilter,10,19)")
allNames<-c("Total Number of tests","Number of B1 (% of total)","Number of B2 (% of total)","Number of B3 with atypia not specified (% B3)",
            "Number of B3 without atypia (% B3)","Number of B3 with atypia (% B3)","Number of B3 (% of total)",
            "Number of B4 (% of total)","Number of B5c (% of B5)","Number of B5b (% of B5)","Number of B5a (% of B5)",
            "Number of B5 (% of total)","Absolute Sensitivity","Specificity (biopsy cases only)",
            "Specificity (Full)","Complete Sensitivity","PPV (B5)","PPV (B4)","PPV (B3)","Negative Predictive Value",
            "False Negative Rate","True False Positive Rate","Miss Rate")
PrimarySortIDHeadings<-c("Clinical_team", "Location_code","Location_name","RA_local_code","RA_local_name", "RA_national_code","Laboratory_code",
                         "Laboratory_name","Path_local_code","Path_local_name","Path_national_code","Loc_method","Radiological_appearance")
SelectionHeadings<-c(PrimarySortIDHeadings[8],PrimarySortIDHeadings[11:13])
  
checker<-read.csv(filessrc[1], skip = 3,stringsAsFactors = F)
if (!grepl("*",checker[1,2],fixed=T)) {
  selector<-"Local_NBSS_Code"
  group<-"NBSS pathologist code"
  winDialog(type = "ok", "Please choose if the data is to be analysed by tests or clients by typing the relevant numerical value into the console screen below")
  ToC<-select.list(c("Tests","Clients"))
  if (ToC == "Tests") {
    ToC<-"T"
  } else {
    ToC<-"C"
  }
  #### this will need to be passed to the code below somehow when subsetting the data
  #### checker$Tests.or.Clients..T.or.C.<-gsub(TRUE,"T",checker$Tests.or.Clients..T.or.C.)
} else {
  winDialog(type = "ok", "Please choose the primary sort ID by typing the relevant numerical value into the console screen below")
  ToC<-"T"
}
rm(checker)

RepeatExtract<-"YES"

while(RepeatExtract == "YES") {
  
  if (!exists("selector")) {
    selector<-select.list(SelectionHeadings)
    group<-c("laboratory", "pathologist", "localisation method", "radiological appearance")[SelectionHeadings==selector]
  }
  TableNames<-file_path_sans_ext(basename(filessrc))
  filessrcData<-lapply(filessrc,DataExtract)
  
  if (selector == "Path_national_code") {
    for (k in 1:length(filessrcData)) {
      filessrcData[[k]][2]<-unlist(lapply(filessrcData[[k]][[2]],FindPathCode))
    } ### renames all pathologists in the list based on their psedonym code or sets them to UNK Pathologist
  } 
  
  xl.workbook.add()
  
  for (k in 1:length(filessrcData)){
    
    numFrame<-data.frame(stringsAsFactors = FALSE)
    denomFrame<-data.frame(stringsAsFactors = FALSE)
    casesFrame<-data.frame(stringsAsFactors = FALSE)
    allsummary<-NULL
    
    bqadf<-filessrcData[[k]] %>%
      group_by(Primary.Sort.ID,Primary.Sort.Value,Secondary.Sort.ID,Secondary.Sort.Value,Tests.or.Clients..T.or.C.,Table.Identifier..A..B.or.C.,Row.Identifier) %>%
      summarise(WBN.B5 = sum(WBN.B5, na.rm=T),WBN.B5a = sum(WBN.B5a, na.rm=T),WBN.B5b = sum(WBN.B5b, na.rm=T),WBN.B5c = sum(WBN.B5c, na.rm=T),
                WBN.B4 = sum(WBN.B4, na.rm=T),WBN.B3 = sum(WBN.B3, na.rm=T),WBN.B3.wa = sum(WBN.B3.wa, na.rm=T),WBN.B3.na = sum(WBN.B3.na, na.rm=T),
                WBN.B3.ns = sum(WBN.B3.ns, na.rm=T),WBN.B2 = sum(WBN.B2, na.rm=T),WBN.B1 = sum(WBN.B1, na.rm=T),Total = sum(Total, na.rm=T))
    bqadf<-bqadf[bqadf$Table.Identifier..A..B.or.C.=="D",] # excludes the table ID other than D
    bqadf<-bqadf[bqadf$Tests.or.Clients..T.or.C.==ToC,] #excludes any rows where tests or clients is other than the specified value
    
    paths<-as.character(unique(bqadf$Primary.Sort.Value)) # creates a unique list Primary Sort Value
    # following code creates a data frame for the currently selected BQA data containing the numbers of cases for each pathologist
    # for each of the pathology categories
    for (j in 19:8) {
      casesBox = NULL
      for (i in paths) {
        bqafilter<-bqadf[bqadf$Primary.Sort.Value %in% i,]
        casesBox = append(casesBox, BoxSelect(bqafilter,9999,j))
      }
      casesBox<-unlist(casesBox)
      casesFrame<-rbind(casesFrame,casesBox)
    }
    
    colnames(casesFrame)<-paths #names the columns of the dataframe by pathologist code
    colnames(casesFrame)[ncol(casesFrame)]<-"TOT"   #renames the last column as the TOT column
    
    # following code creates 2 data frames for the currently selected BQA data containing the numerator and denominators for each 
    # pathologist for each of the calculated stats
    for (j in 1:length(numerators)) {
      numStore = NULL
      denomStore = NULL
      for (i in paths){
        numSum = NULL
        denomSum = NULL
        bqafilter<-bqadf[(bqadf$Primary.Sort.Value %in% i),]
        numSum<-eval(parse(text = paste(numSum,"sum(",numerators[j],")")))
        numStore<-append(numStore,numSum)
        numStore<-replace(numStore,is.na(numStore),0)
        denomSum<-eval(parse(text = paste(denomSum,"sum(",denominators[j],")")))
        denomStore<-append(denomStore,denomSum)
      }
      numFrame<-rbind(numFrame,numStore,stringsAsFactors = FALSE)
      denomFrame<-rbind(denomFrame,denomStore,stringsAsFactors = FALSE)
    }
    colnames(numFrame)<-paths                       #names the columns of the dataframe by pathologist code
    colnames(numFrame)[ncol(numFrame)]<-"TOT"       #renames the last column as the TOT column
    colnames(denomFrame)<-paths                     #names the columns of the dataframe by pathologist code
    colnames(denomFrame)[ncol(denomFrame)]<-"TOT"   #renames the last column as the TOT column
    calcSummarydf<-numFrame/denomFrame
    calcSummarydf[is.na(calcSummarydf)]<-NA
    calcSummarydf<-as.data.frame(lapply(calcSummarydf,percent),stringsAsFactors = FALSE)
    calcSummarydf<-replace(calcSummarydf,calcSummarydf==" NA%","No Cases")
    if (selector != "Path_national_code") {
      paths[nchar(paths)==0]<-paste0("UNK_",selector)
    }
    colnames(calcSummarydf)<-paths
    colnames(calcSummarydf)[ncol(calcSummarydf)]<-"TOT"
    colnames(casesFrame)<-paths
    colnames(casesFrame)[ncol(casesFrame)]<-"TOT"
    casesFrame<-casesFrame[,order(-casesFrame[1,])]    
    allsummary<-rbind(casesFrame,calcSummarydf)
    allsummary<-cbind("BQA_Measure"=allNames,allsummary)[c(1,12,8,7,3,2,11:9,6:4,13:23),]
    TablesList[[TableNames[k]]]<-allsummary
    
    ###produces proportions for calculated stats
    calcframe<-numFrame/denomFrame
    calcframe[is.na(calcframe)]<-0
    ###produces full stats table as values
    chartFrame<-rbind(casesFrame,calcframe)
    chartFrame<-cbind("BQA_Measure"=allNames,chartFrame)[c(1,12,8,7,3,2,11:9,6:4,13:23),]
    ###produces full stats table as values
    propFrame<-casesFrame
    for (i in 9:11) {propFrame[i,]<-propFrame[i,]/propFrame[12,]}
    for (i in 4:6) {propFrame[i,]<-propFrame[i,]/propFrame[7,]}
    for (i in c(2,3,7,8,12)) {propFrame[i,]<-propFrame[i,]/propFrame[1,]}
    propFrame[1,]<-propFrame[1,]/propFrame[1,]
    propFrame[is.na(propFrame)]<-0
    chartpropframe<-rbind(propFrame,calcframe)
    chartpropframe<-cbind("BQA_Measure"=allNames,chartpropframe)[c(1,12,8,7,3,2,11:9,6:4,13:23),]

    ###produces bar charts for the report output
    BCatPlot<-BQABarPlot(2:6,"BQA_Measure","B category", group)
    B5Plot<-BQABarPlot(7:9,"BQA_Measure","B5 sub-category", group)
    B3Plot<-BQABarPlot(10:12,"BQA_Measure","B3 sub-category", group)
    
    ###Prints the charts to an excel workbook
    xl.sheet.add(TableNames[k])
    xl.write(allsummary,xl.get.excel()[["ActiveSheet"]]$Cells(1,1),row.names = FALSE)
    print(BCatPlot)
    xl[a26] = current.graphics(width=1000)
    print(B5Plot)
    xl[a51] = current.graphics(width=1000)
    print(B3Plot)
    xl[a76] = current.graphics(width=1000)
    
  }  
  
  ###### Produces Table using combined data from all selected files
  
  filessrcDataTotal<-bind_rows(filessrcData)
  
  numFrame<-data.frame(stringsAsFactors = FALSE)
  denomFrame<-data.frame(stringsAsFactors = FALSE)
  casesFrame<-data.frame(stringsAsFactors = FALSE)
  allsummary<-NULL
  
  bqadf<-filessrcDataTotal %>%
    group_by(Primary.Sort.ID,Primary.Sort.Value,Secondary.Sort.ID,Secondary.Sort.Value,Tests.or.Clients..T.or.C.,Table.Identifier..A..B.or.C.,Row.Identifier) %>%
    summarise(WBN.B5 = sum(WBN.B5, na.rm=T),WBN.B5a = sum(WBN.B5a, na.rm=T),WBN.B5b = sum(WBN.B5b, na.rm=T),WBN.B5c = sum(WBN.B5c, na.rm=T),
              WBN.B4 = sum(WBN.B4, na.rm=T),WBN.B3 = sum(WBN.B3, na.rm=T),WBN.B3.wa = sum(WBN.B3.wa, na.rm=T),WBN.B3.na = sum(WBN.B3.na, na.rm=T),
              WBN.B3.ns = sum(WBN.B3.ns, na.rm=T),WBN.B2 = sum(WBN.B2, na.rm=T),WBN.B1 = sum(WBN.B1, na.rm=T),Total = sum(Total, na.rm=T))
  bqadf<-bqadf[bqadf$Table.Identifier..A..B.or.C.=="D",] # excludes the table ID other than D
  bqadf<-bqadf[bqadf$Tests.or.Clients..T.or.C.==ToC,] #excludes any rows where tests or clients is other than the specified value
  
  paths<-as.character(unique(bqadf$Primary.Sort.Value)) # creates a unique list Primary Sort Value
  
  ### produces table of pathologists
  if (selector == "Path_national_code") {
    allinfo<-lapply(filessrc,DataExtractAll)
    allinfo<-bind_rows(allinfo)
    pathlook<-Pathologists
    pathtable<-data.frame("NBSS_national_code"=unique(allinfo$Path_national_code),"NBSS_local_name"=NA,"National_P_code_(if_known)"=NA,"National_name_(if_known)"=NA,stringsAsFactors = FALSE)
    for (i in 1:nrow(pathtable)){
      pathtable$NBSS_local_name[i]<-allinfo$Path_local_name[match(pathtable$NBSS_national_code[i],allinfo$Path_national_code)]
      pathtable$National_P_code_.if_known.[i]<-pathlook$P.Code[match(pathtable$NBSS_national_code[i],pathlook$Pathologist.GMC.number)]
      pathtable$National_name_.if_known.[i]<-pathlook$Pathologist.Full.Name[match(pathtable$NBSS_national_code[i],pathlook$Pathologist.GMC.number)]    
    }
    rm(allinfo)
    rm(pathlook)
    write.csv(pathtable,"Pathologist details for selected files.csv",row.names = FALSE)
  }
  
  # following code creates a data frame for the currently selected BQA data containing the numbers of cases for each pathologist
  # for each of the pathology categories
  for (j in 19:8) {
    casesBox = NULL
    for (i in paths) {
      bqafilter<-bqadf[bqadf$Primary.Sort.Value %in% i,]
      casesBox = append(casesBox, BoxSelect(bqafilter,9999,j))
    }
    casesBox<-unlist(casesBox)
    casesFrame<-rbind(casesFrame,casesBox)
  }
  colnames(casesFrame)<-paths #names the columns of the dataframe by pathologist code
  colnames(casesFrame)[ncol(casesFrame)]<-"TOT"   #renames the last column as the TOT column
  
  # following code creates 2 data frames for the currently selected BQA data containing the numerator and denominators for each 
  # pathologist for each of the calculated stats
  for (j in 1:length(numerators)) {
    numStore = NULL
    denomStore = NULL
    for (i in paths){
      numSum = NULL
      denomSum = NULL
      bqafilter<-bqadf[(bqadf$Primary.Sort.Value %in% i),]
      numSum<-eval(parse(text = paste(numSum,"sum(",numerators[j],")")))
      numStore<-append(numStore,numSum)
      numStore<-replace(numStore,is.na(numStore),0)
      denomSum<-eval(parse(text = paste(denomSum,"sum(",denominators[j],")")))
      denomStore<-append(denomStore,denomSum)
    }
    numFrame<-rbind(numFrame,numStore,stringsAsFactors = FALSE)
    denomFrame<-rbind(denomFrame,denomStore,stringsAsFactors = FALSE)
  }
  colnames(numFrame)<-paths                       #names the columns of the dataframe by pathologist code
  colnames(numFrame)[ncol(numFrame)]<-"TOT"       #renames the last column as the TOT column
  colnames(denomFrame)<-paths                     #names the columns of the dataframe by pathologist code
  colnames(denomFrame)[ncol(denomFrame)]<-"TOT"   #renames the last column as the TOT column
  calcSummarydf<-numFrame/denomFrame
  calcSummarydf[is.na(calcSummarydf)]<-NA
  calcSummarydf<-as.data.frame(lapply(calcSummarydf,percent),stringsAsFactors = FALSE)
  calcSummarydf<-replace(calcSummarydf,calcSummarydf==" NA%","No Cases")
  if (selector != "Path_national_code") {
    paths[nchar(paths)==0]<-paste0("UNK_",selector)
  }
  colnames(calcSummarydf)<-paths
  colnames(calcSummarydf)[ncol(calcSummarydf)]<-"TOT"
  colnames(casesFrame)<-paths
  colnames(casesFrame)[ncol(casesFrame)]<-"TOT"
  casesFrame<-casesFrame[,order(-casesFrame[1,])]
  allsummary<-rbind(casesFrame,calcSummarydf)
  allsummary<-cbind("BQA_Measure"=allNames,allsummary)[c(1,12,8,7,3,2,11:9,6:4,13:23),]
  
  ###produces proportions for calculated stats
  calcframe<-numFrame/denomFrame
  calcframe[is.na(calcframe)]<-0
  ###produces full stats table as values
  chartFrame<-rbind(casesFrame,calcframe)
  chartFrame<-cbind("BQA_Measure"=allNames,chartFrame)[c(1,12,8,7,3,2,11:9,6:4,13:23),]
  ###produces full stats table as values
  propFrame<-casesFrame
  for (i in 9:11) {propFrame[i,]<-propFrame[i,]/propFrame[12,]}
  for (i in 4:6) {propFrame[i,]<-propFrame[i,]/propFrame[7,]}
  for (i in c(2,3,7,8,12)) {propFrame[i,]<-propFrame[i,]/propFrame[1,]}
  propFrame[1,]<-propFrame[1,]/propFrame[1,]
  propFrame[is.na(propFrame)]<-0
  chartpropframe<-rbind(propFrame,calcframe)
  chartpropframe<-cbind("BQA_Measure"=allNames,chartpropframe)[c(1,12,8,7,3,2,11:9,6:4,13:23),]

  ###produces bar charts for the report output
  BCatPlot<-BQABarPlot(2:6,"BQA_Measure","B category", group)
  B5Plot<-BQABarPlot(7:9,"BQA_Measure","B5 sub-category", group)
  B3Plot<-BQABarPlot(10:12,"BQA_Measure","B3 sub-category", group)
  
  ###Prints the charts to an excel workbook
  xl.sheet.add()
  xl.sheet.name("Total")
  xl.write(allsummary,xl.get.excel()[["ActiveSheet"]]$Cells(1,1),row.names = FALSE)
  print(BCatPlot)
  xl[a26] = current.graphics(width=1000)
  print(B5Plot)
  xl[a51] = current.graphics(width=1000)
  print(B3Plot)
  xl[a76] = current.graphics(width=1000)
  #xl.sheet.add(TableNames[1])
  xl.sheet.delete("Sheet1")
  xl.workbook.save(paste(selector,"generated",Sys.Date()))
  
  if(!selector=="Local_NBSS_Code") {
    RepeatExtract<-winDialog(type = "yesno", "Do you wish to extract data for another primary sort value? If yes please enter the relevant numerical value in the console below.")
  } else {
    RepeatExtract<-"NO"
  }
  rm(selector)
}
rm(Pathologists)
options(warn = oldw)