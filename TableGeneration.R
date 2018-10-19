### Project to extract BQA data, arrange and analyse in different levels and export analysis to tables and charts for QA visit data packs

oldw<-getOption("warn")
options(warn = -1)

source("Pathologist lookup.R")

library(tools)

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
  datasource<-read.csv(filename, skip = 2,stringsAsFactors = F)
  datasource<-datasource[-nrow(datasource),1:23] #remove last row (states end of data), and extra columns with calculated fields
  datasource<-separate(datasource,col = "Primary.Sort.Value",into = PrimarySortIDHeadings, sep = "\\*") #expands the Primary.Sort.Value field 
  datasourceselected<-datasource[,c(1,match(selector,names(datasource)),15:35)]
  colnames(datasourceselected)[2]<-"Primary.Sort.Value"
  datasourceselected
}

DataExtractAll <- function (filename) {
  datasource<-read.csv(filename, skip = 2,stringsAsFactors = F)
  datasource<-datasource[-nrow(datasource),1:23] #remove last row (states end of data), and extra columns with calculated fields
  datasource<-separate(datasource,col = "Primary.Sort.Value",into = PrimarySortIDHeadings, sep = "\\*") #expands the Primary.Sort.Value field 
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

winDialog(type = "ok", "Please choose the primary sort ID by typing the relevant numerical value into the console screen below")

RepeatExtract<-"YES"

while(RepeatExtract == "YES") {
  
  selector<-select.list(PrimarySortIDHeadings)
  TableNames<-paste("Table",file_path_sans_ext(basename(filessrc)), sep = "_")
  filessrcData<-lapply(filessrc,DataExtract)
  
  if (selector == "Path_national_code") {
    for (k in 1:length(filessrcData)) {
      filessrcData[[k]][2]<-unlist(lapply(filessrcData[[k]][[2]],FindPathCode))
    } ### renames all pathologists in the list based on their psedonym code or sets them to UNK Pathologist
  } 
  
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
  }  
  for (i in seq_along(TablesList)) {
    filename= paste0(TableNames[i],"_",selector,".csv")
    write.csv(TablesList[[i]],filename,row.names = FALSE)
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
  filename<-paste0(winDialogString("Please enter a filename for the data table using all files selected.", paste0("Total of all files_",selector)),".csv")
  write.csv(allsummary,filename,row.names = FALSE)
  
  RepeatExtract<-winDialog(type = "yesno", "Do you wish to extract data for another primary sort value? If yes please enter the relevant numerical value in the console below.")
}
rm(Pathologists)
options(warn = oldw)
