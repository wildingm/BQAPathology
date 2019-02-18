### Use list of pathologists to provide lookup list for:
###                           Assigning pseudonymous codes to pathologists for BQA data tables
###                           Identifying unknown pathologists
###                           Updating list of pathologists with provided data

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
if (!require('excel.link')) {
  install.packages('excel.link')
  library('excel.link')
}

Pathologists<-xl.read.file("\\\\phe.gov.uk\\Health & Wellbeing\\HAW\\Quality Assurance\\QA SHARED\\Breast\\Data\\National pathology audit 2013-16\\CONFIDENTIAL P-Code look up\\Pathologist Codes BQA.xlsx",password = "M4r5hma11oW")
Pathologists<-filter(Pathologists,Pathologists$P.Code != "") ### removes empty rows
PCodes<-xl.read.file("\\\\phe.gov.uk\\Health & Wellbeing\\HAW\\Quality Assurance\\QA SHARED\\Breast\\Data\\National pathology audit 2013-16\\CONFIDENTIAL P-Code look up\\P_Codes.xlsx",password = "M4r5hma11oW")
PCodesUnused<-filter(PCodes,is.na(PCodes$Used))
  
FindPathCode <- function(NationalCode) {
  if (is.na(NationalCode)) {
    NewCode<-"Total"
  } else if (NationalCode=="") {
    NewCode<-"UNK_Pathologist"
  } else if (!is.na(NationalCode)|NationalCode!=""){
    NewCode<-as.character(Pathologists[Pathologists$Pathologist.GMC.number==NationalCode,2],stringsAsFactors = F)
    if (is_empty(NewCode)) {
      NewCode<-"UNK_Pathologist"
    }
  }
  NewCode
}

AssignCode <- function(GMCCode) {
  ### select a code
  newcode<-sample(PCodesUnused$P.Code,1)
  if (GMCCode %in% Pathologists$Pathologist.GMC.number) {
    print(paste("GMC Code",GMCCode,"is already present in the list of pathologists with code", Pathologists$P.Code[Pathologists$Pathologist.GMC.number==GMCCode],"- no action performed"))
  } else {  
    ### Add GMC code and P.Code to Pathologists data frame and YES to PCode data frame
    Pathologists<-rbind(Pathologists,c(NA,newcode,NA,GMCCode,NA,NA,NA))
    PCodes[PCodes$P.Code==newcode,2]<-"YES"
    print(paste("GMC code",GMCCode,"has been added to the list of pathologists with code", newcode))
  }
  newcode
}

MakePathTable<-function(list) {
  df<-bind_rows(list)
  df<-unique(df[,2])
  df2<-sapply(df,FindPathCode)
  pathlookup<-data_frame(df,df2)
  pathlookup
}

### Code to examine BQA downloads, extract new pathologist GMC codes , query the user to check if the codes should be updated and 
### then update the Pathologist and PCodes csv files

CheckFile<-winDialog(type = "yesno", "Do you want to check a BQA csv file for new pathologist GMC codes?")

while(CheckFile=="YES") {
  filessrc<-choose.files() 
  checksource<-read.csv(filessrc, skip = 2)
  checksource<-checksource[-nrow(checksource),1:23] #remove last row (states end of data), and extra columns with calculated fields
  checksource<-separate(checksource,col = Primary.Sort.Value,into = PrimarySortIDHeadings, sep = "\\*") #expands the Primary.Sort.Value field 
  checksourceselected<-checksource[,c(1,match("Path_national_code",names(checksource)),15:35)]
  colnames(checksourceselected)[2]<-"Primary.Sort.Value"
  
  codeslist<-as.character(unique(checksourceselected$Primary.Sort.Value))
  for (i in seq_along(codeslist)) {
    if (!(codeslist[i] %in% Pathologists$Pathologist.GMC.number) & !(is.na(codeslist[i])) & codeslist[i]!=""){
      checkcode<-winDialog(type = "yesno", paste("Is", codeslist[i], "a valid GMC code?"))
      if(checkcode == "YES") {
        AssignCode(codeslist[i])
      }
    }
  }
  xl.save.file(Pathologists, "\\\\phe.gov.uk\\Health & Wellbeing\\HAW\\Quality Assurance\\QA SHARED\\Breast\\Data\\National pathology audit 2013-16\\CONFIDENTIAL P-Code look up\\Pathologist Codes BQA.xlsx", row.names = FALSE, password = "M4r5hma11oW")
  xl.save.file(PCodes, "\\\\phe.gov.uk\\Health & Wellbeing\\HAW\\Quality Assurance\\QA SHARED\\Breast\\Data\\National pathology audit 2013-16\\CONFIDENTIAL P-Code look up\\P_Codes.xlsx", row.names = FALSE, password = "M4r5hma11oW")
  CheckFile <- winDialog(type = "yesno", "Do you want to check another file?")
}

rm(checksource,checksourceselected,checkcode,CheckFile,codeslist,PCodes)

