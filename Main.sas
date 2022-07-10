/* Notes: Added by LeighAMP 
   Last modified: June 26 2022*/

1

2

%macro get_current_path();
  %local path_name name;
  %let path_name=%bquote(%sysget(SAS_EXECFILEPATH));
  %let name=%bquote(%sysget(SAS_EXECFILENAME));
  %sysfunc(tranwrd(&path_name,&name,))
  %put path_name=&path_name;
%mend;
%let CurrentPath=%get_current_path(); 

%include "&CurrentPath.Summary_Report.sas";

/* create reports in current folder */

%summary_report(
  DataSet         = Sashelp.Heart,
  GroupVars       = Sex Status,
  StudyVars       = Smoking_Status BP_Status Weight Cholesterol,
  StudyVarTypes   = cat cat con con,
  Statistics      = N Nmiss Mean STD Median Min Max 
                    P1 P99 QRANGE LCLM UCLM SKEWNESS KURTOSIS,
  MaxDec          = 2,
  Margin          = 0.5,       
  PaperSize       = LETTER,
  Orientation     = Portrait,
  Style           = 2,
  FontSize        = 7,
  RowShadow       = 1,
  Output_RTF_File = "&CurrentPath.heart_summary.rtf"
);

%summary_report(
  DataSet         = Sashelp.cars,
  GroupVars       = Origin DriveTrain,
  StudyVars       = Cylinders MPG_City Make Horsepower,
  StudyVarTypes   = cat con cat con,
  MaxDec          = 2,
  PaperSize       = LEDGER,
  Orientation     = Landscape,
  FirstColumnWidth= 2.0,
  Style           = 2,
  FontSize        = 10,
  RowShadow       = 3,
  Output_RTF_File = "&CurrentPath.cars_summary.rtf"
);


* 3 level of groups, larger paper size and smaller font size;
* style 1;
%summary_report(
  DataSet         = sashelp.prdsale,
  GroupVars       = COUNTRY Region Division,
  StudyVars       = YEAR QUARTER MONTH ACTUAL PREDICT,
  StudyVarTypes   = cat cat cat con con,
  MaxDec          = 2,
  PaperSize       = LEDGER,
  Orientation     = Landscape,
  Style           = 1,
  FontSize        = 7,
  RowShadow       = 0,
  Output_RTF_File = "&CurrentPath.prdsale_summary.rtf"
);
