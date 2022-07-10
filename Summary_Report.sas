/* 

This macro will create a detailed report in rich text format for the dataset
by nested groups for the given categorical and continuous study variables.

P values of the difference for the major groups are calculated. 

Most comments were removed for showing maximum content in one page. 

For how to use the macro, see the sample file Main.sas

*/

%macro summary_report(
  DataSet=,              /* Required. Dataset to be summarized. */
  GroupVars=,            /* Required. Grouping variables (column). */
  StudyVars=,            /* Required. Studying variables (row). */
  StudyVarTypes=,        /* Required. Types of studying variables. Cat-categorical, Con-continue. */
  Statistics=            /* Descriptive statistics */
    N Nmiss Mean STD     /* Support all SAS proc means statistic-keyword(s) except CLM */
    Median Min Max       /* Mean STD, Min Max, P1 P99, LCLM UCLM, SKEW KURT, etc... */
    Q1 Q3,               /* will be automatically paired and positioned. */
  PaperSize=Letter,      /* A1-A5, LETTER, LEGAL, LEDGER, TABLOID, EXECUTIVE */
  Orientation=Landscape, /* Portrait, Landscape */
  FirstColumnWidth=1.2,  /* Inch */
  Margin=1,              /* Left and right margin in inch */
  GrandSubTotal=No,      /* Yes, No */
  ShowMissing=Yes,       /* Show missing of grouping variables. Yes, No */
  MaxDec=2,              /* Decimal */ 
  Style=2,               /* 1 Vertical Lines, 2 Horizontal Lines */
  CharFont=Arial,        /* Font for non-numeric charactors */
  NumFont=Courier,       /* Font for Numbers */
  FontSize=8,            /* pt, minimum 1 */ 
  RowShadow=2,           /* 0-none, 1-row, 2-block, 3-both  */
  Output_RTF_File="summary.rft");

  %macro name_list(name, count);
    %local i; %do i=1 %to &count; &name&i %end;
  %mend;

  %macro scan_(src, index, dlm);
    %* get a word from a list separated by sequenced dlm; 
    %* SAS build-in scan treats separater "abc" same as "cba";
    %local i j l temp;
    %let i=%index(&src, %eval(&index-1)&dlm);
    %let j=%index(&src, *&index&dlm);
    %let l=%length(%eval(&index-1)&dlm);
    %if %eval(&j-&i-&l)>0 %then %do;
    %substr(&src,%eval(&i+&l),%eval(&j-&i-&l)) %end;
  %mend;

  %macro get_word_count(words);
    %local i; %let i=1;
    %do %while (%scan(&words, &i) ne %str());
      %let i=%eval(&i+1);
    %end;%eval(&i-1)
  %mend;

  %macro list_map(list, pattern);
    %*fit the list to question mark position;
    %local i;%let i = 1;
    %do %while (%scan(&list., &i.) ne );
      %do; %unquote(%qsysfunc(tranwrd(&pattern,?,%nrstr(%scan(&list., &i.))))) %end;
      %let i = %eval((&i + 1);
    %end;
  %mend;

  %macro pairs_map(list, pattern);
    %local i p; %let i = 1;
    %do %while (%scan(&list., &i.) ne );
      %let p=%unquote(%qsysfunc(tranwrd(&pattern,?1,%nrstr(%scan(&list., &i.)))));
      %let p=%unquote(%qsysfunc(tranwrd(&p,?2,%nrstr(%scan(&list., %eval(&i+1))))));
        %do; &p %end;
      %let i = %eval((&i + 2);
    %end;
  %mend;

  %macro check_stats(stats, supt_stats);
    %local s i l stat; %let s=1;
    %do %while (%scan(&stats, &s) ne %str());
      %let stat=%scan(&stats, &s);
      %let l=%length(&stat);
      %let i=%index(%upcase(&supt_stats), %upcase(&stat));
        %if &i>0 %then %do; %substr(&supt_stats, &i, %length(&stat)) %end;
      %let s=%eval(&s+1);
    %end;
  %mend;

  %macro pair_stats(stats, pairs, retains, combines, drops);
    %let stats=%qcmpres(&stats)%str( ); 
    %local p1 p2 i j k r c d;
    %let k=1;
    %do %while (%scan(&pairs, &k) ne %str());
      %let p1=%scan(&pairs, &k); %let p2=%scan(&pairs, %eval(&k+1));
      %let i=%index(&stats, &p1%str( )); %let j=%index(&stats, &p2%str( ));
      %if &i>0 and &j>0 %then %do;
        %if &i<&j %then %do; 
          %let stats=%qsysfunc(tranwrd(&stats,&p2%str( ),%str())); 
          %let stats=%qsysfunc(tranwrd(&stats,&p1%str( ),%str(&p1 &p2 ))); 
        %end;
        %if &i>&j %then %do; 
          %let stats=%qsysfunc(tranwrd(&stats,&p1%str( ),%str())); 
          %let stats=%qsysfunc(tranwrd(&stats,&p2%str( ),%str(&p1 &p2 ))); 
        %end;
      %end;
      %let k=%eval(&k+2);
    %end; 
    %let k=1; %let r=&stats; 
    %do %while (%scan(&pairs, &k) ne %str());
      %let p1=%scan(&pairs, &k); %let p2=%scan(&pairs, %eval(&k+1));
      %let i=%index(&stats, &p1%str( )&p2%str( ));
      %if &i>0 %then %do;
        %let r=%qsysfunc(tranwrd(&r,&p1%str( )&p2%str( ),&p1%str(_)&p2%str( ))); 
        %let c=&c%str( )&p1%str( )&p2%str( );
        %let d=&d%str( )&p1%str( )&p2%str( );
      %end;
      %let k=%eval(&k+2);
    %end; 
    %let &retains=&r;
    %let &combines=&c;
    %let &drops=&d;
    %qcmpres(&stats)
  %mend;

  %macro get_var_info(DataSet, var, info_type, return);
    %local i libName dsName info;
    %let i=%index(&DataSet,.);
    %if &i=0 %then %do;%let libName=work;%let dsName=&DataSet;%end;
    %else %do;%let libName=%scan(&DataSet,1,.);%let dsName=%scan(&DataSet,2,.);%end;
    proc sql noprint;
      select &info_type into :info from sashelp.vcolumn
      where libname=upcase("&libName") and memname=upcase("&dsName")and upcase("&var")=upcase(name);
    quit;
    %let &return=&info;
  %mend;

  %macro get_var_layer_count(DataSet, var, return);
    %local count cmissing fmt;
    %get_var_info(&DataSet, &var, format, fmt);
    %if &fmt ne %then %do;
      proc sql noprint;
        select count(distinct(put(&var,&fmt))), nmiss(&var)
        into :count, :cmissing
        from &DataSet;
      quit;
    %end;
    %else %do;
      proc sql noprint;
        select count(distinct(&var)), nmiss(&var)
        into :count, :cmissing
        from &DataSet;
      quit;
      %if &cmissing = 0 %then %let count = &count;
      %else %let count = %eval(&count+1);
    %end;
    %let &return = &count;
  %mend;

  %macro get_var_layers(DataSet, var, return);
    %local miss values fmt;
    %get_var_info(&DataSet, &var, format, fmt);
    proc sql noprint;
      select (case when nmiss(&var) >0 then 1 else 0 end) 
        into :mis from &DataSet;
    quit;
    proc sql noprint;
      select (case when nmiss(&var) >0 then "Missing" else "" end) 
        into :miss from &DataSet;
      %if &fmt ne %then %do;
        select catx("",put(&var,&fmt),"*",monotonic()+&mis)
      %end;
      %else %do;
        select catx("",&var,"*",monotonic()+&mis)
      %end;
        into :values separated by "_sPt_" from 
        (select distinct(&var) from &DataSet)
      where &var is not missing order by &var;
    quit;
    %let values=&mis._sPt_&values._sPt_;
    %if &miss ne %then %let &return=0_sPt_&miss.*&values;%else %let &return=&values;
  %mend;

  %macro delete_temp_datasets(DSlist);
    proc datasets library=work nolist;
      %list_map(&DSlist, %str(delete ?)); run; 
  %mend;

  %macro type_str(total, current);
    %local i type;
    %if &current=0 %then %do i=1 %to &total;%let type=&type.0;%end;
    %else %do i=1 %to &total;
      %if &i <= &current %then %let type=&type.1;%else %let type=&type.0;
    %end;
    &type
  %mend;

  %macro type_strs(count);
    %local i type; %let type=;
    %do i=1 %to &count; %let type= &type.0;%end;"&type"
    %do i=1 %to &count; %let type=1&type;"%substr(&type, 1, &count)" %end;
  %mend;

  %macro create_total_count(DataSet, GroupVars);
    %local i j col_count col; %let col_count = %get_word_count(&GroupVars);
    proc freq data=&DataSet noprint;
      table %qsysfunc(tranwrd(&GroupVars, %str( ), %str(*))) /out=tmp_00_freq_out missing sparse;
    data tmp_00_freq_out(drop=count percent); retain level _type_ &GroupVars; 
      set tmp_00_freq_out end=eof;by &GroupVars;
      length level $2 total $32; array counts(%eval(&col_count)) _temporary_;
      level=%eval(&col_count+1); _type_= "%type_str(&col_count, &col_count)";total="(n=" || strip(count) ||")";output;
      %do i=%eval(&col_count-1)%to 1 %by -1;
        %let col=%scan(&GroupVars, &i);
        if first.&col then counts(&i)=0;
        counts(&i)+count;
        if last.&col then do;
          level=%eval(&i+1);
          _type_= "%type_str(&col_count, &i)";
          %do j=%eval(&i+1) %to &col_count;
            if vtype(%scan(&GroupVars, &j))="N" then %scan(&GroupVars, &j)=.;
            else %scan(&GroupVars, &j)="";
          %end;
          total="(n=" || strip(counts(&i)) ||")";output;
        end;
      %end;
      counts(&col_count)+count;
      if eof then do;
        level=1;
        _type_= "%type_str(&col_count, 0)";
        total="(n=" || strip(counts(&col_count)) ||")";
        %do i=1 %to &col_count;
          if vtype(%scan(&GroupVars, &i))="N" then %scan(&GroupVars, &i)=.;
            else %scan(&GroupVars, &i)="";
        %end;
        output;
      end;
    proc sort data=tmp_00_freq_out;by _type_ &GroupVars;
  %mend;

  %macro create_row_cnt_pct_level(DataSet, Row, RowCode, table, level, _type_, layers, MaxDec);
    proc freq data=&DataSet noprint;
      table &table &Row /out=tmp_01_freq_out_&RowCode._&level missing sparse;
    proc transpose data=tmp_01_freq_out_&RowCode._&level
      out=tmp_02_count_&RowCode._&level(rename=(col1-col&layers = &RowCode.cnt1-&RowCode.cnt&layers));
      by %qsysfunc(tranwrd(&table, *, %str( )));
      var count;
    data tmp_03_percent_&RowCode._&level; set tmp_02_count_&RowCode._&level;
      drop i total &RowCode.cnt1-&RowCode.cnt&layers; total=sum(of &RowCode.cnt1-&RowCode.cnt&layers);
      array cnt(&layers) &RowCode.cnt1-&RowCode.cnt&layers; array pct(&layers) &RowCode.pct1-&RowCode.pct&layers;
      if total = 0 or total =. then do i=1 to &layers; pct(i) = 0; end;
      else do i=1 to &layers; pct(i) = cnt(i) / total * 100; end;
    %local i cnt_pct;
    data tmp_04_n_pct_&RowCode._&level;
      merge tmp_02_count_&RowCode._&level tmp_03_percent_&RowCode._&level;
      %do i=1 %to &layers;
        %let cnt_pct=&RowCode.layer_%sysfunc(putn(&i,z3.));
        v=strip(put(&RowCode.pct&i,best.));
        if length(scan(v,2,"."))>&MaxDec then v=strip(put(&RowCode.pct&i,32.&MaxDec));
        &cnt_pct=strip(&RowCode.cnt&i) || BYTE(13) || " (" || strip(v) || "%)";
        drop v &RowCode.cnt&i &RowCode.pct&i;
      %end;
      level = " &level";_type_ = "&_type_";drop _name_ _label_;
  %mend;

  %macro create_row_cnt_pct(DataSet, Row, RowCode, GroupVars, order, MaxDec);
    %local i table layers layers_count col_count row_label;
    %let col_count=%get_word_count(&GroupVars);
    %get_var_layer_count(&DataSet, &Row, layers_count);
    %let i=1;/* first level;*/
    %create_row_cnt_pct_level(&DataSet, &Row, &RowCode, , &i, %type_str(&col_count, %eval(&i-1)), &layers_count, &MaxDec);
    %do %while (%scan(&GroupVars,&i) ne);* and other levels;
      %let table= &table %scan(&GroupVars,&i) *;
      %let i=%eval(&i+1);
      %create_row_cnt_pct_level(&DataSet, &Row, &RowCode, &table, &i, %type_str(&col_count, %eval(&i-1)), &layers_count, &MaxDec);
    %end;
    %get_var_layers(&DataSet, &Row, layers);
    %get_var_info(&DataSet, &Row, label, row_label);
    %if %length(&row_label)=0 %then %let row_label=&Row;
    data tmp_05_n_pct_&order;
      retain level _type_ &GroupVars &RowCode;
      set %name_list(tmp_04_n_pct_&RowCode._, &i);
      %do i=1 %to &layers_count;
        label &RowCode.layer_%sysfunc(putn(&i,z3.))
              =----------%scan_(&layers, &i, _sPt_);
      %end;
      &RowCode = "";
      label &RowCode = &row_label [n(%)];
  %mend;

  %macro create_cat_rows(DataSet, StudyVars, StudyVarTypes, GroupVars, MaxDec);
    %local Row i order type;
    %let i=1;%let order=0;
    %do %while (%scan(&StudyVarTypes,&i) ne);
      * only for categorical row variables;
      %let type= %scan(&StudyVarTypes, &i);
      %if %upcase(&type) = CAT %then %do;
        %let Row= %scan(&StudyVars, &i);
        * to avoid long dataset/variable name;
        %let RowCode= study_var_%sysfunc(putn(&i,z3.)); 
        %let order=%eval(&order+1);
        %let order=%sysfunc(putn(&order,z2.));
        %create_row_cnt_pct(&DataSet, &Row, &RowCode, &GroupVars, cat_&order, &MaxDec);
      %end;
      %let i=%eval(&i+1);
    %end;
    %if &order^=0 %then %do;data tmp_06_rows_cat; merge 
      %do i=1 %to &order; tmp_05_n_pct_cat_%sysfunc(putn(&i,z2.)) %end;; %end;
  %mend;

  %macro create_row_means(DataSet, Row, RowCode, GroupVars, order, Statistics, MaxDec);
    %local i out_var_names var_name row_label supt_stats pairs rs cs ds;
    %let supt_stats=N T Nmiss Mean STD Median Min Max Sum StdErr Range Mode USS VAR QRange PROBT CSS
      Q1 Q3 P1 P5 P10 P20 P25 P30 P40 P50 P60 P70 P75 P80 P90 P95 P99 CV LCLM UCLM SKEW KURT;
    %let pairs=Mean STD Min Max LCLM UCLM Q1 Q3 P1 P99 P5 P95 P10 P90 P20 P80 P30 P70 P40 P60 P25 P75 SKEW KURT;
    %let Statistics=%qsysfunc(tranwrd(&Statistics,STDDEV,STD));
    %let Statistics=%qsysfunc(tranwrd(&Statistics,PRT,PROBT));
    %let Statistics=%qsysfunc(tranwrd(&Statistics,KURTOSIS,KURT));
    %let Statistics=%qsysfunc(tranwrd(&Statistics,SKEWNESS,SKEW));
    %let Statistics=%check_stats(&Statistics, &supt_stats);
    %let Statistics=%pair_stats(&Statistics, &pairs, rs, cs, ds);
    %let out_var_names=%list_map(&Statistics, %str(?=&RowCode._?));
    proc means data=&DataSet chartype noprint missing;
      output out=tmp_070_mean_&order &out_var_names;
      class &GroupVars;
      var &Row;
    %get_var_info(&DataSet, &Row, label, row_label);
    %if %length(&row_label)=0 %then %let row_label=&Row;
    data tmp_071_mean_&order;
      retain &GroupVars _type_ &RowCode;
      set tmp_070_mean_&order;
      &RowCode = "";
      label &RowCode=&row_label %list_map(&Statistics, %str(&RowCode._?=----------?));
      drop _freq_;
      where _type_ in (%type_strs(%get_word_count(&GroupVars)));
    data tmp_072_mean_&order;
      set tmp_071_mean_&order(rename=(%list_map(&Statistics, %str(&RowCode._?=_&RowCode._?))));
      %let i=1;
      %do %while (%scan(&Statistics,&i) ne);
        %let var_name=&RowCode._%scan(&Statistics,&i);
        &var_name=strip(put(_&var_name,best.));
        if length(scan(&var_name,2,"."))>&MaxDec then &var_name=strip(put(_&var_name,32.&MaxDec));
        if &var_name="." then &var_name="";
        drop _&var_name;
        %let lbl=%scan(&Statistics,&i);
        %* QRange -> IQR;
        %if %upcase(&lbl)=QRANGE %then %let lbl=IQR;
        label &var_name=----------&lbl;
        %let i=%eval(&i+1);
      %end;

    %let strCombinePattern=%nrstr(&RowCode._?1_?2=" (" || strip(&RowCode._?1) || BYTE(13) || ", " || strip(&RowCode._?2) || ")"; label &RowCode._?1_?2=----------(?1, ?2););
    %* Mean STD -> Mean(SD);
    %let strMeanSD=%nrstr(&RowCode._?1_?2=strip(&RowCode._?1) || BYTE(13) || " (" || strip(&RowCode._?2) || ")"; label &RowCode._?1_?2=----------Mean (SD););
    %* LCLM UCLM -> 95% CI;
    %let strCI95=%nrstr(&RowCode._?1_?2=" (" || strip(&RowCode._?1) || BYTE(13) || ", " || strip(&RowCode._?2) || ")"; label &RowCode._?1_?2=----------%nrstr(95% CI););
    %let strPairRename=%qsysfunc(tranwrd(%pairs_map(&cs,&strCombinePattern),%pairs_map(Mean STD,&strCombinePattern),%pairs_map(Mean STD,&strMeanSD)));
    %let strPairRename=%qsysfunc(tranwrd(&strPairRename,%pairs_map(LCLM UCLM,&strCombinePattern),%pairs_map(LCLM UCLM,&strCI95)));

    data tmp_07_mean_&order(drop=%list_map(&ds, %str(&RowCode._?)));
      retain &GroupVars _type_ &RowCode %list_map(&rs, %str(&RowCode._?));
/*      set tmp_072_mean_&order(rename=(%list_map(&Statistics, %str(&RowCode._?=_&RowCode._?))));*/
      set tmp_072_mean_&order;
      %unquote(&strPairRename);
  %mend;

  %macro create_con_rows(DataSet, StudyVars, StudyVarTypes, GroupVars, Statistics, MaxDec);
    %local Row i order type;
    %let i=1;%let order=0;
    %do %while (%scan(&StudyVarTypes,&i) ne);
      * only for continue row variables;
      %let type= %scan(&StudyVarTypes, &i);
      %if %upcase(&type) = CON %then %do;
        %let Row= %scan(&StudyVars, &i);
        %let RowCode= study_var_%sysfunc(putn(&i,z3.)); 
        %let order=%eval(&order+1);
        %let order=%sysfunc(putn(&order,z2.));
        %create_row_means(&DataSet, &Row, &RowCode, &GroupVars, con_&order, &Statistics, &MaxDec);
      %end;
      %let i=%eval(&i+1);
    %end;
    %if &order^=0 %then %do;data tmp_08_rows_con; merge 
      %do i=1 %to &order; tmp_07_mean_con_%sysfunc(putn(&i,z2.)) %end; %end;
  %mend;

  %macro rename_num_cols(col_names, col_types);
    %local i col_name;
    %let i=1;
    %do %while (%scan(&col_names, &i) ne);
      %if %scan(&col_types, &i) = num %then %do;
        %let col_name=%scan(&col_names, &i);
        &col_name=_num_&i
      %end;
      %let i=%eval(&i+1);
    %end;
  %mend;

  %macro num_to_chr(ds, out);
    %* need macro rename_num_cols, scan_;
    %local i lib col_names col_types col_labels col_formats dlm;
    %let i=%index(&ds,.); %let dlm=_sPt_;
    %if &i=0 %then %let lib=work;
    %else %do;%let lib=%scan(&ds,1,.);%let ds=%scan(&ds,2,.);%end;
    proc sql noprint;
      select strip(name), strip(type), 
        catx("",label,"*",monotonic()), 
        catx("",format,"*",monotonic())
      into :col_names separated by " ", :col_types separated by " ",
           :col_labels separated by "&dlm", :col_formats separated by "&dlm"
      from sashelp.vcolumn
      where libname="%upcase(&lib)" and memtype="DATA" and memname="%upcase(&ds)";
    quit;
    %let col_labels=0&dlm&col_labels&dlm;
    %let col_formats=0&dlm&col_formats&dlm;
    data &out;
      retain &col_names;
      set &ds(rename=(%rename_num_cols(&col_names, &col_types)));
      %local col_name label format;
      %let i=1;
      %do %while (%scan(&col_names, &i) ne);
        %let col_name=%scan(&col_names, &i);
        %let col_label = %scan_(&col_labels, &i, &dlm);
        label &col_name = &col_label;
        %if %scan(&col_types, &i) = num %then %do;
          %let format=%scan_(&col_formats, &i, &dlm);
          %if &format eq %then %let format=best.;
          &col_name=strip(put(_num_&i, &format));
          if &col_name="." then &col_name="";
          drop _num_&i;
        %end;
        %let i=%eval(&i+1);
      %end;
    run;
  %mend;

  %macro convert_num_chr(DataSet, GroupVars, GrandSubTotal, ShowMissing, Style);
    %local i col tbls;
    %let nGroupVars=%get_word_count(&GroupVars);
    data _null_;X=Sleep(0.1,1);run;
    %if %sysfunc(exist(tmp_06_rows_cat)) %then %let tbls=tmp_06_rows_cat;
    %if %sysfunc(exist(tmp_08_rows_con)) %then %let tbls=&tbls. tmp_08_rows_con;
    data tmp_09_rows_cat_con_num;retain level _type_ &GroupVars blank;
      merge tmp_00_freq_out &tbls;by _type_ &GroupVars;
    proc sort data=tmp_09_rows_cat_con_num;by &GroupVars _type_;
    %num_to_chr(tmp_09_rows_cat_con_num, tmp_10_rows_cat_con_chr);
    proc sort data=tmp_10_rows_cat_con_chr;by &GroupVars _type_;
    data tmp_10_rows_cat_con_chr(drop=i);
      length &GroupVars $256;
      set tmp_10_rows_cat_con_chr;
      array cols(*) &GroupVars;
      array gvars(&nGroupVars)$ gvar1-gvar%get_word_count(&GroupVars);
      %do i=1 %to &nGroupVars;
        %let col=%scan(&GroupVars, &i);
        &col=putc(&col,vformat(&col));
      %end;
      do i=1 to dim(cols);
        gvars(i)=cols(i);
        if substr(_type_,i,1)="0" then cols(i)="Total";
        /* delete grand total & subtotal if not lowest level*/
        %if %upcase(&GrandSubTotal)=NO %then %do;
          if substr(_type_,i,1)="0" and level < &nGroupVars then delete;
        %end;
        if substr(_type_,i,1)="1" and cols(i)="" then cols(i)="Missing";
        if cols(i)="Missing" and total="(n=0)" then delete;
        %if %upcase(&ShowMissing)=NO %then %do;
          if cols(i)="Missing" then delete;
        %end;
      end;
    %let levels = &nGroupVars;
    %if &style=2 %then %do;
      %if &levels>1 %then %do;
        data tmp_10_rows_cat_con_chr; set tmp_10_rows_cat_con_chr end=eof; by &GroupVars _type_ notsorted;
          array ary(*) _character_; output; added=0; drop i added;
          %do i=1 %to %eval(&levels-1);
            %let col=%scan(&GroupVars, &i);
            if _n_ > 1 then do;
              if last.&col and eof=0 then do;
                do i=%eval(&i+0) to dim(ary); ary(i)=""; end;
                total = "000"; level="-" || "&i"; 
                if added=0 then output; added=1;
              end;
            end;
          %end;
      %end; 
    %end;
    proc transpose data=tmp_10_rows_cat_con_chr(drop=&GroupVars level _type_ total gvar1-gvar%get_word_count(&GroupVars))
      out=tmp_110_summary(rename=(_name_=VarCodes _label_=Variables));
      var _all_;
    data tmp_111_summary; set tmp_110_summary;
      attrib _all_ label="";
      study_var=substr(VarCodes,1,13); row_id=_n_;
    proc sort data=tmp_111_summary; by study_var row_id;
    data tmp_112_summary; set tmp_111_summary;
      if Variables="" then Variables=VarCodes;
      if length(VarCodes)=13 then do; VarType=1;VarGroup+1; end;
      else VarType=2;
    data tmp_11_summary;set tmp_112_summary(drop=VarCodes study_var row_id);by VarGroup;
      if first.Vargroup then VarOrder=0;
      VarOrder+1;
      length pvalue $200; pvalue=""; label pvalue='P-Value';
  %mend;

  %macro save_options;
    %local i opt opt_list; %let i=1;
    %let opt_list=source source2 notes mprint mlogic symbolgen quotelenmax;
    %do %while (%scan(&opt_list, &i) ne);
      %let opt=%scan(&opt_list, &i);
      %*sysfunc(getoption(&opt,keyword));
      %sysfunc(getoption(&opt))
      %let i=%eval(&i+1);
    %end;
  %mend;

  %macro change_options;
    options nosource nosource2 nonotes nomprint nomlogic nosymbolgen options noquotelenmax;;
    /*dm 'odsresults; clear';*/
  %mend;

  %macro add_pvalue(DataSet, StudyVar, GroupVar, Type, Order);
    %if %upcase(&Type)=CAT %then %do;
      proc freq data=&DataSet;
        ods output "CHI-SQUARE TESTS"=tmp_pvalue;
        tables &StudyVar * &GroupVar/chisq;
      run; quit;
      ods output close; 
      data tmp_pvalue(keep=pvalue VarGroup VarOrder Test);
        set tmp_pvalue; length TestDesc $200;
        if upcase(STATISTIC)="CHI-SQUARE" ;
        if prob<0.001 then pvalue="<.001"; else pvalue=put(prob,5.3);
        VarGroup=&Order; VarOrder=2; Test="P"; 
      run; quit;
    %end;
    %if %upcase(&Type)=CON %then %do;
      proc glm data=&DataSet PLOTS(MAXPOINTS=NONE) ;
        ods output "Overall ANOVA" = tmp_pvalue;
        class &GroupVar;
        model &StudyVar=&GroupVar;
      run;quit;
      ods output close;
      data tmp_pvalue(keep=pvalue VarGroup VarOrder Test);
        set tmp_pvalue; length TestDesc $200;
        if upcase(Source)="MODEL"; 
        if probf<0.001 then pvalue="<.001"; else pvalue=put(probf,5.3);
        VarGroup=&Order; VarOrder=2; Test="A"; 
      run; quit;
    %end;
    proc append base=tmp_pvalues data=tmp_pvalue;
  %mend;

  %macro add_pvalues(DataSet, StudyVars, StydyVarTypes, GroupVars);
    %local order studyvar groupvar type ;
    %let order=1; %let groupvar=%scan(&GroupVars,1); 
    ods listing close; ods results=off;
    %do %while (%scan(&StudyVarTypes,&order) ne);
      %let type= %scan(&StudyVarTypes, &order);
      %let studyvar= %scan(&StudyVars, &order);
      %add_pvalue(&DataSet, &studyvar, &groupvar, &type, &order);
      %let order=%eval(&order+1);
    %end; 
    ods results=on; ods listing;
    data tmp_11_summary; merge tmp_11_summary tmp_pvalues; by VarGroup VarOrder;
  %mend;

  %macro bw(levels, level);
    %eval((&levels-&level+1)*2-2)
  %mend;

  %macro bd(position,width,style,color);
    %if (&color=) %then %let color=black; %if (&color=w) %then %let color=white;
    %if (&position=t) %then %let position=top; %if (&position=b) %then %let position=bottom;
    %if (&position=l) %then %let position=left; %if (&position=r) %then %let position=right;
    %if (&style=) %then %let style=solid; %if (&style=n) %then %let style=none;
    %if &width=0 %then %let color=white;
    border&position.width=&width border&position.color=&color border&position.style=&style
  %mend;

  %macro PaperWidth(PaperSize, Orientation);
    %if %upcase(&PaperSize)=LETTER and %upcase(&Orientation)=PORTRAIT %then 8.5;
    %if %upcase(&PaperSize)=LETTER and %upcase(&Orientation)=LANDSCAPE %then 11;
    %if %upcase(&PaperSize)=LEGAL and %upcase(&Orientation)=PORTRAIT %then 8.5;
    %if %upcase(&PaperSize)=LEGAL and %upcase(&Orientation)=LANDSCAPE %then 14;
    %if %upcase(&PaperSize)=LEDGER and %upcase(&Orientation)=PORTRAIT %then 11;
    %if %upcase(&PaperSize)=LEDGER and %upcase(&Orientation)=LANDSCAPE %then 17;
    %if %upcase(&PaperSize)=TABLOID and %upcase(&Orientation)=PORTRAIT %then 11;
    %if %upcase(&PaperSize)=TABLOID and %upcase(&Orientation)=LANDSCAPE %then 17;
    %if %upcase(&PaperSize)=EXECUTIVE and %upcase(&Orientation)=PORTRAIT %then 7.25;
    %if %upcase(&PaperSize)=EXECUTIVE and %upcase(&Orientation)=LANDSCAPE %then 10.5;
    %if %upcase(&PaperSize)=A1 and %upcase(&Orientation)=PORTRAIT %then 23.39;
    %if %upcase(&PaperSize)=A1 and %upcase(&Orientation)=LANDSCAPE %then 33.11;
    %if %upcase(&PaperSize)=A2 and %upcase(&Orientation)=PORTRAIT %then 16.54;
    %if %upcase(&PaperSize)=A2 and %upcase(&Orientation)=LANDSCAPE %then 23.39;
    %if %upcase(&PaperSize)=A3 and %upcase(&Orientation)=PORTRAIT %then 11.69;
    %if %upcase(&PaperSize)=A3 and %upcase(&Orientation)=LANDSCAPE %then 16.54;
    %if %upcase(&PaperSize)=A4 and %upcase(&Orientation)=PORTRAIT %then 8.27;
    %if %upcase(&PaperSize)=A4 and %upcase(&Orientation)=LANDSCAPE %then 11.69;
    %if %upcase(&PaperSize)=A5 and %upcase(&Orientation)=PORTRAIT %then 5.83;
    %if %upcase(&PaperSize)=A5 and %upcase(&Orientation)=LANDSCAPE %then 8.27;
  %mend;

  %macro output_rtf_style1(DataSet, Output_RTF_File, head_str, tot_str, groups, levels, 
    bdwidths,cellwidths,PaperSize,Orientation,Margin,FirstColumnWidth,
    CharFont,NumFont,FontSize,RowShadow);
    %local w wb wc wm s ss sd vs ps SumDS cellspacing cellwidth LastColumnWidth clr1 clr2 clr3 clr4;
    %let ss=%nrstr(/*font_face="&CharFont"*/ font_size=&FontSize.pt);
    proc sql noprint; select distinct Test 
      into: str_tests separated by "*" 
      from tmp_11_summary; quit;
    %let w=%bw(&levels,1); %let cellspacing=0; %let LastColumnWidth=0.7;
    %let wm=%eval(%bw(&levels, 1)+1); %if &wm<4 %then %let wm=4;
    %let SumDS = tmp_11_summary;
    %let cellwidth=%sysevalf((%PaperWidth(&PaperSize,&Orientation)-&FirstColumnWidth
         -&cellspacing*&groups-&margin*2-&LastColumnWidth-0.1)/&groups);
    options papersize=&PaperSize topmargin=0.1in bottommargin=0.5in leftmargin=&Margin.in rightmargin=&Margin.in
            orientation=&Orientation nodate nonumber;
    ods escapechar='^'; footnote; title;
    filename fname &Output_RTF_File;
    ods listing close; ods rtf file=fname;
    title1 justify=l "&DataSet._Summary";
    footnote j=r "^{pageof}";
    proc report data=&SumDS nowd
      style(column)={just=l vjust=middle font_size=&FontSize.pt borderwidth=0 cellpadding=3}
      style(header)={rules=none vjust=bottom font_size=&FontSize.pt 
                     font_face="&CharFont" foreground=black background=white cellpadding=3}
      style(report)={rules=none frame=void cellspacing=&cellspacing.in borderwidth=&wm};
      %let s='^S={%bd(,0) %bd(l,&w)}'; %let ps=pvalue;
      %do i=2 %to &levels; %let ps=(&s &ps); %end; %let ps=('^S={%bd(,0) %bd(t,&wm) %bd(l,&w)}' &ps);
      %let s='^S={%bd(,0) %bd(r,&w)}'; %let vs=Variables;
      %do i=2 %to &levels; %let vs=(&s &vs); %end; %let vs=('^S={%bd(,0) %bd(t,&wm) %bd(r,&w)}' &vs);
      cols &vs &head_str &ps VarGroup VarType VarOrder Test temp;
      *define VarGroup / order noprint;
      define Variables / display  'Variable'
        style(column)={cellwidth=&FirstColumnWidth.in font_face='Arial' %bd(r,&w)}
        style(header)={%bd(r,&w) %bd(b,3.5) just=l};
      define pvalue / display 'P-Value'
        style(column)={font_face="&CharFont" %bd(l,&w) just=r cellwidth=&LastColumnWidth.in}
        style(header)={%bd(l,&w) %bd(b,3.5) just=r};
      define VarGroup / display noprint;
      define VarType / display noprint;
      define VarOrder / display noprint;
      define Test / display noprint;
      %do i=1 %to &groups;
        %let wb=%scan(&bdwidths,&i);
        %let wc=%sysevalf(%scan(&cellwidths,&i)+0.01);
        %let tot=%scan(&tot_str, &i); 
        %if %substr(&tot,1,1) eq %str(n) %then %do;
          define col&i / display "(&tot)"
        %end;
        %else %do;
          define col&i / display ""
        %end;
        style(column)={%bd(l,&wb)
          rules=cols just=c cellwidth=%sysevalf(&cellwidth * &wc)in font_face="&NumFont"}
        style(header)={%bd(l,&wb) %bd(b,3.5)};
      %end;
      define temp / computed noprint;
      compute temp;
        if substr(Variables, 1, 10) eq "----------" then
          Variables=substr(Variables,11);
        if VarOrder=2 then pvalue=strip(pvalue) || "^S={%bd(,1,n) font_weight=bold} ^{super " || strip(Test) || "}";
        %let clr1=cxffffff; %let clr2=cxffffff; %let clr3=cxffffff; %let clr4=cxffffff;
        %if &RowShadow=1 %then %do;
          %let clr1=cxffffff; %let clr2=cxefefef; %let clr3=cxffffff; %let clr4=cxefefef;
        %end;
        %if &RowShadow=2 %then %do;
          %let clr1=cxffffff; %let clr2=cxffffff; %let clr3=cxefefef; %let clr4=cxefefef;
        %end;
        %if &RowShadow=3 %then %do;
          %let clr1=cxffffff; %let clr2=cxefefef; %let clr3=cxefefef; %let clr4=cxe0e0e0;
        %end;
        if VarType=1 and mod(VarGroup,2)=1 then 
          call define(_row_,'style','style={background=&clr3 font_weight=bold %bd(t,3.5)}');
        if VarType=1 and mod(VarGroup,2)=0 then 
          call define(_row_,'style','style={background=&clr1 font_weight=bold %bd(t,3.5)}');
        if VarType=2 and mod(VarGroup,2)=1 and mod(VarOrder,2)=1 then 
          call define(_row_,'style','style={background=&clr3}'); 
        if VarType=2 and mod(VarGroup,2)=1 and mod(VarOrder,2)=0 then 
          call define(_row_,'style','style={background=&clr4}'); 
        if VarType=2 and mod(VarGroup,2)=0 and mod(VarOrder,2)=1 then 
          call define(_row_,'style','style={background=&clr1}'); 
        if VarType=2 and mod(VarGroup,2)=0 and mod(VarOrder,2)=0 then 
          call define(_row_,'style','style={background=&clr2}'); 
        if VarType=2 then
          call define('_C1_','style','style={fontstyle=italic indent=20}'); 
      endcomp;
      compute after _page_/left;
        line "^S={%bd(t,6)}";  
        %let i=1;
        %do %while (%scan(&str_tests,&i,*) ne);
          %let s=%scan(&str_tests,&i,*); 
          %if &s=P %then %let sd=Pearson Chi-Square test for top level groups;
          %if &s=A %then %let sd=Analysis of Variance for top level groups;
          line "^S={&ss font_weight=bold} ^{super &s} ^S={&ss} &sd";
          line ""; 
          %let i=%eval(&i+1);
        %end;
      endcomp;
    run;
    ods rtf close; ods listing;
  %mend;

  %macro output_style1(DataSet, GroupVars, Output_RTF_File,PaperSize,Orientation,
    Margin,FirstColumnWidth,CharFont,NumFont,FontSize,RowShadow);
    %local i j n s s1 s2 col levels groups w wt wm bdwidths cellwidths clr colors head_str tot_str ctlTable;
    %let ctlTable=tmp_10_rows_cat_con_chr;
    %let levels = %get_word_count(&GroupVars);
    data _null_; set &ctlTable; by &GroupVars _type_ notsorted;
      call symput("groups", strip(_n_)); w=0; if level="0" then cw=0; else cw=1;
      %do i =1 %to &levels;
        %let col=%scan(&GroupVars, &i);
        if first.&col then do; w1=%bw(&levels, &i); if w1>w then w=w1; end;
        
        if &col="Missing" then call symput("colors", resolve('&colors') || " Red");
        else call symput("colors", resolve('&colors') || " Black");
      %end;
      call symput("bdwidths", resolve('&bdwidths') || " " || strip(w));
      call symput("cellwidths", resolve('&cellwidths') || " " || strip(cw));
    data _null_; set tmp_10_rows_cat_con_chr; by &GroupVars _type_ notsorted;
      call symput("n",strip(_n_));
      %do i =1 %to &levels;
        %let col=%scan(&GroupVars, &i); 
        %let s1="bdstr1_" || strip(&i) || "_" || resolve('&n');
        %let s2="bdstr2_" || strip(&i) || "_" || resolve('&n');
        call symput("value", strip(&col));
        lbl = vlabel(&col); lbl="";
        if first.&col then do;
          %if &i <= &levels %then %do;
            call symput("head_str", resolve('&head_str') || " ('^S={" || &s1 || "}" ||
             /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
          %end;
        end;
      %end;
      call symput("head_str", resolve('&head_str') || " Col" || strip(_n_));
      call symput("tot_str", resolve('&tot_str') || " " || strip(total));
      %do i =&levels %to 1 %by -1;
        %let col=%scan(&GroupVars, &i);
        if last.&col then do; 
          %if &i <= &levels %then %do;
            call symput("head_str", resolve('&head_str') || ")"); 
          %end;
        end;
      %end;
    run;
    %let wm=%eval(%bw(&levels, 1)+1); %if &wm<4 %then %let wm=4;
    %do j=&groups %to 1 %by -1;
      %do i=&levels %to 1 %by -1;
        %let clr=%scan(&colors, %eval(&levels*(&j-1)+&i);
        %let w=%scan(&bdwidths,&j); %let s=bdstr1_&i._&j; 
        %if &i=1 %then %let wt=&wm; %else %let wt=%eval(%bw(&levels, &i)+2);
        %let d=%bd(b,0) %bd(r,0) %bd(t,&wt) %bd(l,&w) foreground=&clr;
        %let head_str=%qsysfunc(tranwrd(&head_str,&s,&d));
        %let s=bdstr2_&i._&j;
        %let d=%bd(,0) %bd(l,&w) foreground=&clr;
        %let head_str=%qsysfunc(tranwrd(&head_str,&s,&d));
      %end;
    %end;
    %let head_str=%unquote(&head_str);
    %output_rtf_style1(&DataSet, &Output_RTF_File,&head_str,&tot_str,&groups,&levels,
      &bdwidths,&cellwidths,&PaperSize,&Orientation,&Margin,&FirstColumnWidth,
      &CharFont,&NumFont,&FontSize,&RowShadow);
  %mend;

  %macro output_rtf_style2(DataSet, Output_RTF_File, head_str, tot_str, groups, levels, 
    bdwidths,cellwidths,blankcols,PaperSize,Orientation,Margin,FirstColumnWidth,
    CharFont,NumFont,FontSize,RowShadow);
    %local b w wc wm s ss sd vs ps SumDS cellspacing cellwidth blankwidth 
           LastColumnWidth clr1 clr2 clr3 clr4 str_tests;
    proc sql noprint; select distinct Test 
      into: str_tests separated by "*" 
      from tmp_11_summary; quit;
    %let ss=%nrstr(/*font_face="&CharFont"*/ font_size=&FontSize.pt);
    %let w=%bw(&levels,1); %let cellspacing=0; %let LastColumnWidth=0.7; %let blankwidth=0.1;
    %let SumDS = tmp_11_summary;
    %let cellwidth=%sysevalf((%PaperWidth(&PaperSize,&Orientation)-&FirstColumnWidth
       -&blankwidth*&blankcols-&margin*2-&LastColumnWidth-0.1)/(&groups-&blankcols));
    options papersize=&PaperSize topmargin=0.1in bottommargin=0.5in leftmargin=&Margin.in rightmargin=&Margin.in
            orientation=&Orientation nodate nonumber;
    ods escapechar='^'; footnote; title;
    filename fname &Output_RTF_File;
    ods listing close; ods rtf file=fname startpage=no;
    title1 justify=l "&DataSet._Summary";
    footnote j=r "^{pageof}";
    proc report data=&SumDS nowd
      style(column)={just=l vjust=middle font_size=&FontSize.pt borderwidth=0 cellpadding=3}
      style(header)={rules=none vjust=bottom font_size=&FontSize.pt 
                     font_face="&CharFont" foreground=black background=white cellpadding=3}
      style(report)={rules=none frame=void cellspacing=&cellspacing.in borderwidth=6};
      %let s='^S={}'; %let ps=pvalue;
      %do i=2 %to &levels; %let ps=(&s &ps); %end; %let ps=('^S={%bd(t,6)}' &ps);
      %let s='^S={}'; %let vs=Variables;
      %do i=2 %to &levels; %let vs=(&s &vs); %end; %let vs=('^S={%bd(t,6)}' &vs);

      cols &vs &head_str &ps VarGroup VarType VarOrder Test temp;
      define Variables / display  'Variable'
        style(column)={cellwidth=&FirstColumnWidth.in font_face='Arial'}
        style(header)={%bd(b,3.5) just=l};
      define pvalue / display 'P-Value'
        style(column)={font_face="&CharFont" just=r cellwidth=&LastColumnWidth.in}
        style(header)={%bd(b,3.5) just=r};
      define VarGroup / display noprint;
      define VarType / display noprint;
      define VarOrder / display noprint;
      define Test / display noprint;
      %do i=1 %to &groups;
        %let b=%scan(&cellwidths,&i);
        %if &b=2 %then %let wc=&cellwidth; %else %let wc=&blankwidth;
        %let tot=%scan(&tot_str, &i); 
        %if %substr(&tot,1,1) eq %str(n) %then %do;
          define col&i / display "(&tot)"
        %end;
        %else %do;
          define col&i / display ""
        %end;
        style(column)={
          rules=cols just=c cellwidth=&wc.in font_face="&NumFont"}
          style(header)={%bd(b,3.5)};
      %end;
      define temp / computed noprint;
      compute temp;
        if substr(Variables,1,10)="----------" then 
          Variables=substr(Variables,11);
        if VarOrder=2 then pvalue=strip(pvalue) || "^S={%bd(,1,n) font_weight=bold} ^{super " || strip(Test) || "}";
        %let clr1=cxffffff; %let clr2=cxffffff; %let clr3=cxffffff; %let clr4=cxffffff;
        %if &RowShadow=1 %then %do;
          %let clr1=cxffffff; %let clr2=cxefefef; %let clr3=cxffffff; %let clr4=cxefefef;
        %end;
        %if &RowShadow=2 %then %do;
          %let clr1=cxffffff; %let clr2=cxffffff; %let clr3=cxefefef; %let clr4=cxefefef;
        %end;
        %if &RowShadow=3 %then %do;
          %let clr1=cxffffff; %let clr2=cxefefef; %let clr3=cxefefef; %let clr4=cxe0e0e0;
        %end;
        if VarType=1 and mod(VarGroup,2)=1 then 
          call define(_row_,'style','style={background=&clr3 font_weight=bold %bd(t,3.5)}');
        if VarType=1 and mod(VarGroup,2)=0 then 
          call define(_row_,'style','style={background=&clr1 font_weight=bold %bd(t,3.5)}');
        if VarType=2 and mod(VarGroup,2)=1 and mod(VarOrder,2)=1 then 
          call define(_row_,'style','style={background=&clr3}'); 
        if VarType=2 and mod(VarGroup,2)=1 and mod(VarOrder,2)=0 then 
          call define(_row_,'style','style={background=&clr4}'); 
        if VarType=2 and mod(VarGroup,2)=0 and mod(VarOrder,2)=1 then 
          call define(_row_,'style','style={background=&clr1}'); 
        if VarType=2 and mod(VarGroup,2)=0 and mod(VarOrder,2)=0 then 
          call define(_row_,'style','style={background=&clr2}'); 
        if VarType=2 then
          call define('_C1_','style','style={fontstyle=italic indent=20}'); 
      endcomp;
      compute after _page_/left;
        line "^S={%bd(t,6)}";  
        %let i=1;
        %do %while (%scan(&str_tests,&i,*) ne);
          %let s=%scan(&str_tests,&i,*); 
          %if &s=P %then %let sd=Pearson Chi-Square test for top level groups;
          %if &s=A %then %let sd=Analysis of Variance for top level groups;
          line "^S={&ss font_weight=bold} ^{super &s} ^S={&ss} &sd";
          line ""; 
          %let i=%eval(&i+1);
        %end;
      endcomp;
    run;
    ods rtf close; ods listing;
  %mend;

  %macro output_style2(DataSet, GroupVars, Output_RTF_File,PaperSize,Orientation,
    Margin,FirstColumnWidth,CharFont,NumFont,FontSize,RowShadow);
    %local i j n s s1 s2 s3 s4 s5 col levels groups w wt wm bdwidths cellwidths blankcols clr colors head_str tot_str ctlTable;
    %let ctlTable=tmp_10_rows_cat_con_chr;
    %let levels = %get_word_count(&GroupVars);
    data _null_; set &ctlTable; by &GroupVars _type_ notsorted;
      call symput("groups", strip(_n_)); w=0; 
      if substr(level,1,1)="-" then do; cw=1; blankcols+1; end; else cw=2;
      %do i =1 %to &levels;
        %let col=%scan(&GroupVars, &i);
        if first.&col then do; w1=%bw(&levels, &i); if w1>w then w=w1; end;
        
        if &col="Missing" then call symput("colors", resolve('&colors') || " Red");
        else call symput("colors", resolve('&colors') || " Black");
      %end;
      call symput("bdwidths", resolve('&bdwidths') || " " || strip(w));
      call symput("cellwidths", resolve('&cellwidths') || " " || strip(cw));
      call symput("blankcols", strip(blankcols));
    data _null_; set tmp_10_rows_cat_con_chr; by &GroupVars _type_ notsorted;
      call symput("n",strip(_n_));
      %do i =1 %to &levels;
        %let col=%scan(&GroupVars, &i); 
        %let s1="bdstr1_" || strip(&i) || "_" || resolve('&n');
        %let s2="bdstr2_" || strip(&i) || "_" || resolve('&n');
        %let s3="bdstr3_" || strip(&i) || "_" || resolve('&n');
        %let s4="bdstr4_" || strip(&i) || "_" || resolve('&n');
        %let s5="bdstr5_" || strip(&i) || "_" || resolve('&n');
        call symput("value", strip(&col));
        lbl = vlabel(&col); lbl="";
        if first.&col then do;
          %if &i = 1 and &levels = 1 %then %do;
              call symput("head_str", resolve('&head_str') || " ('^S={" || &s2 || "}" ||
                /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
          %end;
          %if &i = 1 and &levels > 1 %then %do;
            if level > 0  then
              call symput("head_str", resolve('&head_str') || " ('^S={" || &s1 || "}" ||
                /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
            else
              call symput("head_str", resolve('&head_str') || " ('^S={" || &s2 || "}" ||
                /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
          %end;
          %if &i > 1 and &i < &levels %then %do;
            if level > 0  then
              call symput("head_str", resolve('&head_str') || " ('^S={" || &s3 || "}" ||
                /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
            else
              call symput("head_str", resolve('&head_str') || " ('^S={" || &s5 || "}" ||
                /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
          %end;
          %if &i > 1 and &i = &levels %then %do;
            if level > 0  then
              call symput("head_str", resolve('&head_str') || " ('^S={" || &s4 || "}" ||
                /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
            else
              call symput("head_str", resolve('&head_str') || " ('^S={" || &s5 || "}" ||
                /* strip(lbl) || " = " ||*/ resolve('&value') || "'");
          %end; 
        end;
      %end;
      call symput("head_str", resolve('&head_str') || " Col" || strip(_n_));
      call symput("tot_str", resolve('&tot_str') || " " || strip(total));
      %do i =&levels %to 1 %by -1;
        %let col=%scan(&GroupVars, &i);
        if last.&col then do; 
          %if &i <= &levels %then %do;
            call symput("head_str", resolve('&head_str') || ")"); 
          %end;
        end;
      %end;
    run;
    %do j=&groups %to 1 %by -1;
      %do i=&levels %to 1 %by -1;
        %let clr=%scan(&colors, %eval(&levels*(&j-1)+&i);
        %let s=bdstr1_&i._&j;
        %let d=%bd(l,0) %bd(r,0) %bd(t,6) %bd(b,3.5) foreground=&clr;
        %let head_str=%qsysfunc(tranwrd(&head_str,&s,&d));
        %let s=bdstr2_&i._&j;
        %let d=%bd(l,0) %bd(r,0) %bd(t,6) %bd(b,0) foreground=&clr;
        %let head_str=%qsysfunc(tranwrd(&head_str,&s,&d));
        %let s=bdstr3_&i._&j;
        %let d=%bd(l,0) %bd(r,0) %bd(t,0) %bd(b,3.5) foreground=&clr;
        %let head_str=%qsysfunc(tranwrd(&head_str,&s,&d));
        %let s=bdstr4_&i._&j;
        %let d=%bd(l,0) %bd(r,0) %bd(t,0) %bd(b,0) foreground=&clr;
        %let head_str=%qsysfunc(tranwrd(&head_str,&s,&d));
        %let s=bdstr5_&i._&j;
        %let d=%bd(l,0) %bd(r,0) %bd(t,0) %bd(b,0) foreground=&clr;
        %let head_str=%qsysfunc(tranwrd(&head_str,&s,&d));
      %end;
    %end;
    %let head_str=%unquote(&head_str);
    %output_rtf_style2(&DataSet, &Output_RTF_File,&head_str,&tot_str,&groups,&levels,
      &bdwidths,&cellwidths,&blankcols,&PaperSize,&Orientation,&Margin,
      &FirstColumnWidth,&CharFont,&NumFont,&FontSize,&RowShadow);
  %mend;

  %local saved_opts; %let saved_opts=%save_options; %change_options; 

  %delete_temp_datasets(tmp_:);

  %create_total_count(&DataSet, &GroupVars);

  %* create n percent tables for categorical row variables;
  %create_cat_rows(&DataSet, &StudyVars, &StudyVarTypes, &GroupVars, &MaxDec);

  %* create statistic tables for continue row variables;
  %create_con_rows(&DataSet, &StudyVars, &StudyVarTypes, &GroupVars, &Statistics, &MaxDec);

  %* combine all tables into one character table;
  %convert_num_chr(&DataSet, &GroupVars, &GrandSubTotal, &ShowMissing, &Style);

  %* add p-values;
  %add_pvalues(&DataSet, &StudyVars, &StudyVarTypes, &GroupVars);

  data Summary_Report_Output_Table(drop=VarType VarGroup VarOrder); set tmp_11_summary; run;

  %* output rtf file;
  %if &Style=1 %then %output_style1(&DataSet,&GroupVars,&Output_RTF_File,&PaperSize,
    &Orientation,&Margin,&FirstColumnWidth,&CharFont,&NumFont,&FontSize,&RowShadow);
  %if &Style=2 %then %output_style2(&DataSet,&GroupVars,&Output_RTF_File,&PaperSize, 
    &Orientation,&Margin,&FirstColumnWidth,&CharFont,&NumFont,&FontSize,&RowShadow);

  %delete_temp_datasets(tmp_:); /* clean workspace */

  %* restore options;
  options &saved_opts;

%mend;
