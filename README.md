# utl-excel-changing-cell-contents-inside-proc-report
Excel changing cell contents inside proc report

    Excel changing cell contents inside proc report                                                                    
                                                                                                                       
    If you are going to use the cell properties created by proc report you                                             
    are going to have to set the properties of cells a20 and b21 to match 'report excel table'.                        
                                                                                                                       
    So add a blank row to the end of class or use XLconnect to set properties(messy).                                  
                                                                                                                       
    github                                                                                                             
    https://tinyurl.com/y3pu6yj5                                                                                       
    https://github.com/rogerjdeangelis/utl-excel-changing-cell-contents-inside-proc-report                             
                                                                                                                       
    SAS Forum                                                                                                          
    https://tinyurl.com/y6oyfbbl                                                                                       
    https://communities.sas.com/t5/New-SAS-User/Help-required-with-proc-report/m-p/580716                              
                                                                                                                       
    *_                   _                                                                                             
    (_)_ __  _ __  _   _| |_                                                                                           
    | | '_ \| '_ \| | | | __|                                                                                          
    | | | | | |_) | |_| | |_                                                                                           
    |_|_| |_| .__/ \__,_|\__|                                                                                          
            |_|                                                                                                        
    ;                                                                                                                  
                                                                                                                       
    options missing=' ';                                                                                               
    data class;                                                                                                        
    retain name sex;                                                                                                   
    length sex $2;                                                                                                     
    set sashelp.class(obs=3) end=dne;                                                                                  
    output;                                                                                                            
    if dne then do;                                                                                                    
       call missing(of _all_);                                                                                         
       output;                                                                                                         
    end;                                                                                                               
    run;                                                                                                               
                                                                                                                       
    PROC SQL;                                                                                                          
    SELECT                                                                                                             
    DISTINCT COUNT(*) + 1 INTO :CNT1 trimmed                                                                           
    FROM sashelp.class;                                                                                                
    Quit;                                                                                                              
                                                                                                                       
    Up to 40 obs WORK.CLASS total obs=5                                                                                
                                                                                                                       
    Obs     NAME      SEX    AGE    HEIGHT    WEIGHT                                                                   
                                                                                                                       
     1     Alfred      M      14     69.0      112.5                                                                   
     2     Alice       F      13     56.5       84.0                                                                   
     3     Barbara     F      13     65.3       98.0                                                                   
     4                                               ** note the empty row                                             
                                                        we will updat these cells                                      
    And macro variable                                  inside 'proc report'                                           
                                                                                                                       
    %put &=cnt1;                                                                                                       
                                                                                                                       
       CNT1=20                                                                                                         
                                                                                                                       
    *            _               _                                                                                     
      ___  _   _| |_ _ __  _   _| |_                                                                                   
     / _ \| | | | __| '_ \| | | | __|                                                                                  
    | (_) | |_| | |_| |_) | |_| | |_                                                                                   
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                  
                    |_|                                                                                                
    ;                                                                                                                  
                                                                                                                       
                                                                                                                       
       d:/xls/wants.xlsx                                                                                               
                                                                                                                       
        WORKBOOK d:/xls/class.xlsx with sheet class                                                                    
                                                                                                                       
       d:/xls/want.xlsx                                                                                                
          +----------------------------------------------------------------+                                           
          |     A      |    B       |     C      |    D       |    E       |                                           
          +----------------------------------------------------------------+                                           
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                           
          +------------+------------+------------+------------+------------+                                           
       2  | ALFRED     |    M       |    99      |    69      |  112.5     |   ** note age changed to 99;              
          +------------+------------+------------+------------+------------+                                           
       3  | BARBARA    |    F       |    13      |    58      |  101.5     |                                           
          +------------+------------+------------+------------+------------+                                           
       4  | T          |    20      |            |            |            |  RULE: put 'T' in A4 and '20' in B4       
          +------------+------------+------------+------------+------------+                                           
                                                                                                                       
       [CLASS]                                                                                                         
                                                                                                                       
    *          _       _   _                                                                                           
     ___  ___ | |_   _| |_(_) ___  _ __                                                                                
    / __|/ _ \| | | | | __| |/ _ \| '_ \                                                                               
    \__ \ (_) | | |_| | |_| | (_) | | | |                                                                              
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                                              
                                                                                                                       
    ;                                                                                                                  
    title;                                                                                                             
    footnote;                                                                                                          
    ods _all_ close;                                                                                                   
    ods Excel file="d:/xls/classadd.xlsx"                                                                              
    STYLE= sasdocprinter                                                                                               
    options(sheet_name='class'                                                                                         
    Orientation= "landscape"                                                                                           
    Absolute_Column_Width="20"                                                                                         
    Frozen_headers= 'yes'                                                                                              
    embedded_titles='yes'                                                                                              
    embedded_footnotes='yes'                                                                                           
    );                                                                                                                 
                                                                                                                       
    proc report data=class;                                                                                            
    compute name;;                                                                                                     
    if name="" then name="T";                                                                                          
    endcomp;                                                                                                           
    compute sex;                                                                                                       
    if sex=" " then sex="&cnt1";                                                                                       
    endcomp;                                                                                                           
    run;quit;                                                                                                          
                                                                                                                       
    ods excel close;                                                                                                   
                                                                                                                       
                                                                                                                       
