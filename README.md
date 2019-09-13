# utl-single-click-and-eight-excel-tabs-are-converted-to-csv-files
Single click and eight excel tabs are converted to csv files  
    SAS Forum: Single click and eight excel tabs are converted to csv files                                       
                                                                                                                  
      There are four ways you can do the conversion                                                               
                                                                                                                  
         1. click on RMB                                                                                          
         2. Type shoot_csv on the classic original clean command line in any window                               
         3. %shoot_csv in program editor and 'run program'                                                        
         4. Hit Function key PF1                                                                                  
                                                                                                                  
    FYI: It is trivial to pop up a window and ask for tab names                                                   
         or use sql or proc contents to get all the sheet names;                                                  
                                                                                                                  
                                                                                                                  
    * No need for proc xport.                                                                                     
                                                                                                                  
    github                                                                                                        
    https://tinyurl.com/yxmlsvr4                                                                                  
    https://github.com/rogerjdeangelis/utl-single-click-and-eight-excel-tabs-are-converted-to-csv-files           
                                                                                                                  
    Supporting repos                                                                                              
    dee github                                                                                                    
    https://tinyurl.com/y74wdoya                                                                                  
    https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=classic+in%3Aname&type=&language=        
                                                                                                                  
    SAS Forum                                                                                                     
    https://tinyurl.com/yyk7vz9f                                                                                  
    https://communities.sas.com/t5/SAS-Programming/Converting-XLSX-to-CSV-in-SAS/td-p/543577                      
                                                                                                                  
    *_                   _                                                                                        
    (_)_ __  _ __  _   _| |_                                                                                      
    | | '_ \| '_ \| | | | __|                                                                                     
    | | | | | |_) | |_| | |_                                                                                      
    |_|_| |_| .__/ \__,_|\__|                                                                                     
            |_|                                                                                                   
    ;                                                                                                             
                                                                                                                  
     * this creates named ranges however you cab change to ''tab$'n; or 'sheet$1'n;                               
                                                                                                                  
    %utlfkil(d:/xls/tabs.xlsx);                                                                                   
    libname xel "d:/xls/tabs.xlsx";                                                                               
    data xel.ken xel.mat xel.jan;                                                                                 
        set sashelp.class(obs=6);                                                                                 
        select;                                                                                                   
           when (mod(_n_,3)=1 ) output xel.ken;                                                                   
           when (mod(_n_,3)=2 ) output xel.mat;                                                                   
           when (mod(_n_,3)=0 ) output xel.jan;                                                                   
        end;                                                                                                      
    run;quit;                                                                                                     
    libname xel clear;                                                                                            
                                                                                                                  
    NOTE: There were 19 observations read from the data set SASHELP.CLASS.                                        
    NOTE: The data set XEL.ken has 19 observations and 5 variables.                                               
    NOTE: The data set XEL.mat has 19 observations and 5 variables.                                               
    NOTE: The data set XEL.jan has 19 observations and 5 variables.                                               
                                                                                                                  
                                                                                                                  
     d:/xls/tabs.xlsx                                                                                             
                                                                                                                  
       +--------------------------------------+                                                                   
       |  A    |  B    |  C    |  D    |  E   |                                                                   
       +--------------------------------------+                                                                   
     1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|                                                                   
       +-------+-------+-------+-------+------|                                                                   
     2 |Alfred |14     |M      |69     |112.5 |                                                                   
       +-------+-------+-------+-------+------+                                                                   
     3 |Carol  |13     |F      |56.5   |84    |                                                                   
       ---------------------------------------+                                                                   
       [KEN]                                                                                                      
                                                                                                                  
                                                                                                                  
       +--------------------------------------+                                                                   
       |  A    |  B    |  C    |  D    |  E   |                                                                   
       +--------------------------------------+                                                                   
     1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|                                                                   
       +-------+-------+-------+-------+------|                                                                   
     2 |Alice  |14     |F      |62.8   |102.5 |                                                                   
       +-------+-------+-------+-------+------+                                                                   
     3 |Henry  |14     |M      |63.5   |88    |                                                                   
       ----------------------------------------                                                                   
       [MAT]                                                                                                      
                                                                                                                  
                                                                                                                  
       +--------------------------------------+                                                                   
       |  A    |  B    |  C    |  D    |  E   |                                                                   
       +--------------------------------------+                                                                   
     1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|                                                                   
       +-------+-------+-------+-------+------+                                                                   
     2 |Barbara|14     |M      |63.5   |88    |                                                                   
       ----------------------------------------                                                                   
     3 |James  |13     |M      |73.5   |78    |                                                                   
       ----------------------------------------                                                                   
       [class]                                                                                                    
                                                                                                                  
    *            _               _                                                                                
      ___  _   _| |_ _ __  _   _| |_                                                                              
     / _ \| | | | __| '_ \| | | | __|                                                                             
    | (_) | |_| | |_| |_) | |_| | |_                                                                              
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                             
                    |_|                                                                                           
    ;                                                                                                             
                                                                                                                  
     d:/csv/ken.csv                                                                                               
                                                                                                                  
    "NAME","SEX","AGE","HEIGHT","WEIGHT"                                                                          
    "Alfred","M",14,69.0,112.5                                                                                    
    "Carol","F",14,62.8,102.5                                                                                     
                                                                                                                  
                                                                                                                  
     d:/csv/mat.csv                                                                                               
                                                                                                                  
     NAME","SEX","AGE","HEIGHT","WEIGHT"                                                                          
    "Alice","F",13,56.5,84.0                                                                                      
    "Henry","M",14,63.5,102.5                                                                                     
                                                                                                                  
                                                                                                                  
     d:/csv/jan.csv                                                                                               
                                                                                                                  
     "NAME","SEX","AGE","HEIGHT","WEIGHT"                                                                         
    "Barbara","F",13,65.3,98                                                                                      
    "James","M",12,57.3,83                                                                                        
                                                                                                                  
    *                                                                                                             
     _ __  _ __ ___   ___ ___  ___ ___                                                                            
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                           
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                           
    | .__/|_|  \___/ \___\___||___/___/                                                                           
    |_|                                                                                                           
    ;                                                                                                             
                                                                                                                  
    Put this on right mouse button                                                                                
                                                                                                                  
    RMB note;notesubmit '%shoot_csv:'; /* can map to anny mouse action */                                         
    or                                                                                                            
    PF1 note;notesubmit '%shoot_csv:';                                                                            
                                                                                                                  
    * put this in your autocall librar;                                                                           
                                                                                                                  
    %macro shoot_csv() / cmd;                                                                                     
                                                                                                                  
        %array(dsns,values=ken mat jan);                                                                          
                                                                                                                  
        %let rc=%sysfunc(dosubl('                                                                                 
          %do_over(dsns,phrase=%str(                                                                              
                                                                                                                  
            libname xel "d:/xls/tabs.xlsx";                                                                       
                                                                                                                  
            ods csv file="d:/csv/?.csv"                                                                           
            options(doc="Help" delimiter=",");                                                                    
               proc print data=xel.? noobs;                                                                       
            run;quit;                                                                                             
            ods csv close;                                                                                        
                                                                                                                  
            libname xel clear;));                                                                                 
          '));                                                                                                    
    %mend shoot_csv;                                                                                              
                                                                                                                  
    %shoot_csv;                                                                                                   
                                                                                                                  
                                                                                                                  
                                                                                                                  
