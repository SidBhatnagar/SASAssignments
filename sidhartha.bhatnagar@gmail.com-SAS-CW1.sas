libname libcw '/folders/myfolders';

proc import datafile = '/folders/myfolders/OLYMPICS_DIGI.XLS'
DBMS=  XLS
OUT= Libcw.OLYMPICS;
GETNAMES = YES;
RUN;	

proc import datafile = '/folders/myfolders/OLYMPICS.XLS'
DBMS=  XLS
OUT= libcw.OLYMPICS1;
GETNAMES = YES;
RUN;	


DATA Descriptiona ;
Set  libcw.OLYMPICS1 ;
Label Account_Name = Client Name
Opportunity_Owner = Sales Person(Champion)
Primary_Contact = Primary contact from Client Side
Created_Date = Deal creation in CRM Application
Probability____ = Probability assigned to different stages of the deal
Total_Property_s__Budget_Currenc = Currency
Total_Property_s__Budget = As of now money commited for the different slots
Comments = Comments
Stage = Current Stage of the deal
Sports_Elements_seleceted = Ad slot for the sports element
Opportunity_Name = CRM Application generated name for the deal
Description = Description of the deal
Deal_Comments = Comments from the sales person
Industry = Client serving industry
Last_Modified_Date = Latest of the deal modification in the CRM system
Total_Media_Value_Currency = Currency
Total_Media_Value = Estimated deal value;
Format Total_Media_Value dollar26.; 
Run;

data Format;
set descriptiona;
array ch(*) _character_;
do _n_ = 1 to dim(ch);
ch(_n_) = coalescec(ch(_n_), '-');
end;
run;

data Useful;
Set Format;
keep Account_Name Opportunity_Owner Created_Date Probability____ Deal_Comments Total_Media_Value;
run;

data calculation1;
set useful;
format Tot_Forecast dollar26.;
Tot_Forecast = Total_Media_Value*Probability____/100;
run; 

data calculation2;
set calculation1;
format Tot_Budget dollar26.;
Tot_Budget = sum(Total_Media_Value);
run;

data Olympicsq;
set libcw.olympics;
rename _probability___ = Probability____;
run; 

data Useful1;
Set Olympicsq;
format Total_Media_Value dollar26.;
keep Account_Name Opportunity_Owner Created_Date Probability____ Deal_Comments Total_Media_Value;
run;

data calculation3;
set useful1;
Format Digital_Bugt dollar26.;
Digital_Bugt = Total_Media_Value;
run;

data calculation4;
set calculation3;
Format D_forecast dollar26.;
D_Forecast = Total_Media_Value*Probability____/100;
run;

options  nodate  ;
ods  PDF startpage=no startpage=now
color=full  Dpi=300  

File='/folders/myfolders/REPORT1.PDF';
    



Proc tabulate data = calculation2  out = summary1
style = [color = black borderbottomcolor=black bordercolor=black borderleftcolor=black 
borderrightcolor=black bordertopcolor=black
 COLOR = BLACK  borderwidth=2]; 
Title bold italic color = black 'London olympic pipeline as of 17th November 2012';
where Probability____>0;
CLASS Probability____ /descending ;
keyword N / style=[background=gray color = black 
                       bordertopcolor=black borderbottomcolor=black borderwidth=2 
                       borderleftcolor=black borderrightcolor=black];
Keyword all / style=[background=gray color = black 
                       bordertopcolor=black borderbottomcolor=black borderwidth=2 
                       borderleftcolor=black borderrightcolor=black];
ClassLEV Probability____/style=[Color = Black background = white bordertopcolor=black borderbottomcolor=black borderwidth=2 
                       borderleftcolor=black borderrightcolor=black ];     
VAR Tot_Budget/style= [backgroundcolor = gray foreground = black 
     borderbottomcolor=black bordercolor=black borderleftcolor=black
             borderrightcolor=black bordertopcolor=black  borderwidth=2]; 
VAR Tot_Forecast /style= [backgroundcolor = gray foreground = black 
     borderbottomcolor=black bordercolor=black borderleftcolor=black
             borderrightcolor=black bordertopcolor=black  borderwidth=2];
     
TABLES Probability____='' all=''*[Style=[background=gray color = black  borderbottomcolor=black 
                       bordercolor=black borderleftcolor=black
                       borderrightcolor=black bordertopcolor=black fontweight=bold borderwidth=2]],
       N = Nbr_of_Optys sum=''*Tot_Budget*f=dollar26. sum=''*Tot_Forecast*f=dollar26. 
       / box='Probability' box=[style = [color = black background = gray  borderwidth=2  
       Fontweight = Bold borderbottomcolor=black bordercolor=black borderleftcolor=black 
       borderrightcolor=black bordertopcolor=black]] nocellmerge; 

 
Proc tabulate data = calculation4 style = [color = black borderbottomcolor=black bordercolor=black borderleftcolor=black 
borderrightcolor=black bordertopcolor=black
 borderwidth=2] out = summary2;
Title bold italic color = black 'London olympic pipeline as of 17th November 2012';
where Probability____>0;
CLASS Probability____ /descending ;
keyword N / style=[background=gray color = black 
                       bordertopcolor=black borderbottomcolor=black borderwidth=2 
                       borderleftcolor=black borderrightcolor=black];
Keyword all / style=[background=gray color = black 
                       bordertopcolor=black borderbottomcolor=black borderwidth=2 
                       borderleftcolor=black borderrightcolor=black];
ClassLEV Probability____/style=[Color = Black background = white bordertopcolor=black borderbottomcolor=black borderwidth=2 
                       borderleftcolor=black borderrightcolor=black ];                       
                       
Var Digital_Bugt/
     style= [backgroundcolor = gray foreground = black 
     borderbottomcolor=black bordercolor=black borderleftcolor=black
             borderrightcolor=black bordertopcolor=black  borderwidth=2];
Var D_Forecast/
     style= [backgroundcolor = gray foreground = black 
     borderbottomcolor=black bordercolor=black borderleftcolor=black
             borderrightcolor=black bordertopcolor=black  borderwidth=2];
TABLES Probability____=''all=''*[Style=[background=gray color = black  borderbottomcolor=black 
                       bordercolor=black borderleftcolor=black
                       borderrightcolor=black bordertopcolor=black fontweight=bold borderwidth=2]],
       N = 'Nbr_of_Optys' sum=''*Digital_Bugt*f=dollar26. sum=''*D_Forecast*f=dollar26. 
       / box='Probability' box=[style = [color = black background = gray  borderwidth=2  
       Fontweight = Bold borderbottomcolor=black bordercolor=black borderleftcolor=black 
       borderrightcolor=black bordertopcolor=black]] nocellmerge; 
                       
run;

proc sql;
create table DataMerge as
select L.*, R.*
from Calculation2 as L
Left JOIN Calculation4 as R
on L.Probability____=R.Probability____ and L.Account_Name=R.Account_Name;
quit;

options missing = '-';
proc report data = DataMerge out = Procrep
style(Column) = [color = black borderbottomcolor=black bordercolor=black borderleftcolor=black
              borderrightcolor=black bordertopcolor=black COLOR = BLACK borderwidth=2]
Style(Header) = [ Backgroundcolor=gray color = black borderbottomcolor=black bordercolor=black borderleftcolor=black
              borderrightcolor=black bordertopcolor=black COLOR = BLACK borderwidth=2]
Style(Summary) = [color = black borderbottomcolor=black bordercolor=black borderleftcolor=black
              borderrightcolor=black bordertopcolor=black COLOR = BLACK borderwidth=2];
Title Bold italic Color = Black 'Olympics pipeline (London)- by Probability as of 17th November 2012 ';
Column Probability____ Account_Name Opportunity_Owner Created_Date Tot_Budget Digital_Bugt Deal_Comments;
Define Probability____ /display 'Probability' order descending ;
Define Account_Name / display 'Client' style=[cellwidth=5];
Define Opportunity_Owner / display 'Champ' ;
Define Created_Date/ display 'Modified';
Define Tot_Budget / display 'Total_Budget';
Define Digital_Bugt / display 'Digital_Bugt';
Define Tot_Budget / analysis Sum '' format = dollar26.;
Define Digital_Bugt/analysis Sum '' format = dollar26.;
Define Deal_Comments/ display 'Deal_comments' Style=[Cellwidth=350];
break after  Probability____ / summarize suppress style= [Backgroundcolor=gray] ;
run;

ODS PDF Close   ;  

Proc copy in = work out = libcw;
Run;



