/* Changes to be made in column names 

'AdjusterRegion'n       TO         'Region'n
'AdjusterTeamBranch'n   TO         'RPT_ENTITY'n
'AdjNm'n                TO         'EMP_NM'n
'AdjNumber'n            TO         'EMP_ID'n
'PayPeriodEnding'n      TO         'Payperiodend'n
'NewServices'n
'NewExposures'n -- 'CompletedExposures'n
'TotalHours'n           TO         'Total_Hours'n
'TotalPTOHrsPaid'n
'TotalOTHrsPaid'n
'AdjPosition'n          TO          'JOB_TTL_FULL'N
'Supv#'n                TO         'SPVS_EMP_ID'N
'SupvNam'n              TO         'SPVS_NM'N
payperioddate           TO         'Payperiodend_dt'n
   
		
		
		
*/
PROC IMPORT DATAFILE='/ere/local/shared/actuarial/Share/Claims/WPC/TESTING 2/Adj Prod by Property_Template.xlsx' 
	DBMS=XLSX 
	REPLACE 
	OUT=both1;
	sheet=  "PropertyWL";
	GETNAMES=YES; 
RUN;

data both1; 
set both1; 
payperiodending = put(payperiodend, 8.);
run;
	
	proc sql;
	   create table dups as 
	   select *
	      from both1
	      group by 'AdjNumber'n, 'PayPeriodEnding'n
	      having count(*) gt 1;
	quit;

	
	data both;
		set both1;
		format payperioddate mmddyy10.;

		keep 'AdjusterRegion'n 
		'AdjusterTeamBranch'n
		'AdjNm'n
		'AdjNumber'n
		'PayPeriodEnding'n
		'NewServices'n
		'NewExposures'n -- 'CompletedExposures'n
		'TotalHours'n
		'TotalPTOHrsPaid'n
		'TotalOTHrsPaid'n
		'AdjPosition'n
		'Supv#'n
		'SupvNam'n
		payperioddate

		;
		payperioddate = mdy(input(substr('payperiodending'n,5,2),2.),input(substr('payperiodending'n,7,2),2.),input(substr('payperiodending'n,1,4),4.));

		;
	run;
	/*fix adj titles*/

	data both;
		format AdjPosition $50.;
		set both;
		if AdjPosition = 'IPA' then AdjPosition = 'Inside Property Adjuster';
		else if AdjPosition = 'Property Claims Spct' then AdjPosition = 'Property Claims Specialist';
		else if AdjPosition = 'Property Specialist' then AdjPosition = 'Property Claims Specialist';
		else if AdjPosition = 'Sr Property Specialist' then AdjPosition = 'Sr Property Claims Specialist';
		else if AdjPosition = 'Sr Property Claims Spct' then AdjPosition = 'Sr Property Claims Specialist';
		else if AdjPosition = 'CAT Adjuster' then AdjPosition = 'Catastrophe Property Adjuster';
		else if AdjPosition = 'Property Adjuster I' then AdjPosition = 'Property Adjuster';
		else if AdjPosition = 'Property Adjuster II' then AdjPosition = 'Property Adjuster';
		where AdjPosition ne 'PRPA' and AdjPosition ne 'Property Repair Program Adjuster'
			; 

	run;
		/*fix branch*/

	data both;
		set both;
		length AdjusterRegion $10.;
		if AdjusterTeamBranch = 'Columbus' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'Roanoke' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Charlotte' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Silver Spring' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Hagerstown' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Tennessee' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'Philadelphia' then AdjusterRegion = 'Northeast';
			else if AdjusterTeamBranch = 'Murrysville' then AdjusterRegion = 'Northeast';
			else if AdjusterTeamBranch = 'Harrisburg' then AdjusterRegion = 'Northeast';
			else if AdjusterTeamBranch = 'Wisconsin' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'Erie' then AdjusterRegion = 'Northeast';
			else if AdjusterTeamBranch = 'Allentown' then AdjusterRegion = 'Northeast';
			else if AdjusterTeamBranch = 'Indiana' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'New York' then AdjusterRegion = 'Northeast';
			else if AdjusterTeamBranch = 'West Virginia' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Richmond' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Pittsburgh' then AdjusterRegion = 'Northeast';
			else if AdjusterTeamBranch = 'Raleigh' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Fort Wayne' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'Illinois' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'Canton' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'Kentucky' then AdjusterRegion = 'West';
			else if AdjusterTeamBranch = 'Waynesboro' then AdjusterRegion = 'Southeast';
			else if AdjusterTeamBranch = 'Claims' then AdjusterRegion = 'Claims';
		run;

/* combine */
PROC SQL;
	CREATE TABLE propcurr.hcoutput As
		Select
			'AdjusterRegion'n	
			,'AdjusterTeamBranch'n	
			,'AdjPosition'n	
			,'AdjNumber'n	
			,payperioddate	
			,'NewServices'n	
			,'NewExposures'n	
			,'OpenExposures'n	
			,'CompletedExposures'n	
			,'TotalHours'n
			,'TotalOTHrsPaid'n
			,'TotalPTOHrsPaid'n	
			,'Supv#'n
			,'SupvNam'n
			,'AdjNm'n
		From both 

	;
QUIT;

	data propcurr.hcoutput2;
		set propcurr.hcoutput;
		drop 
		'Supv#'n
		'SupvNam'n
		'AdjNm'n
	;
	run;




/* create adjuster/branch/position decode */
PROC SQL;
	CREATE TABLE adjusters_decode As
		Select distinct a.'AdjNm'n	
		,a.'AdjPosition'n
		,a.'AdjusterTeamBranch'n
		,a.'AdjNumber'n
		,cat('AdjusterTeamBranch'n,'AdjPosition'n) as group
		From both as A 
		where a.payperioddate 
		order by 
		a.'AdjusterTeamBranch'n
		,a.'AdjPosition'n
		,a.'AdjNm'n
		,a.'AdjNumber'n	
	;
QUIT;
	/* add counter and lookup field */
	data adjusters_decode;
		set adjusters_decode;
		by group;
		retain agt;
		if first.group = 1 then agt = 1; else agt = agt+1;
	run;	
	 PROC SQL;
		CREATE TABLE propcurr.adjusters_decode2 As
			Select 
				cat(agt,trim('AdjusterTeamBranch'n),'AdjPosition'n) as bradjseq
				,'AdjNm'n	
				,'AdjPosition'n
				,'AdjusterTeamBranch'n
				,'AdjNumber'n
			From adjusters_decode 
			;
		QUIT;


proc sql;
   create table propcurr.adj_decode as 
   select distinct 
			compress(a.'AdjNm'n) as adjnm
			,a.'AdjNumber'n
      from both1 as a
;
quit;


libname propcurr clear;

