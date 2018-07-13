dm 'log;clear;output;clear;';
options mprint;

***Enter the START and END date for the window you want to produce title picker results for - note that these dates are INCLUSIVE***;
***Format should be mm/dd/yyyy e.g. 05/07/2018***;
%let inp_start = 05/21/2017;
%let inp_end = 05/24/2019;

/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ */ 
/* ---------------------------DON'T TOUCH CODE BELOW THIS LINE-------------------------------------- */
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ */ 

***Convert to SAS date format;
%let tp_start = %sysfunc(inputn(&inp_start, mmddyy10.));
%let tp_end = %sysfunc(inputn(&inp_end, mmddyy10.));
/*****************************************************************
Step 1. Define file names and macro variable for the
project-specific token.
*******************************************************************/
*** Text file for API parameters that the define the request sent to
REDCap API. Will be created in a DATA step. Extension can be .csv,
.txt, .dat ***;
filename my_in "P:\Mapping Project\Title Picker items\SAS\api_info.txt";
*** .CSV output file to contain the exported data ***;
filename my_out "P:\Mapping Project\Title Picker items\SAS\title_picker_data.csv"; 

*** Output file to contain PROC HTTP status information returned from
REDCap API (this is optional) ***;
filename status "P:\Mapping Project\Title Picker items\SAS\redcap_status.txt";

*** Project- and user-specific token obtained from REDCap ***;
%let duketoken = API_TOKEN&; 
 /**********************************************************
Step 2. Request all observations (CONTENT=RECORDS) with one
row per record (TYPE=FLAT). Note: Obtain your site-specific
 url from your local REDCap support team.
******************************************************/


*Import Data from RTI REDCAP;

*** Create the text file to hold the API parameters. ***; 
data _null_ ;
file my_in ;
put "%NRStr(content=record&format=csv&type=flat&exportSurveyFields=true&exportDataAccessGroups=Duke&token=)&duketoken";
run;
*** PROC HTTP call. Everything except HEADEROUT= is required. ***;
proc http
 in = my_in
 out = my_out
 headerout = status
 url ='https://redcap.duke.edu/redcap/api/'
 method="post";
run; 

*%include will run the SAS program which reads in the .csv file created above and assigns labels;
%include 'P:\Mapping Project\Title Picker items\SAS\sas_inputs.txt';

*Now we want to pull records that were created between tp_start and tp_end and completed;
proc sql;
create table redcap_filter as select * from redcap
	where title_picker_timestamp between &tp_start and &tp_end and title_picker_complete = 2;
quit;

*Create variables that will display in the report by appending proficiencies if not = 0;
data tp_report; /*Keep statement to only keep the variables that are going in the report*/
	set redcap_filter end = last;

	length operations $5000
		ethics $5000
		data $5000
		scicon $5000
		leadership $5000
		sitestudyman $5000
		progport $5000
		clinical_var $5000;

*Create compentency descriptions by concatenating appropriate competencies;
	operations = catx(' ',
		ifc( oy_1 = 1, put(o1_contracts, o1_contracts_. ),''),
		ifc( oy_2 = 1, put(o2_fda, o2_fda_. ),''),
		ifc( oy_3 = 1, put(o3_regpp, o3_regpp_. ),''),
		ifc( oy_4 = 1, put(o4_international, o4_international_. ),''),
		ifc( oy_5 = 1, put(o5_ip, o5_ip_. ),''),
		ifc( oy_6 = 1, put(o6_monaudit, o6_monaudit_. ),''),
		ifc( oy_7 = 1, put(o7_partdoc, o7_partdoc_. ),''),
		ifc( oy_8 = 1, put(o8_retention, o8_retention_. ),''),
		ifc( oy_9 = 1, put(o9_recruitment, o9_recruitment_. ),''),
		ifc( oy_10 = 1, put(o10_screen, o10_screen_. ),''),
		ifc( oy_11 = 1, put(o11_sop, o11_sop_. ),''),
		ifc( oy_12 = 1, put(o12_specimen, o12_specimen_. ),''),
		ifc( oy_13 = 1, put(o13_studydoc, o13_studydoc_. ),''),
		ifc( oy_14 = 1, put(o14_visits, o14_visits_. ),''),
		ifc( oy_15 = 1, put(o15_meetings, o15_meetings_. ),''));

	ethics = catx(' ',
		ifc( ey_1 = 1, put(e1_ae, e1_ae_. ),''),
		ifc( ey_2 = 1, put(e2_consent, e2_consent_. ),''),
		ifc( ey_3 = 1, put(e3_devconsentplan, e3_devconsentplan_. ),''),
		ifc( ey_4 = 1, put(e4_irb, e4_irb_. ),''),
		ifc( ey_5 = 1, put(e5_sponregreport, e5_sponregreport_. ),''));

	data = catx(' ',
		ifc( dy_1 = 1, put(d1_datacollect, d1_datacollect_.),''),
		ifc( dy_2 = 1, put(d2_qa, d2_qa_.),''),
		ifc( dy_3 = 1, put(d3_security, d3_security_.),''),
		ifc( dy_4 = 1, put(d4_dataflow, d4_dataflow_.),''),
		ifc( dy_5 = 1, put(d5_techinno, d5_techinno_.),''));

	scicon = catx(' ',
		ifc( scy_1 = 1, put(sc1_fundprop, sc1_fundprop_.),''),
		ifc( scy_2 = 1, put(sc2_litreview, sc2_litreview_.),''),
		ifc( scy_3 = 1, put(sc3_protocol, sc3_protocol_.),''),
		ifc( scy_4 = 1, put(sc4_resdesign, sc4_resdesign_.),''),
		ifc( scy_5 = 1, put(sc5_scholarworks, sc5_scholarworks_.),''));

	leadership = catx(' ',
		ifc( ly_1 = 1, put(l1_prodev, l1_prodev_.),''),
		ifc( ly_2 = 1, put(l2_externalaware, l2_externalaware_.),''),
		ifc( ly_3 = 1, put(l3_orgagility, l3_orgagility_.),''),
		ifc( ly_4 = 1, put(l4_adapt, l4_adapt_.),''),
		ifc( ly_5 = 1, put(l5_sme, l5_sme_.),''),
		ifc( ly_6 = 1, put(l6_teamwork, l6_teamwork_.),''));

	sitestudyman = catx(' ',
		ifc( ssmy_1 = 1, put(ssm1_visits, ssm1_visits_.),''),
		ifc( ssmy_2 = 1, put(ssm2_electronic, ssm2_electronic_.),''),
		ifc( ssmy_3 = 1, put(ssm3_feasibility, ssm3_feasibility_.),''),
		ifc( ssmy_4 = 1, put(ssm4_resources, ssm4_resources_.),''),
		ifc( ssmy_5 = 1, put(ssm5_risk, ssm5_risk_.),''),
		ifc( ssmy_6 = 1, put(ssm6_oppplans, ssm6_oppplans_.),''),
		ifc( ssmy_7 = 1, put(ssm7_closeout, ssm7_closeout_.),''));

	progport = catx(' ',
		ifc( progport1_infra > 0, put(progport1_infra, progport1_infra_.),''),
		ifc( progport2_progport > 0, put(progport2_progport, progport2_progport_.),''),
		ifc( progport3_lias > 0, put(progport3_lias, progport3_lias_.),''));
	
	clinical_var = catx('^n• ',' ',
		ifc(abul_admin = 1, vlabel(abul_admin),''),
		ifc(adult_med = 1, vlabel(adult_med),''),
		ifc(ped_med = 1, vlabel(ped_med),''),
		ifc(or_med = 1, vlabel(or_med),''),
		ifc(blood_admin = 1, vlabel(blood_admin),''),
		ifc(peripheral_iv = 1, vlabel(peripheral_iv),''),
		ifc(alaris_pump = 1, vlabel(alaris_pump),''),
		ifc(blood_draw = 1, vlabel(blood_draw),''),
		ifc(external_genital = 1, vlabel(external_genital),''),
		ifc(vaginal = 1, vlabel(vaginal),''),
		ifc(fetal_monitor = 1, vlabel(fetal_monitor),''),
		ifc(adult_chemo = 1, vlabel(adult_chemo),''),
		ifc(adult_bone_marrow = 1, vlabel(adult_bone_marrow),''),
		ifc(adult_dys = 1, vlabel(adult_dys),''),
		ifc(basic_dys = 1, vlabel(basic_dys),''),
		ifc(vesoactive = 1, vlabel(vesoactive),''),
		ifc(other_clinical_res = 1, catx(': ',vlabel(other_clinical_res),clinical_desc),''));


*Keep only the variables we will need for the report;
	keep record_id 
		name_requester 
		man_title
		man_fname
		man_lname
		type_position
		operations
		ethics
		data
		scicon
		leadership
		sitestudyman
		progport
		perc_port
		progport_desc
		progport_number
		progport_staff
		progport_liason
		clinical_var
		research_type
		special_skills; 

*Set the macro variables for the first record and last record;
	if _n_ = 1 then call symput('first_record', record_id);
	if last then call symput('last_record', record_id);
run;

/* This template will be used for the rtf files generated below*/
proc template;
	define style tp_style;
	parent = styles.rtf;
	style body from document /
		leftmargin=1.0in
		rightmargin=1.0in
		topmargin=1.0in
		bottommargin=1.0in;
	end;
run;
/* Options that make the rtf output look better */
options nodate nonumber;
title;
ods escapechar = '^';


***Need to generate the word document - one per record***;
***Use a loop that generates macro variables from an individual record, then, creates the report output using ODS rtf text = " macrovariables", then start on the next record; 

*This macro loop will run from each record from first_record to last_record (defined above);
%macro rtf_gen;
%do I = &first_record %to &last_record;

	*PROC SQL code will be used to filter to the specific record;
	proc sql noprint;
			select record_id,
				name_requester,
				man_title,
				man_fname,
				man_lname,
				type_position,
				operations,
				ethics,
				data,
				scicon,
				leadership,
				sitestudyman,
				progport,
				perc_port,
				progport_desc,
				progport_number,
				progport_staff,
				progport_liason,
				clinical_var,
				research_type,
				special_skills
			into :record_id TRIMMED,
				:name_requester TRIMMED,
				:man_title TRIMMED,
				:man_fname TRIMMED,
				:man_lname TRIMMED,
				:type_position TRIMMED,
				:operations TRIMMED,
				:ethics TRIMMED,
				:data TRIMMED,
				:scicon TRIMMED,
				:leadership TRIMMED,
				:sitestudyman TRIMMED,
				:progport TRIMMED,
				:perc_port TRIMMED,
				:progport_desc TRIMMED,
				:progport_number TRIMMED,
				:progport_staff TRIMMED,
				:progport_liason TRIMMED,
				:clinical_var TRIMMED,
				:research_type TRIMMED,
				:special_skills TRIMMED
			from tp_report
			where record_id = "&I";
	quit;

	/*Use ods tagsets.rtf to generate the word documents - tagsets makes it NOT a table*/
	ods tagsets.rtf file = "P:\Mapping Project\Title Picker items\SAS\output\Job_Description_&I..rtf" style = tp_style options (tables_off = 'usertext');
	ods tagsets.rtf text = "^{style [fontweight=bold]Record ID: &record_id}^n
		^{style [fontweight=bold]Title of requester: &name_requester}^n
		^{style [fontweight=bold]Title of person who will manage position: &man_title} ^n
		^{style [fontweight=bold]Name of person who will manage the position: &man_fname &man_lname} ^n
		^{style [fontweight=bold]Type of position: &type_position} ^n^n
		^{style [textdecoration=underline]Operations:}^n &operations ^n ^n
		^{style [textdecoration=underline]Ethics:}^n &ethics ^n ^n
		^{style [textdecoration=underline]Data:}^n &data ^n ^n
		^{style [textdecoration=underline]Science:}^n &scicon ^n ^n
		^{style [textdecoration=underline]Study and Site Management:}^n &sitestudyman ^n ^n 
		^{style [textdecoration=underline]Leadership:}^n &leadership ^n ^n ^n
		^{style [textdecoration=underline]Description of Portfolio Responsibilities:}^n ^n &progport ^n ^n
		Portfolio Management (Effort &perc_port%): ^n
		&progport_desc ^n
		&progport_number ^n
		&progport_staff ^n
		&progport_liason ^n ^n
		^{style [textdecoration=underline]Description of Clinical Responsibilities:}^n ^n
		Clinical responsibilities: ^n
		• &clinical_var ^n ^n ^n
		^{style [textdecoration=underline]Type of Research:}^n  &research_type ^n ^n
		^{style [textdecoration=underline]Special skills:}^n &special_skills";

	ods tagsets.rtf close;
	
%END;
%MEND rtf_gen;

%rtf_gen;

