 
#from tkinter import FIRST Spire.XLS
import psycopg2
 
import datetime
from datetime import date, timedelta,datetime
#import xlsxwriter
#from openpyxl import load_workbook
#from openpyxl.workbook import Workbook
#from openpyxl.cell import Cell  
import pandas as pd
import numpy as np
import os
import sys
import logging
import locale
#from openpyxl.styles import Alignment
#import dateutil.relativedelta
#from openpyxl.utils.dataframe import dataframe_to_rows
#from openpyxl.styles import Color, Fill, Font, Border,Side, PatternFill
#from openpyxl.styles import colors
#import openpyxl 
#from openpyxl import Workbook
 
#from win32com import client
 
#import pathlib
#from functools import reduce
#import win32api
#from shutil import copyfile
from config import config
from config2 import config2
import dateutil.relativedelta 
 
from dateutil.relativedelta import relativedelta
#from win32com.client import Dispatch
#import win32com.client as win32
import codecs
import shutil
#from pyexcelerate import Workbook
import psutil
import win32com.client as win32
from pyexcelerate import Workbook, Color, Style, Fill
import pyexcelerate


dirpath = os.getcwd()
 
today =  date.today()
  
 
file_template = dirpath +    '\\setup_audit_template.xlsx' 
#filename1 = "O:\\ANALYTICS\\Setup_Audit\\Set Up Audit "  +  today.strftime("%m/%d/%Y").replace('/','.')   + "(Company Node).xlsx"
filename1_1 = "c:\\APP\\Set Up Audit "  +  today.strftime("%m/%d/%Y").replace('/','.')   + "(Company Node).xlsx" 
filename1_1_tmp = "c:\\APP\\Set Up Audit "  +  today.strftime("%m/%d/%Y").replace('/','.')   + "(Company Node).xlsb"
filename1_2 = "O:\\ANALYTICS\\Setup_Audit\\Setup Audit "  +  today.strftime("%m/%d/%Y").replace('/','.')   + "(Company Node).xlsb"


filename2_1 = "c:\\APP\\Set Up Audit "  +  today.strftime("%m/%d/%Y").replace('/','.')   + "(Criminal Policy).xlsx" 
filename2_2 = "O:\\ANALYTICS\\Setup_Audit\\Setup Audit "  +  today.strftime("%m/%d/%Y").replace('/','.')   + "(Criminal Policy).xlsb"
  
''' 
def connect_acct_manager():
    params = config2()
    conn = psycopg2.connect(**params)       
    cur = conn.cursor() 
    SQL = ("""  
     select trim(company_code) company_code,  max(rs_account_manager)  account_manager from   ycrm_company where coalesce(rs_account_manager,'') <> '' and coalesce(company_code,'')<>''
group by company_code     
   
""" ) 
    cur.execute(SQL) 
    col = [i[0] for i in cur.description]
    conn.set_client_encoding('ISO-8859-1') 
    rows = cur.fetchall()
    acct_am = pd.DataFrame(rows, columns = col)
    cur.close()
    return acct_am
'''  
def generate_terminate_code ( ):
    conn = None
    params = config2()
    
    conn = psycopg2.connect(**params )
 
    cur = conn.cursor()

    SQL = ("""   select trim(company_code)  from "Test_Code_Master"  ; """   )
        
    cur.execute(SQL) 
    #col = cur.description
    col = [i[0] for i in cur.description]
    a = np.array(col)
    rows =   [item[0] for item in cur.fetchall()]
            
    #df_codes = pd.DataFrame(rows, columns = a)
    cur.close()
    conn.close()
    return rows

list_term_code =  generate_terminate_code () 

##### weekly

def generate_company_sql ():
    SQL =    ( """  set client_encoding to 'SQL_ASCII' ; 
select trim(company_tbl.company_code) as "Company_Code", company_tbl.contact_name, ac.contact_email, ac.account_manager , --'    ' as "Account Manager",
case ac.pricing_model	WHEN '0' THEN 'NOT SET' WHEN '1' THEN 'Transactional - Standard'	WHEN '2' THEN 'Transactional - Bundled' WHEN '3' THEN 'Unit-Monthly' WHEN '4' THEN 'Unit-Quarterly' WHEN '5' THEN 'Unit_Annual' WHEN '6' THEN 'Manual' WHEN '7' THEN 'Mixed' ELSE NULL END AS "Company_Pricing_Model", 
case when isfee_mgmt_company = false then 'No' else 'Yes' end as "Fee_Manager_Company" ,
case when export_to_ycrm = false then 'No' else 'Yes' end as "Export_to_yCRM",
case when enable_exec_access_control =false then 'No' else 'Yes' end as "Enable_Exec_Access_Control", 
case when Allow_FCRA_Letter_Generation =false then 'No' else 'Yes' end as "Allow_FCRA_Letter_Generation_(RSEXEC)" , 
case when map_sitecode =false then 'No' else 'Yes' end as "Map_Sitecode", 
 case when guarantorasreject =false then 'No' else 'Yes' end  as "PMS_Workflow_Applicant_w_Guarantor_REC_as_Reject",
case when sync_with_voyager =false then 'No' else 'Yes' end as "Sync_Property_Access_from_Voyager",
case when   autogenerate_fcraletter  =false then  'No' else 'Yes' end as "Show_FCRA_Letter_PDF",
case when autogenerate_fcra_pdf =false then 'No' else 'Yes' end  "Auto_Generate_FCRA_PDF",
case when show_matrix_report =false then 'No' else 'Yes' end   "Show_Benchmark_Report",
case when enable_prequal_screening =false then 'Disabled For Company' else 'Eabled For Company' end as    "Enable_Pre-Qual",
case when suppress_voyager_landing_page =false then 'Disabled For Company' else 'Enabled For Company' end as  "Suppress_Full_Applicant_File",
             UPPER(CAST (ac.tsr_email_delivery as text)) as tsr_email_delivery,            
              translate(lower( trim(ac.tsr_notification_email)  ), ' ''àáâãäéèëêíìïîóòõöôúùüûçÇ', '--aaaaaeeeeiiiiooooouuuucc') tsr_notification_email,

              ac.suppress_validation_workflow,  ac.suppress_cancel_voi , ac.suppress_rralink, 
client_pre_decision_email as "Criminal Records Assessment Email",  
autoemail_fcraletter    "Auto_Email_FCRA_Letters",
case when hold_until_active_app_processs = TRUE then 'Yes' else 'No' end AS "Hold_until_all_Active_Applicants_in_Group_are_Processed",
case when coalesce(send_fcraletter_after_hours,0) <=1 then send_fcraletter_after_hours::text || ' Hour' else send_fcraletter_after_hours::text || ' Hours' end as "Email_FCRA_Letters_every",
case when use_fcra_template_comp =false then 'No' else 'Use System Template' end as "Enable_Template", 
case when include_logo_comp='fcraletter' then 'FCRA Letter' when include_logo_comp='off' then 'Off' else include_logo_comp end as "Include_Logo",
'Enabled For Company' as "Show_Specific_Reasons_for_Result",
left(company_customized_text, 30000) as "Custom_Conditional_Text",
left(fcra_custom_reject_text, 30000) as "Custom_Reject_Text",  

case when fcra_batch_processing =false then 'No' else 'Yes' end as "Enable_FCRA_Batch_Processing",--
case when showcredit_summaryrsexec =false then 'No' else 'Yes' end as "Show_Credit_Summary_in_RSExec",
case when showcredit_summaryyrs =false then 'No' else 'Yes' end as "Show_Credit_Summary_in_YRS",
case when credit_no_ssn =false then 'No' else 'Yes' end as "Process_Credit_for_NO_SSN" ,
		voi_invite_expiry as "VOI Email Invite Expiry"
  from company_tbl LEFT JOIN auxiliary_company AS ac  ON company_tbl.company_code = ac.company_code 
 left join (select comp_code,  autoemail_fcraletter,  autogenerate_fcraletter,  hold_until_active_app_processs ,   send_fcraletter_after_hours send_fcraletter_after_hours
 from autogenerate_fcraletter_config where id in (select max(id) from autogenerate_fcraletter_config where COALESCE(comp_code,'')<>''  and COALESCE(node_code,'')=''  group by comp_code)
  ) fcra on  comp_code = company_tbl.company_code
inner join (select company_code as com_code from node_tbl where canceled = false group by company_code) node_tbl on node_tbl.com_code = company_tbl.company_code
where  active = 'True' 
              and upper(trim(company_tbl.company_code)) not in ('APP')
order by 1     
             
  

"""  )
    return SQL

def generate_node_sql ():
    SQL =    ( """  set client_encoding to 'SQL_ASCII' ; 
   SELECT 
	/* General */
       nt.company_code,
nt.node_code,
              
              
	nt.node_name AS property_name,

	/* NODE ADDRESS AND BILLING INFO*/
  
	nt.node_city, 
	nt.node_state, 
    yardi_prop_code,
              
nt.policy AS screening_policy,  market_rate_units as "Unit Count" , --''''||
  CASE WHEN nt.service_level ='1' THEN 'T1' 
       WHEN nt.service_level ='2' THEN 'T2' 
       WHEN nt.service_level ='3' THEN 'T3' 
       WHEN nt.service_level = '5' then 'TSE' 
ELSE nt.service_level::text end AS totalscreen_level,
 
CASE WHEN an.outgoing_notification=0 THEN 'None'
			 WHEN an.outgoing_notification=1 THEN 'Fax'
       WHEN an.outgoing_notification=2 THEN 'Email'
						ELSE NULL       
		END AS Outgoing_response,
	CASE WHEN nt.allow_corp_applications='FALSE' THEN 'No'
			 WHEN nt.allow_corp_applications='TRUE' THEN 'Yes'
						ELSE NULL
						END AS allow_corporate_apps, 
	CASE WHEN an.biz_product=0 THEN 'None'
			 WHEN an.biz_product=1 THEN 'IntelliScore'
			 WHEN an.biz_product=2 THEN 'D&B'
						ELSE 'None'
						END as biz_product,

UPPER(CAST (fcra_batch_enabled   as text)) fcra_batch_enabled  , 
 
UPPER(CAST (allow_custom_fcra_reasons as text)) allow_custom_fcra_reasons,
UPPER(CAST (an.allow_fcra_letter_generation as text)) allow_fcra_letter_generation,

UPPER(CAST (autoemail_fcraletter as text)) as Auto_Email_Adverse_Action_Letters, 
UPPER(CAST (hold_until_active_app_processs as text)) as "Hold Until All Active Applicants are Processed", 
send_fcraletter_after_hours Email_FCRA_Letters_every, 
left(an.fcra_customized_text,100) as "Custom Conditional Text", 
              left(an.fcra_custom_reject_text,100) "Custom Reject Text", 
 to_char(created_on , 'MM/DD/YYYY'::text) start_date,
 
   UPPER(CAST (an.tsr_email_delivery as text)) as tsr_email_delivery, 
cp.cred_prod_name AS credit_product,
	an.vax_policy AS "Credit Policy (CDS)",
	an.delq_cutoff AS Delinq_Cutoff,

	UPPER(CAST (an.score_medical AS TEXT)) AS score_medical, 
	UPPER(CAST (an.score_student_loans AS TEXT)) AS score_student_loans, 
  UPPER(CAST (an.private_owner AS TEXT)) AS Private_Owner_Passed_Res,
	UPPER(CAST (an.reject_apt_collect AS TEXT)) AS enable_apartment_filter, 
	CASE WHEN reject_apt_collect ='TRUE' AND an.rac_num_months ISNULL THEN 	(	SELECT CAST (resource_value AS INT)
																																						FROM resources 
																																						WHERE resource_group = 'AptCollectionFilter' AND resource_key='DefaultMonths4Reject'
																																					) 	/*Must account for system defaults */
					ELSE an.rac_num_months 
					END AS apt_filter_reject_months,
	CASE WHEN reject_apt_collect ='TRUE' AND an.rac_max_amount ISNULL THEN (	SELECT CAST (resource_value AS INT)
																																						FROM resources 
																																						WHERE resource_group = 'AptCollectionFilter' AND resource_key='DefaultAmount4Reject'
																																					) 	/*Must account for system defaults */
					ELSE an.rac_max_amount
					END AS apt_filter_reject_$_value, 
	CASE WHEN reject_apt_collect ='TRUE' AND an.rac_num_debts ISNULL THEN  (	SELECT CAST (resource_value AS INT)
																																						FROM resources 
																																						WHERE resource_group = 'AptCollectionFilter' AND resource_key='DefaultAlwaysRejectQuantity'
																																					)		/*Must account for system defaults */
					ELSE an.rac_num_debts
					END AS apt_filter_reject_count,
UPPER(CAST (an.util_collect_reject as text)) util_enable_rejection_criteria,
UPPER(CAST (an.util_collect_exclude as  text)) util_exclude_from_credit_scoring,
an.util_collect_num_months,
an.util_collect_max_amount,
an.util_collect_max_num,
CASE WHEN an.show_prem_crim_res=0 THEN 'Never'
			 WHEN an.show_prem_crim_res=1 THEN 'Always'
       WHEN an.show_prem_crim_res=2 THEN 'Does Not Meet Property Requirements'
       WHEN an.show_prem_crim_res=3 THEN 'Inconclusive'
       WHEN an.show_prem_crim_res=4 THEN 'Does Not Meet Property Requirements OR Inconclusive'
       WHEN an.show_prem_crim_res=5 THEN 'Meets Property Requirements'
					ELSE NULL
					END AS "Show Criminal/Offense Results",
	CASE WHEN an.show_prem_evic_res=0 THEN 'Never'
	     WHEN an.show_prem_evic_res=1 THEN 'Always'
       WHEN an.show_prem_evic_res=2 THEN 'Does Not Meet Property Requirements'
       WHEN an.show_prem_evic_res=3 THEN 'Inconclusive'
       WHEN an.show_prem_evic_res=4 THEN 'Does Not Meet Property Requirements OR Inconclusive'
       WHEN an.show_prem_evic_res=5 THEN 'Meets Property Requirements'
					ELSE NULL
					END AS show_civil_court_results, 
	CASE WHEN an.recs_in_out_resp=0 THEN 'Never'
	     WHEN an.recs_in_out_resp=1 THEN 'Always'
       WHEN an.recs_in_out_resp=2 THEN 'Does Not Meet Property Requirements'
       WHEN an.recs_in_out_resp=3 THEN 'Inconclusive'
       WHEN an.recs_in_out_resp=4 THEN 'Does Not Meet Property Requirements OR Inconclusive'
       WHEN an.recs_in_out_resp=5 THEN 'Meets Property Requirements'
					ELSE NULL
					END AS records_in_outgoing_responses, 
	CASE WHEN an.create_linked_pcc IN (0,4) THEN 'Always'
	     WHEN an.create_linked_pcc IN (1,5) THEN 'Best Practice'
       WHEN an.create_linked_pcc IN (2,6) THEN 'Never'
					ELSE NULL
					END AS create_supplemental_criminal_request,
	 no_suppl_crim_states as Skip_Supplemental_Searches_for_States,
	UPPER(CAST(an.launch_supplemental AS TEXT)) AS enable_supplemental_criminal,
	UPPER(CAST((create_linked_pcc&4 > 0)AS TEXT)) as filter_only_for_criminal, 
  UPPER(CAST (an.exclude_sex_offender AS TEXT)) AS exclude_sex_offender,
 UPPER(CAST (an.rejection_reason_custom_criminal AS TEXT)) AS rejection_reason_custom_criminal,
client_crim_pre_decision as crim_record_assessment, 
UPPER(CAST (an.crim_email AS TEXT)) AS crim_email,
 UPPER (cast(enable_rent_to_income as TEXT))      as "Ccalculate RTI Ratio w/o Credit",  
	UPPER(CAST((an.comp_recommendation&1>0) AS TEXT)) AS comprehensive_recommendation,
	replace( replace( 
replace(  
replace(replace( replace( replace( 
replace( replace( replace(replace(
replace(replace(replace(replace(
replace( replace(replace(an.default_recommendation,
      '+',   ' '), 
     '%2B' , ' '),
     '%252B',' '), 
     '%09',  ' '),
     '%27',  '''' ), 
     '%%2F', '/' ),
     '%%2F' , '/' ),
  '%28' , '(' ),
  '%29' , ')' ),
  '%24' , '$' ),
  '%26' , '&' ),
  '%2C' , ',' ),
  '%21' , '!' ),
  '%3D' , '=' ),
  '%3A' , ':' ),
  '%40' , '@' ),
  '%%3F' , '?' ),
  '%%7E' , '~' ) as  default_recommendation,
	UPPER(CAST(an.gen_comp AS TEXT)) AS "Generate Comp Score W/O Credit", 
 
	an.score_without_credit, 
/* Automatic Service */
	--tsc.ts_credit_report, 6/20/18
 case when coalesce(cred_tier,0)=0 then null else ns.cred_tier end cred_tier, 
(select max(vendor_name) from services  where ns.cred = service_id) credit_report,

case when coalesce(crim_tier,0)=0 then null else ns.crim_tier end crim_tier,
   (select max(vendor_name) from services  where ns.crim = service_id) criminal,

 case when coalesce(civilcrt_tier,0)=0 then null else ns.civilcrt_tier end civilcrt_tier,
   (select max(vendor_name) from services  where ns.civilcrt = service_id) civil_court,


 case when coalesce(rhist_tier,0)=0 then null else ns.rhist_tier end rhist_tier,
  (select max(vendor_name) from services  where ns.rhist = service_id) rental_history,

case when coalesce(ofac_tier,0)=0 then null else ns.ofac_tier end ofac_tier,
 (select max(vendor_name) from services  where ns.ofac = service_id) as "OFAC",  
 
       case when coalesce(cust_crim_tier,0)=0 then null else cust_crim_tier end cust_crim_tier,     
(select max(vendor_name) from services  where ns.cust_crim = service_id) custom_criminal, 

 

 case when coalesce(offense_tier,0)=0 then null else offense_tier end offense_tier,    
  (select max(vendor_name) from services  where ns.offense = service_id) offense_alert, 


-- voi
 (select max(vendor_name) from services  where ns.voi = service_id) "VOI", 
		      case when coalesce(voi_tier,0)=0 then null else voi_tier end "VOI Tier",     
 (select max(vendor_name) from services  where ns.voi_ondmd = service_id)   "On Demand VOI",
  case coalesce(trim(enable_dollar_validation_2),'')
           when 'On' then 'Enable with Rejection'
           when 'Rejection Pending Workflow' then 'Enable with Pending Workflow'
           when 'Off' then 'Off' when '' then 'Off' end 
  as enable_dollar_valication,   ---3

 
 case when  enable_voi_notification_emails = true then 'Yes' else 'No' end   as send_notification_emails   ,  ---4

             an.enable_voi_embedded as "Enable VOI Embedded", 


  case when  trim ((select max(vendor_name) from services  where ns.voi = service_id))='The Work Number SSV' then  ssv_num_months
when  trim ((select max(vendor_name) from services  where ns.voi_ondmd = service_id))='The Work Number SSV' then  ssv_num_months

end ssv_num_of_month, 

   ts_service_tier "Worknumber SSV Select", ts_selected_tier "Worknumber SSV Tier",ts_on_demand "Worknumber SSV On Demand",
   py_service_tier "Paystub Select" ,py_selected_tier "Paystub Tier",py_on_demand "Paystub On Demand",
   av_service_tier "Asset Verification Select",av_selected_tier "Asset Verification Tier",av_on_demand "Asset Verification On Demand",
   vs_ssv_num_months as "TWN SSV Number of Months",

-- voi

   (select max(vendor_name) from services  where ns.cred_ondmd = service_id)   on_demand_credit,
(select max(vendor_name) from services  where ns.civilcrt_ondmd = service_id)   on_demand_civil_court,
  (select max(vendor_name) from services  where ns.rhist_ondmd = service_id)   on_demand_rental,
(select max(vendor_name) from services  where ns.ofac_ondmd = service_id)   "On Demand OFAC",
  (select max(vendor_name) from services  where ns.crim_ondmd = service_id)   on_demand_crim,
(select max(vendor_name) from services  where ns.cust_crim_ondmd = service_id)   on_demand_cust_crim,
(select max(vendor_name) from services  where ns.offense_ondmd = service_id)   on_demand_Offense,
 


	(case when ns.reeval_func=19 then 'On' else null end) on_demand_reevaluation_request,
	(case when ns.sbond_func=20 then 'On' else null end) on_demand_sure_deposit,

 
  
/*Group Scoring */
	CASE 	WHEN gsm.method_description ISNULL THEN gsm2.method_description
									ELSE gsm.method_description 
									END AS group_scoring, 
	CASE 	WHEN gsm.method_description ISNULL THEN gsuco.method_parameter_1
									ELSE gsu.method_parameter_1
									END AS scoring_table, 

/*Credit Report Options*/
	UPPER(CAST((credit_report&1 > 0) AS TEXT)) AS suppress_details, 
UPPER(CAST((credit_report&4 > 0) AS TEXT)) AS Suppress_Items_for_Review,
 an.suppress_voyager_page as suppress_applicant_file,   an.node_suppress_validation_workflow,  an.node_suppress_cancel_voi, voi_manual_additional_income, 
 UPPER(CAST (include_logo as text)) include_logo ,
 
        UPPER(CAST((an.rental_history&4>0) AS TEXT)) AS show_tenant_information,
        UPPER(CAST((an.rental_history&8>0) AS TEXT)) AS show_collections,
        UPPER(CAST((an.rental_history&16>0) AS TEXT)) AS show_statement,
UPPER(CAST((an.rental_history&128>0) AS TEXT)) AS  rental_history_scoring, --RHscoring from NodeInfo where node_code = 'RGDE1';
	UPPER(CAST((an.rental_history&64>0) AS TEXT)) AS show_reasons,
	UPPER(CAST((rhs.late_payments_scoring) AS TEXT)) AS late_payments_scoring, 
	rhs.late_payments_max, 
	rhs.late_payments_period, 
	UPPER(CAST(rhs.nsf_scoring AS TEXT)) AS nsf_scoring, 
	rhs.nsf_max, 
	rhs.nsf_period, 
	UPPER(CAST(rhs.outstanding_balances_scoring AS TEXT)) AS outstanding_balances_scoring, 
	rhs.outstanding_balances_max, 
	rhs.outstanding_balances_period, 
	UPPER(CAST(rhs.write_offs_scoring AS TEXT)) AS write_offs_scoring, 
	rhs.write_offs_amount, 
	rhs.write_offs_period, 
	UPPER(CAST(rhs.collections_scoring AS TEXT)) AS collections_scoring, 
	rhs.collections_amount, 
	rhs.collections_period,

/*Interface Options */
/*new addition */
case yardi_interface when 1 then 'One-Way Interface' when 2 then 'Two-Way Interface' end as interface_functionality,
(select display_value  from list_values where list_id =1 and "index" = pms_if limit 1) "PMS Interface",
 (select display_value  from list_values where list_id =2 and "index" = oli_if limit 1) Online_Leasing_Interface,
 	CASE	WHEN an.disable_edit_applicant=0 THEN 'Do Not Disable'
				WHEN an.disable_edit_applicant=1 THEN 'Edit Applicant'
				WHEN an.disable_edit_applicant=2 THEN 'Enter New Applicant'
				WHEN an.disable_edit_applicant=3 THEN 'Edit and Enter New Applicant'
					ELSE NULL
					END AS disable_edit_applicant,

  	UPPER(CAST(include_checkpoint_msg as text)) include_checkpoint_msg,
	UPPER(CAST(an.enable_group_apps AS TEXT)) AS enable_group_apps,
	ylp.interface_username, 
	ylp.server_name AS database_server_name, 
	ylp.database_name, 
	ylp.platform, 
	ylp.interface_entity, 
	ylp.web_service_uri AS "URL", 
/* Risk Score Options */
	rs.a_active, 
	rs.a_startbreakpoint, 
	rs.a_endbreakpoint,
	rs.b_active, 
	rs.b_startbreakpoint,
	rs.b_endbreakpoint,
	rs.c_active, 
	rs.c_startbreakpoint,
	rs.c_endbreakpoint, 
	rs.d_active, 
	rs.d_startbreakpoint, 
	rs.d_endbreakpoint, 
	rs.f_active,
	rs.f_startbreakpoint, 
	rs.f_endbreakpoint, 
	rs.g_active, 
	rs.g_startbreakpoint, 
	rs.g_endbreakpoint, 
	rs.r_active, 
	rs.r_startbreakpoint, 
	rs.r_endbreakpoint,
	an.def_credit_score  AS no_risk_score_available,
   to_char(nt.create_stamp , 'MM/DD/YYYY'::text)     AS setup_date,

	

  /*General Continued*/
 
 
 /* Property Information*/

/*aal setting */
case when icraa = 'No'  then 'Disable' when icraa = 'Yes' then 'Enabled' else  'Not Applicable' end as ICRAA, 
 case when crim_email='t' then 'Yes' else 'No' end   as  "Email Conditional Offer of Pending Housing", 
case when crim_pre_aal = 'Yes' then 'Enabled' else 'Disabled' end as  PreAdverse_Action_Letter, 
case when preaal_template_id=36 then 'Pre-AAL Rev 121819' else null end  as  PreAdverse_Action_Letter_Template, 
reconsider_request_period || ' ' ||  reconsider_period_in as Reconsideration_Request_Period, 
reconsider_review_period || ' ' || reconsider_period_in as Reconsideration_Review_Period,
  nova_credit_scoring, nova_risk_score,  case when coalesce(sso_enabled,'false') = 'false' then 'No' else 'Yes' end sso_enabled,
link_node as Alternate_Screening_Criteria,case when coalesce(an.b2b_screening_workflow,false)= false then 'Off' else 'On' end as Alternate_Screening_Criteria_Rules, 
case when credit_report<260 then 'False' when credit_report>=260 then 'True'   end as Suppress_Items_to_Review , last_applicant,
case an.hold_voyager_move_in when 'OFAC' then 'OFAC Messages' when 'Checkpoint' then 'Checkpoint Messages' when 'OR' then 'Either Of' else an.hold_voyager_move_in end as Review_Report_Acknowledgment,

/*contact info*/
      an.suspension_contact_name as activity_contact_name, an.suspension_email as activity_email,   an.suspension_phone as activity_phone, an.suspension_fax as activity_fax,
      nt.contact_name,   COALESCE(nt.contact_phone,'') ||  COALESCE(nt.contact_phone_extension,'') as  contact_phone, nt.fax_phone,  
       email_address as contact_email, email_address_man as manager_email,
	CASE WHEN an.email_fcra_letter=1 THEN 'No'
			 WHEN an.email_fcra_letter=2 THEN 'Applicant and Property'

			 WHEN an.email_fcra_letter=3 THEN 'Property Only'
			 WHEN an.email_fcra_letter=4 THEN 'Applicant Only'
						ELSE NULL
						END as Email_FCRA_Letters_To, 
  	an.fcra_email_addr as Adverse_Action_Email_Address,  
region, manager, manager2, manager3,   property_type_description  as property_type, 
/*sure deposit */
CASE WHEN sbond_func =20 THEN 'Yes' else 'No' END as Surety_Bond, 

case  when enable_new_suredep_v =false then null else case sure_deposit_layout when 0 THEN 'Normal' when 1 THEN 'Property Configured for Overrides' END end Surety_Bond_Layout ,
case  when enable_new_suredep_v =false then null else case surety_bond_vENDor when 0 THEN 'Not Set' when 1 THEN 'SureDeposit' when 2 THEN 'DepositIQ' END end Surety_Bond_Vendor,
case coalesce(pandemic_era_civil_court_filter,'') when '' then 'Disabled' 
when '4' then 'Enable Both' 
when '0' then 'Disabled'
when '1' then 'Enable Civil Court'  
when '2' then 'Enable Rental History' 
end as pandemic_era_civil_court_filter,  to_char( pandemic_start_date , 'MM/DD/YYYY' ) pandemic_start_date,  to_char( pandemic_end_date , 'MM/DD/YYYY' )  as pandemic_end_date,
              ac.account_manager
  

FROM node_tbl nt
LEFT JOIN account_tbl as BA ON nt.node_code = BA.account_code
LEFT JOIN auxiliary_node an ON nt.node_code=an.node_code
LEFT JOIN credit_products cp ON an.cred_prod_id=cp.cred_prod_id
LEFT JOIN group_score_users gsu ON gsu.node_or_company_code=nt.node_code
LEFT JOIN group_score_methods gsm ON gsu.method=gsm.method
LEFT JOIN group_score_users gsuco ON gsuco.node_or_company_code=nt.company_code /* We join the group scoring tables again by company for those nodes that take on the company score*/
LEFT JOIN group_score_methods gsm2 ON gsuco.method=gsm2.method
LEFT JOIN auxiliary_company ac ON nt.company_code=ac.company_code
LEFT JOIN rhs_policies rhs ON nt.node_code=rhs.policy_code
LEFT JOIN yardi_login_parameters ylp ON nt.node_code=ylp.node_code
 

 
LEFT JOIN (
					SELECT 
						rpd.node_code, 
						CASE 	WHEN start_break_point ISNULL THEN 'FALSE'
									ELSE 'TRUE' 
									END AS a_active, 	
						rpd.start_break_point AS a_startbreakpoint, 
						rpd.end_break_point AS a_endbreakpoint, 
						CASE 	WHEN brpd.b_startbreakpoint ISNULL THEN 'FALSE'
									ELSE 'TRUE' 
									END AS b_active, 	
						brpd.b_startbreakpoint,
						brpd.b_endbreakpoint, 
						CASE 	WHEN crpd.c_startbreakpoint ISNULL THEN 'FALSE'
									ELSE 'TRUE' 
									END AS c_active, 	
						crpd.c_startbreakpoint,
						crpd.c_endbreakpoint, 
						CASE 	WHEN drpd.d_startbreakpoint ISNULL THEN 'FALSE'
									ELSE 'TRUE' 
									END AS d_active,
						drpd.d_startbreakpoint,
						drpd.d_endbreakpoint, 
						CASE 	WHEN frpd.f_startbreakpoint ISNULL THEN 'FALSE'
									ELSE 'TRUE' 
									END AS f_active,
						frpd.f_startbreakpoint,
						frpd.f_endbreakpoint, 
						CASE 	WHEN grpd.g_startbreakpoint ISNULL THEN 'FALSE'
									ELSE 'TRUE' 
									END AS g_active,
						grpd.g_startbreakpoint,
						grpd.g_endbreakpoint, 
						CASE 	WHEN rrpd.r_startbreakpoint ISNULL THEN 'FALSE'
									ELSE 'TRUE' 
									END AS r_active,
						rrpd.r_startbreakpoint,
						rrpd.r_endbreakpoint
					FROM 	risk_score_policy_details rpd
					LEFT JOIN 	(
											SELECT 
												node_code, 
												CASE 	WHEN start_break_point ISNULL THEN 'FALSE'
															ELSE 'TRUE' 
															END AS b_active,
												start_break_point AS b_startbreakpoint, 
												end_break_point AS b_endbreakpoint	
											FROM 	risk_score_policy_details
											WHERE score = 'B'
											) brpd ON brpd.node_code=rpd.node_code
					LEFT JOIN 	(
											SELECT 
												node_code,  
												CASE 	WHEN start_break_point ISNULL THEN 'FALSE'
															ELSE 'TRUE' 
															END AS c_active,
												start_break_point AS c_startbreakpoint, 
												end_break_point AS c_endbreakpoint	
											FROM 	risk_score_policy_details
											WHERE score = 'C'
											) crpd ON crpd.node_code=rpd.node_code
					LEFT JOIN 	(
											SELECT 
												node_code,  
												CASE 	WHEN start_break_point ISNULL THEN 'FALSE'
															ELSE 'TRUE' 
															END AS d_active,
												start_break_point AS d_startbreakpoint, 
												end_break_point AS d_endbreakpoint	
											FROM 	risk_score_policy_details
											WHERE score = 'D'
											) drpd ON drpd.node_code=rpd.node_code
					LEFT JOIN 	(
											SELECT 
												node_code, 
												CASE 	WHEN start_break_point ISNULL THEN 'FALSE'
															ELSE 'TRUE' 
															END AS f_active,
												start_break_point AS f_startbreakpoint, 
												end_break_point AS f_endbreakpoint	
											FROM 	risk_score_policy_details
											WHERE score = 'F'
											) frpd ON frpd.node_code=rpd.node_code
					LEFT JOIN 	(
											SELECT 
												node_code, 
												CASE 	WHEN start_break_point ISNULL THEN 'FALSE'
															ELSE 'TRUE' 
															END AS g_active,
												start_break_point AS g_startbreakpoint, 
												end_break_point AS g_endbreakpoint	
											FROM 	risk_score_policy_details
											WHERE score = 'G'
											 
											) grpd ON grpd.node_code=rpd.node_code
					LEFT JOIN 	(
											SELECT 
												node_code, 
												CASE 	WHEN start_break_point ISNULL THEN 'FALSE'
															ELSE 'TRUE' 
															END AS r_active,
												start_break_point AS r_startbreakpoint, 
												end_break_point AS r_endbreakpoint	
											FROM 	risk_score_policy_details
											WHERE score = 'R'
											) rrpd ON rrpd.node_code=rpd.node_code
      
					WHERE score = 'A'
					) AS rs ON rs.node_code=nt.node_code
         LEFT JOIN (select primary_node, array_agg(link_node)::text link_node from b2b_node  where active<>'N' and primary_node <> link_node  group by primary_node) Link_Node on primary_node = nt.node_code 
            
         LEFT JOIN (
                      SELECT comp_code, node_code, autoemail_fcraletter, hold_until_active_app_processs, send_fcraletter_after_hours, created_on::date 
                        FROM autogenerate_fcraletter_config --where node_code = 'K7956'

             ) fc on fc.node_code = nt.node_code and fc.comp_code = nt.company_code
             LEFT JOIN node_services ns on    trim(ns.pr_code) = trim(nt.node_code) and trim(ns.co_code) = trim(nt.company_code)
             LEFT JOIN property_types pts on pts.property_type=an.property_type

                 left join    (select node_code, service_tier as ts_service_tier,  selected_tier ts_selected_tier, on_demand as ts_on_demand
            from vs_node_services where  service_code = 'VS_TWNSSV') ts_service on nt.node_code = ts_service.node_code
          
        LEFT JOIN (select node_code, service_tier as  py_service_tier,  selected_tier  py_selected_tier, on_demand as  py_on_demand
            from vs_node_services where  service_code = 'VS_REQ_PSTB') py_service on nt.node_code = py_service.node_code

        LEFT JOIN (select node_code, service_tier as av_service_tier,  selected_tier av_selected_tier, on_demand as av_on_demand
            from vs_node_services where  service_code = 'VS_ASSET_VER') av_service on nt.node_code = av_service.node_code

WHERE nt.canceled <> 'T'   
and  nt.company_code  not in ('APP')
ORDER BY nt.company_code, nt.node_code   
   
  

"""  )
    return SQL

def generate_address_sql():
     
    SQL =    ( """   select company_code, node_tbl.node_code,  node_name,  node_street_1  ,  node_street_2  ,  node_city,  node_state,   
              --''''||coalesce(node_zip,'') as 
              node_zip, market_rate_units,
              email_for_bulk_invoices as "E-mail For Bulk Invoices"
From node_tbl left join auxiliary_node on auxiliary_node.node_code = node_tbl.node_code  where   canceled = false
and    company_code  not in ('APP')
 
    
order by 1,   3,  2  
 

"""  )
    return SQL



def generate_civil_sql ():
    SQL =    ( """  select company_tbl.company_code, company_tbl.company_name,  policy_name,
case when COALESCE(record_type,'')='' and record_seq >0 and record_seq<4 
  then (select record_type from rs_policy_civilcourt_record aa where aa.policy_id =rs_policy_civilcourt_record.policy_id and record_seq =0 and COALESCE(record_type,'')<>'' limit 1) else record_type end record_type,
problem_type, quantity, timeframe, minimum_value, recommendation ,policy_ref_id as policy_id --, record_seq
from    rs_policy_details, 
rs_policy_civilcourt_record , company_tbl
where rs_policy_civilcourt_record.policy_id =rs_policy_details.policy_id     and  record_seq<>0 and company_tbl.company_code = rs_policy_details.company_code
and  (trim(lower(policy_status)) in ( 'draft', 'active',  'ppreview')  and trim(lower(status)) ='active' ) 
and    company_tbl.company_code not in ( 'APP','APP2','APPCK','AUTO','BETAA','BZ002','BZ655','BZ805','BZ890','BZ907','BZ922','BZ965','BZB27','BZB92','DEMO','FAKE','GEN2','GREYT','INTFC','INTF2','INTF3','INTF4','RCCRM','RICHC','RICHD','SALES','TEST ','TEST9','TRAIN','TSEXE','TSTER','VOYA','XTEST','YASC','YRK')   and company_tbl.active ='True'
    and company_tbl.company_code not in (select company_code from node_tbl group by company_code having max(canceled::int) = min(canceled::int) and min(canceled::int)=1)
         and not exists (select attachmentname, company_code  from attachment_property_reference,  attachment_property_master 
	
	where  attachment_property_reference.attachmentpropmasterid =attachment_property_master."id"
	and  refid in (select max(refid) from attachment_property_reference, attachment_property_master 
	where  attachment_property_reference.attachmentpropmasterid =attachment_property_master."id"
	 group by attachmentname, company_code )
	and lower(attachment_property_reference.status)  in ( 'inactive' )
	and attachment_property_master.company_code = company_tbl.company_code and policy_name = attachmentname)
and  company_tbl.company_code  not in ('APP')
order by 1, 2, 3,   record_seq   
--limit 11
             
  

"""  )
    return SQL

def generate_recommendation_sql ():
    SQL =    ( """  select company_tbl.company_code,   company_tbl.company_name,  policy_name,  initcap(grade) grade, initcap(risk) risk, initcap(recommendation) recommendation,
policy_ref_id as policy_id --, record_seq
 from rs_policy_details, rs_policy_grade_risk_recommendation, company_tbl
 where     record_seq<>0 and   (trim(lower(policy_status)) in ( 'draft', 'active',  'ppreview')  and trim(lower(status)) ='active' ) 
and rs_policy_grade_risk_recommendation.policy_id =rs_policy_details.policy_id and company_tbl.company_code = rs_policy_details.company_code
and company_tbl.company_code not in ( 'APP','APP2','APPCK','AUTO','BETAA','BZ002','BZ655','BZ805','BZ890','BZ907','BZ922','BZ965','BZB27','BZB92','DEMO','FAKE','GEN2','GREYT','INTFC','INTF2','INTF3','INTF4','RCCRM','RICHC','RICHD','SALES','TEST ','TEST9','TRAIN','TSEXE','TSTER','VOYA','XTEST','YASC','YRK') and company_tbl.active ='True'
 and not exists (select attachmentname, company_code  from attachment_property_reference,  attachment_property_master 
	
	where  attachment_property_reference.attachmentpropmasterid =attachment_property_master."id"
	and  refid in (select max(refid) from attachment_property_reference, attachment_property_master 
	where  attachment_property_reference.attachmentpropmasterid =attachment_property_master."id"
	 group by attachmentname, company_code )
	and lower(attachment_property_reference.status)  in ( 'inactive' )
	and attachment_property_master.company_code = company_tbl.company_code and policy_name = attachmentname)
               and company_tbl.company_code not in (select company_code from node_tbl group by company_code having max(canceled::int) = min(canceled::int) and min(canceled::int)=1)
and  company_tbl.company_code not in (
'APP','APP2','APPCK','AUTO','BETAA','BZ001 ','BZ002','BZ655','BZ689','BZ805','BZ890','BZ903','BZ907','BZ965','BZB22','BZB27','BZB92','BZC78','BZC79','BZF80','BZG46','DEMO','ESLES','FAKE','GEN2','GREYT','INTF2','INTF3','INTF4','INTFC','PRCRD','RCCRM','RGSAL','RGSW','RGTST','RICHC','RICHD','RICHE','RICHK','SALES','STIJL','SWDMO','SWRKS','TEST9','TESTR','TRAIN','TSEXE','TSTER','VOYA','XTEST','YASC','YRK','SYLJ','RICHT','2SYLJ','BZL65')
order by 1, 2, 3,  rs_policy_grade_risk_recommendation.record_seq    

    --  limit 11       
  

"""  )
    return SQL

def generate_aal_sql ():
    SQL =    ( """              select node_tbl.company_code, node_tbl.node_code, node_name, category, file_type, file_name, node_static_files_mapping.dc_workflow as "Conditional Offer Email", 
              node_static_files_mapping.preaal_workflow as "Tiered AAL Email"   from node_static_files_mapping, node_tbl,company_static_files  where --node_static_files_mapping.node_code =  'ED544'
                  node_tbl.node_code = node_static_files_mapping.node_code 
                 
								and node_static_files_mapping.company_static_files_id = company_static_files."id"
                and canceled = 'f' --and category = 'PREAAL' 
                and status = 't'
				and node_tbl.company_code not in ('APP')				 
                order by node_tbl.company_code, node_name, node_code
    --  limit 11       
  

"""  )
    return SQL

def generate_criminal_sql():
     
    SQL =    ( """   select       company_tbl.company_code,   company_tbl.company_name,  policy_name,
    offense,  felony ,     pending_felony ,   
rs_policy_criminal_record.misdemeamor ,  rs_policy_criminal_record.pending_misdemeamor, pattern_of_misdemeamor, 
 case when length(return_records) =0    then result1 
      when length(return_records) BETWEEN 1 and 3 and return_records <>'ANY'  then return_records 
      else initcap(return_records) end  return_records    ,  policy_ref_id as policy_id --, record_seq
from    rs_policy_details inner join rs_policy_criminal_record on rs_policy_criminal_record.policy_id =rs_policy_details.policy_id
 
inner join company_tbl on company_tbl.company_code = rs_policy_details.company_code
left join   (
select  case when return_records = 'ANY' then 'Any' when length(return_records)<=3 then return_records else  initcap( return_records) end  result1, policy_id
from rs_policy_criminal_record where record_seq = 1 and COALESCE(return_records,'')<>'') pol_desc
 on  pol_desc.policy_id = rs_policy_criminal_record.policy_id 
where rs_policy_criminal_record.policy_id =rs_policy_details.policy_id  and     (trim(lower(policy_status)) in ( 'draft', 'active',  'ppreview')  and trim(lower(status)) ='active' )    and  record_seq<>0
  and company_tbl.company_code = rs_policy_details.company_code and company_tbl.active ='True'
and    company_tbl.company_code not in ( 'APP','APP2','APPCK','AUTO','BETAA','BZ002','BZ655','BZ805','BZ890','BZ907','BZ922','BZ965','BZB27','BZB92','DEMO','FAKE','GEN2','GREYT','INTFC','INTF2','INTF3','INTF4','RCCRM','RICHC','RICHD','SALES','TEST ','TEST9','TRAIN','TSEXE','TSTER','VOYA','XTEST','YASC','YRK')   
    and company_tbl.company_code not in (select company_code from node_tbl group by company_code having max(canceled::int) = min(canceled::int) and min(canceled::int)=1)
               and not exists (select attachmentname, company_code  from attachment_property_reference,  attachment_property_master 
	
	where  attachment_property_reference.attachmentpropmasterid =attachment_property_master."id"
	and  refid in (select max(refid) from attachment_property_reference, attachment_property_master 
	where  attachment_property_reference.attachmentpropmasterid =attachment_property_master."id"
	 group by attachmentname, company_code )
	and lower(attachment_property_reference.status)  in ( 'inactive' )
	and attachment_property_master.company_code = company_tbl.company_code and policy_name = attachmentname)
and  company_tbl.company_code  not in ('APP')
 order by 1, 2, 3,   record_seq     

 -- limit 11

"""  )
    return SQL


def connect_rentgrow_data_frame(cus_sql):
     
    conn = None
 
    params = config()
    conn = psycopg2.connect(**params)       
    cur = conn.cursor()
  
   
    
    string_replace = "  ('" + "','".join(name.upper().strip().replace("'", r"\'") for name in list_term_code) + "'"  + ')'
   # SQL_1 =  (o_SQL).replace( regexp=True, to = "request.company_code not in ('APP') ", value = string_replace    )
    cus_sql =  (cus_sql).replace(   "('APP')",   string_replace     ) #live
    
    SQL =  cus_sql
    cur.execute(SQL) 
    col = [i[0] for i in cur.description]
    conn.set_client_encoding('ISO-8859-1') 
    rows = cur.fetchall()
    cur.close()
    data = pd.DataFrame(rows, columns = col)
   
    return data
# Print the list 
  

def format_book(wb, ws, df_db, node_code=""):
          
    #for x in range(1,len(df_db.columns)+1):
         
      #  ws.cell(1, x).fill = redFill     
      #  ws.cell(1, x).font = Font(color="FFFFFF", name="Verdana", size=12)   
    
    #ws.auto_filter.ref = 'A1:W1'
    ws.auto_filter =True
    if (node_code == "Node"):
        rows = 1
        columns = 8
    elif  (node_code == "Company"):
        rows = 1
        columns = 1
    else:
        rows = 1
        columns=0
 
    ws.panes = pyexcelerate.Panes(columns, rows)
    ''' 
    if (node_code == "Node" or node_code == "Company"):
        style = Style(fill=Fill(background=Color(135,206,250)))
        for row in range(1, len(df_db) + 1, 2):
            ws.set_row_style(row, style)
    '''
    #ws.set_row_style(1, Style(fill=Fill(background=Color(192,192,192))))
    #ws.get_row_style(1).fill.background = Color(192,192,192)
    ws.get_row_style(1).fill.background = Color (91, 155, 213)
    ws.get_row_style(1).font.bold = True
    ws.get_row_style(1).font.color = Color(255,255,255)
    #ws.get_row_style(1).alignment.wrap_text = True 
    
    if (node_code == "Node"): 
        for col in range (68,85):
            ws.set_cell_style(1, col, Style(fill=Fill(background=Color(255,218,185))))
            ws.set_col_style(col, Style(size=25))
    elif (node_code == "Company"): 
            for col in range(1, len(df_db.columns)+1,1):
                ws.set_col_style(col, Style(size=25))
    else:
        for col in range(1, len(df_db.columns)+1,1):
 
           ws.set_col_style(col, Style(size=-1))

   # if  (node_code == "Company" ): 
        #    ws.set_col_style(27, Style(size=25))  #Custom_Conditional_Text
        #    ws.set_col_style(28, Style(size=25))
        #    ws.set_col_style(19, Style(size=25))  #tsr email
        #    ws.set_col_style(20, Style(size=25))
           #ws.set_col_style(col, Style(alignment=Alignment( horizontal = 'left') ))  

    #for col in range(1, len(df_db.columns)+1,1):
        #ws.set_col_style(col, Style(size=-1))
    #    ws.set_col_style(col, Style(size=20))
        #ws.set_col_style(col, Style(alignment=Alignment( horizontal = 'left') )) and wrap not work
   
 
    '''   
    if (node_code == "Company" or node_code == "Node"):
       
        for col in range(1, len(df_db.columns)+1,1):
        #ws.set_col_style(col, Style(size=-1))
            #ws.set_col_style(col, Style(size=20))
            pass

    else :    
       
        for col in range(1, len(df_db.columns)+1,1):
 
            ws.set_col_style(col, Style(size=-1))
           # ws.set_col_style(col, Style(alignment=Alignment( horizontal = 'left') ))  
'''
def send_to_finance(filename1="", filename2=""):
    
    signature_file = "C:\App\Data\Sig\sig.htm"
    if os.path.exists(signature_file):
        
        #signature_path = os.path.join((os.environ['USERPROFILE']), sig_files_path) # Finds the path to Outlook signature files with signature name "Work"
        #html_doc = os.path.join((os.environ['USERPROFILE']),sig_html_path)     #Specifies the name of the HTML version of the stored signature
        #html_doc = html_doc.replace('\\\\', '\\')


        html_file = codecs.open(signature_file, 'r', 'utf-8', errors='ignore') #Opens HTML file and converts to UTF-8, ignoring errors
        signature_code = html_file.read()               #Writes contents of HTML signature file to a string

        #signature_code = signature_code.replace((signature_name + '_files/'), signature_path)      #Replaces local directory with full directory path
        html_file.close()

    
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = "xiaobin.zhang@yardi.com"

        
        mail.Subject = "Weekly Setup Audit - RS_AM@Yardi.Com ; rs_is@yardi.com;   "  
        
        #filename2 =  filename1.replace('O:\\ANALYTICS\\New Property Lists\\', '\\\ysifwfs07\\Vol2\ANALYTICS\\New Property Lists\\') # + filename.replace("/", ".") + ".xlsx"
            #path  = "\"\\\\windows_Server\\golobal_directory\\the folder\\file yyymm.xlsx\""
        path = '"' + filename1 + '"'
        string  = """<a href=""" +  path + ' style=text-decoration: none>' + filename1 +  '<' +  r'\a'  + '>'
        string =  string.replace('O:\\ANALYTICS\\Setup_Audit\\', '\\\ysifwfs07\\Vol2\ANALYTICS\\Setup_Audit\\')

        path_2 = '"' + filename2 + '"'
        string_2  = """<a href=""" +  path_2 + ' style=text-decoration: none>' + filename2 +  '<' +  r'\a'  + '>'
        string_2 =  string_2.replace('O:\\ANALYTICS\\Setup_Audit\\', '\\\ysifwfs07\\Vol2\ANALYTICS\\Setup_Audit\\')
        
        #  string.replace('\\a>', '\a>')
        #mail.body = string
        
        mail.HTMLbody =   string + " <BR> " + string_2 +" <BR><BR><BR> "  +signature_code + " <BR><BR><BR> "
    
    
        mail.send


def columnToLetter(column):
    letter = ''
    while column > 0:
        temp = (column - 1) % 26
        letter = chr(temp + 65) + letter
        column = (column - temp - 1) // 26
    return letter
 
def letterToColumn(letter):
    column = 0
    length = len(letter)
    for i in range(length):
        column += (ord(letter[i].upper()) - 64) * 26**(length - i - 1)
    return column

 

def write_to_book(wb, df, sheet_name):
    values = [df.columns] + list(df.values)
    
    ws = wb.new_sheet(sheet_name, data=values)
    
    format_book(wb, ws,df, sheet_name )

def kill_excel():
    
    for proc in psutil.process_iter():
        if proc.name() == "EXCEL.EXE":
            proc.kill()
   
def dispatch(app_name:str):
    try:
        from win32com import client
       # app = client.gencache.EnsureDispatch(app_name)
        app = win32.dynamic.Dispatch("Excel.Application")
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    return app         
 
def run_part_two():
    wb = Workbook()
     

    df_new_address = connect_rentgrow_data_frame(generate_address_sql())
    df_new_address.replace( np.nan, '',inplace = True)
    write_to_book(wb, df_new_address, "Property Address")
    print ("address done")


    df_civil= connect_rentgrow_data_frame(generate_civil_sql())
    df_civil.replace( np.nan, '',inplace = True)
    write_to_book(wb, df_civil, "Civil")
    print ("civil done") 
 

    df_new_criminal = connect_rentgrow_data_frame (generate_criminal_sql())
    #df_new_company.to_excel (writer2,  sheet_name= 'Company', index=False, startrow=0,engine='xlsxwriter')
    df_new_criminal.replace( np.nan, '',inplace = True)
    write_to_book(wb, df_new_criminal, "Criminal")
    print ("Criminal done") 

    df_new_recommendation = connect_rentgrow_data_frame(generate_recommendation_sql())
    df_new_recommendation.replace( np.nan, '',inplace = True)
    write_to_book(wb, df_new_recommendation, "Recommendation")

    df_new_aal = connect_rentgrow_data_frame(generate_aal_sql())
    df_new_aal.replace( np.nan, '',inplace = True)
    write_to_book(wb, df_new_aal, "AAL Documents")


    

   

    wb.save(filename2_1)
 
   
    excel =   win32.dynamic.Dispatch("Excel.Application") 
    excel.Interactive = False
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False

    wb = excel.Workbooks.Open(filename2_1)
 
    wb.SaveAs(filename2_2, FileFormat=50)  # 50 is the code for XLSB file format
    excel.Application.Quit()
 
   # send_book (filename1, 'xiaobin.zhang@yardi.com' )

    if os.path.exists(filename2_2):
        os.remove(filename2_1)
        print("Successfully! The file has been removed")
    else:
        print("Cannot delete the file as it doesn't exist")    
        
def run_part_one():
     

    wb = Workbook()
 
    df_new_company = connect_rentgrow_data_frame (generate_company_sql())
   # list_am = connect_acct_manager()
    #df_new_company['Account Manager']=  df_new_company['Company_Code'].map(list_am.set_index('company_code')['account_manager']).fillna(df_new_company['Account Manager'])
    #df_new_company.to_excel (writer2,  sheet_name= 'Company', index=False, startrow=0,engine='xlsxwriter')
    #df_new_company.columns = [col.replace('_', ' ') for col in df_new_company.columns]
    df_new_company.replace( np.nan, '',inplace = True)
    write_to_book(wb, df_new_company, "Company")
    print ("comany done")
   


    df_new_prop = connect_rentgrow_data_frame(generate_node_sql())
    df_new_prop.replace( np.nan, '',inplace = True)
  
    write_to_book(wb, df_new_prop, "Node")
    print ("Node done") 


    df_template = pd.read_excel(file_template)
    write_to_book(wb, df_template, "What's New")

    wb.save(filename1_1)
 

    excel =  win32.dynamic.Dispatch("Excel.Application")
    excel.Interactive = False
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False

    wb = excel.Workbooks.Open(filename1_1)
 
    wb.SaveAs(filename1_1_tmp, FileFormat=50)  # 50 is the code for XLSB file format
    excel.Application.Quit()
 
   # send_book (filename1, 'xiaobin.zhang@yardi.com' )

    if os.path.exists(filename1_1_tmp):
        shutil.copyfile(filename1_1_tmp, filename1_2)

    if os.path.exists(filename1_2):
        os.remove(filename1_1)
        os.remove(filename1_1_tmp)
        print("Successfully! The file has been removed")
    else:
        print("Cannot delete the file as it doesn't exist")    


if __name__ == '__main__':
    
     #kill_excel()
    #shutil.copy (file_temp, filename1)
    
    #writer2=pd.ExcelWriter(filename1, engine='openpyxl' )
    #writer2=pd.ExcelWriter(filename1  )

 
    if os.path.exists(filename1_2):
        pass
    else:
          
        run_part_one()
        
        
    print ("Part One Done!!!")


    if os.path.exists(filename2_2):
        pass 
    else: 
        run_part_two()
    

    send_to_finance(filename1_2, filename2_2 ) 

    

    
sys.exit(0)
quit()
