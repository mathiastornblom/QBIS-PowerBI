// This file contains your Data Connector logic
section Qbis;
 
// Data Source Kind description
Qbis = [
	Authentication = [
		UsernamePassword = [
			UsernameLabel= Extension.LoadString("UsernameLabel"),
			PasswordLabel= Extension.LoadString("PasswordLabel")
		]
	],
	Label = Extension.LoadString("DataSourceLabel")
];

// Data Source UI publishing description
Qbis.Publish = [
	Beta = true,
	Category = "Other",
	ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
	//LearnMoreUrl = "https://powerbi.microsoft.com/",
	SourceImage = Qbis.Icons,
	SourceTypeImage = Qbis.Icons
];

Qbis.Icons = [
	Icon16 = { Extension.Contents("Qbis16.png"), Extension.Contents("Qbis20.png"), Extension.Contents("Qbis24.png"), Extension.Contents("Qbis32.png") },
	Icon32 = { Extension.Contents("Qbis32.png"), Extension.Contents("Qbis40.png"), Extension.Contents("Qbis48.png"), Extension.Contents("Qbis64.png") }
];

//
//Variables
//

//https://bin.qbis.se/WS/QBIS_DS.asmx
QbisActualURL = "https://apps.qbis.se/WS/QBIS_DS.asmx";

// https://bin.qbis.se/ws/qbis_mobileapp.asmx
QbisMobileActualURL = "https://bin.qbis.se/ws/qbis_mobileapp.asmx";

//
// Implementation
// 

[DataSource.Kind="Qbis", Publish="Qbis.Publish"]
shared Qbis.Contents =  Value.ReplaceType(QbisNavTable, type function () as any);

QbisNavTable = () as table =>
	let
		source = #table({"Name", "Data", "ItemKind", "ItemName", "IsLeaf"}, {
			{ Extension.LoadString("GetActivityMembers"), GetActivityMembers(), "Table", "Table", true },
			{ Extension.LoadString("GetActivityTime"), GetActivityTime(), "Table", "Table", true },
            { Extension.LoadString("GetArticles"), GetArticles(), "Table", "Table", true },
            { Extension.LoadString("GetCRMTypes"), GetCRMTypes(), "Table", "Table", true },
            { Extension.LoadString("GetCompanyInformation"), GetCompanyInformation(), "Table", "Table", true },
            { Extension.LoadString("GetContactWays"), GetContactWays(), "Table", "Table", true },
            { Extension.LoadString("GetCurrencies"), GetCurrencies(), "Table", "Table", true },
            { Extension.LoadString("GetCurrencyRates"), GetCurrencyRates(), "Table", "Table", true },
            { Extension.LoadString("GetCustomerContacts"), GetCustomerContacts(), "Table", "Table", true },
            { Extension.LoadString("GetCustomers"), GetCustomers(), "Table", "Table", true },
            { Extension.LoadString("GetDepartments"), GetDepartments(), "Table", "Table", true },
            { Extension.LoadString("GetEmployees"), GetEmployees(), "Table", "Table", true },
            { Extension.LoadString("GetExpenseCodes"), GetExpenseCodes(), "Table", "Table", true },
            { Extension.LoadString("GetExpenseItems"), GetExpenseItems(), "Table", "Table", true },
            { Extension.LoadString("GetExpenseSheets"), GetExpenseSheets(), "Table", "Table", true },
            { Extension.LoadString("GetJobs"), GetJobs(), "Table", "Table", true },
            { Extension.LoadString("GetLodgingItems"), GetLodgingItems(), "Table", "Table", true },
            { Extension.LoadString("GetLodgingTypes"), GetLodgingTypes(), "Table", "Table", true },
            { Extension.LoadString("GetModules"), GetModules(), "Table", "Table", true },
            { Extension.LoadString("GetOpportunities"), GetOpportunities(), "Table", "Table", true },
            { Extension.LoadString("GetPermissions"), GetPermissions(), "Table", "Table", true },   
            { Extension.LoadString("GetPositionHistories"), GetPositionHistories(), "Table", "Table", true },  
            { Extension.LoadString("GetPositions"), GetPositions(), "Table", "Table", true },  
            { Extension.LoadString("GetProjectActivities"), GetProjectActivities(), "Table", "Table", true },
            { Extension.LoadString("GetProjectMaterials"), GetProjectMaterials(), "Table", "Table", true },
            { Extension.LoadString("GetProjectMembers"), GetProjectMembers(), "Table", "Table", true },
            { Extension.LoadString("GetProjectMilestones"), GetProjectMilestones(), "Table", "Table", true },   
            { Extension.LoadString("GetProjectPhases"), GetProjectPhases(), "Table", "Table", true },   
            { Extension.LoadString("GetProjects"), GetProjects(), "Table", "Table", true },         
            { Extension.LoadString("GetSalaryActivities"), GetSalaryActivities(), "Table", "Table", true }, 
            { Extension.LoadString("GetSalaryTime"), GetSalaryTime(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleAgreementActivities"), GetScheduleAgreementActivities(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleAgreementSpecialDays"), GetScheduleAgreementSpecialDays(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleAgreements"), GetScheduleAgreements(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleBreaks"), GetScheduleBreaks(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleDays"), GetScheduleDays(), "Table", "Table", true }, 
            //{ Extension.LoadString("GetScheduleDepartments"), GetScheduleDepartments(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleEmployees"), GetScheduleEmployees(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleIntervals"), GetScheduleIntervals(), "Table", "Table", true }, 
            //{ Extension.LoadString("GetSchedulePositions"), GetSchedulePositions(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleSalaryActivities"), GetScheduleSalaryActivities(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleSalaryActivityDays"), GetScheduleSalaryActivityDays(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleSalaryActivityWeeks"), GetScheduleSalaryActivityWeeks(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleTimes"), GetScheduleTimes(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleVersions"), GetScheduleVersions(), "Table", "Table", true }, 
            { Extension.LoadString("GetScheduleWeeks"), GetScheduleWeeks(), "Table", "Table", true }, 
            { Extension.LoadString("GetSchedules"), GetSchedules(), "Table", "Table", true }, 
            { Extension.LoadString("GetSuppliers"), GetSuppliers(), "Table", "Table", true },          
            { Extension.LoadString("GetTasks"), GetTasks(), "Table", "Table", true },          
            { Extension.LoadString("GetTravelItems"), GetTravelItems(), "Table", "Table", true },          
            { Extension.LoadString("GetTravelTypes"), GetTravelTypes(), "Table", "Table", true },          
            { Extension.LoadString("GetUnitTypes"), GetUnitTypes(), "Table", "Table", true },          
            { Extension.LoadString("GetWorkingTime"), GetWorkingTime(), "Table", "Table", true },         

            //Mobile
            { Extension.LoadString("GeTeams"), GeTeams(), "Table", "Table", true },
            { Extension.LoadString("GeTeamsMembers"), GeTeamsMembers(), "Table", "Table", true },
            { Extension.LoadString("GetCRMCategories"), GetCRMCategories(), "Table", "Table", true },
            //{ Extension.LoadString("GetCRMCompanies"), GetCRMCompanies(), "Table", "Table", true },
            //{ Extension.LoadString("GetCRMContacts"), GetCRMContacts(), "Table", "Table", true },
            //{ Extension.LoadString("GetCompanyCRMCategories"), GetCompanyCRMCategories(), "Table", "Table", true },
            //{ Extension.LoadString("GetCompanySDCategories"), GetCompanySDCategories(), "Table", "Table", true },
            //{ Extension.LoadString("GetContactTypes"), GetContactTypes(), "Table", "Table", true },
            { Extension.LoadString("GetPipelinePercentages"), GetPipelinePercentages(), "Table", "Table", true },
            { Extension.LoadString("GetPipelineStages"), GetPipelineStages(), "Table", "Table", true },
            { Extension.LoadString("GetPriorityLevels"), GetPriorityLevels(), "Table", "Table", true },
            //{ Extension.LoadString("GetProducts"), GetProducts(), "Table", "Table", true },
            //{ Extension.LoadString("GetSalesProjects"), GetSalesProjects(), "Table", "Table", true },
            //{ Extension.LoadString("GetServiceDeskContacts"), GetServiceDeskContacts(), "Table", "Table", true },
            { Extension.LoadString("GetServiceDeskRequestCategories"), GetServiceDeskRequestCategories(), "Table", "Table", true },
            //{ Extension.LoadString("GetServiceDeskRequests"), GetServiceDeskRequests(), "Table", "Table", true },
            { Extension.LoadString("GetServiceDeskSchedue"), GetServiceDeskSchedue(), "Table", "Table", true },
            { Extension.LoadString("GetServiceDeskSettings"), GetServiceDeskSettings(), "Table", "Table", true },
            //{ Extension.LoadString("GetServiceDeskStatusLog"), GetServiceDeskStatusLog(), "Table", "Table", true },
            //{ Extension.LoadString("GetServiceRequestsActions"), GetServiceRequestsActions(), "Table", "Table", true },
            //{ Extension.LoadString("GetServiceRequestsMaterial"), GetServiceRequestsMaterial(), "Table", "Table", true },
            //{ Extension.LoadString("GetServicedeskCompanies"), GetServicedeskCompanies(), "Table", "Table", true },
            { Extension.LoadString("GetStatusTypes"), GetStatusTypes(), "Table", "Table", true },
            { Extension.LoadString("GetTeamsCategories"), GetTeamsCategories(), "Table", "Table", true }          
 		}),
		navTable = Table.ToNavigationTable(source, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
	in
		navTable;

GetActivityMembers = () as table =>
let
   source = fQBISRestAPI("GetActivityMembers"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"activity_member_activity_id", Int64.Type}, {"activity_member_employee_id", Int64.Type}, {"activity_member_allocated_minutes", Int64.Type}, {"activity_member_completed_by_member", type logical}, {"activity_member_completed_by_manager", type logical}, {"activity_member_cost_per_hour", Int64.Type}, {"activity_member_price_per_hour", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetActivityTime = () as table =>
let
   source = fQBISRestAPI("GetActivityTime"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded activity_time_notes_internal" = fExpand(Table6, "activity_time_notes_internal"),
    #"Expanded activity_time_notes_external" = fExpand(#"Expanded activity_time_notes_internal", "activity_time_notes_external"),   
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded activity_time_notes_external",{{"ACTIVITY_TIME_ID", Int64.Type}, {"activity_time_activity_id", Int64.Type}, {"activity_time_employee_id", Int64.Type}, {"activity_time_minutes", Int64.Type}, {"activity_time_notes_external", type text}, {"activity_time_notes_internal", type text}, {"activity_time_factor", type text}, {"activity_time_date", type datetimezone}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"activity_time_factor"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"activity_time_factor", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetArticles = () as table =>
let
   source = fQBISRestAPI("GetArticles"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded article_unit" = fExpand(Table6, "article_unit"),
    #"Expanded article_notes" = fExpand(#"Expanded article_unit", "article_notes"),
    #"Expanded article_gl_code" = fExpand(#"Expanded article_notes", "article_gl_code"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded article_gl_code",{{"ARTICLE_ID", Int64.Type}, {"article_code", type text}, {"article_name", type text}, {"article_unit", type text}, {"article_cost", type text}, {"article_price", type text}, {"article_supplier_id", Int64.Type}, {"article_notes", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"article_cost", "article_price"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"article_cost", Currency.Type}, {"article_price", Currency.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetCRMTypes = () as table =>
let
   source = fQBISRestAPI("GetCRMTypes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"CRM_TYPE_ID", Int64.Type}, {"crm_type_name", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetCompanyInformation = () as table =>
let
   source = fQBISRestAPI("GetCompanyInformation"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"company_name", type text}, {"company_organisation_number", type text}, {"company_minimum_lunch", Int64.Type}, {"company_maximum_lunch", Int64.Type}, {"company_use_lunch_rule", type logical}, {"company_monday_schedule", Int64.Type}, {"company_tuesday_schedule", Int64.Type}, {"company_wednesday_schedule", Int64.Type}, {"company_thursday_schedule", Int64.Type}, {"company_friday_schedule", Int64.Type}, {"company_saturday_schedule", Int64.Type}, {"company_sunday_schedule", Int64.Type}, {"company_allow_timeclock", type logical}, {"company_enforce_lunch_timeclock", type logical}, {"company_enforce_lunch_min_timeclock", Int64.Type}, {"company_enforce_lunch_from_timeclock", type datetimezone}, {"company_show_inout", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetContactWays = () as table =>
let
   source = fQBISRestAPI("GetContactWays"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"CONTACT_WAY_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetCurrencies = () as table =>
let
   source = fQBISRestAPI("GetCurrencies"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"CURRENCY_ID", Int64.Type}, {"currency_name", type text}, {"currency_region", type text}, {"currency_default", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetCurrencyRates = () as table =>
let
   source = fQBISRestAPI("GetCurrencyRates"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"CURRENCY_RATE_ID", Int64.Type}, {"currency_rate_currency_id", Int64.Type}, {"currency_rate_rate", type text}, {"currency_rate_date", type datetimezone}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"currency_rate_rate"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"currency_rate_rate", Currency.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetCustomerContacts = () as table =>
let
   source = fQBISRestAPI("GetCustomerContacts"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded customer_contact_title" = fExpand(Table6, "customer_contact_title"),
    #"Expanded customer_contact_phone_work" = fExpand(#"Expanded customer_contact_title", "customer_contact_phone_work"),   
    #"Expanded customer_contact_phone_mobile"= fExpand(#"Expanded customer_contact_phone_work", "customer_contact_phone_mobile"),   
    #"Expanded customer_contact_email" = fExpand(#"Expanded customer_contact_phone_mobile", "customer_contact_email"),   
    #"Expanded customer_contact_social_number" = fExpand(#"Expanded customer_contact_email", "customer_contact_social_number"),  
    #"Expanded customer_contact_type" = fExpand(#"Expanded customer_contact_social_number", "customer_contact_type"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded customer_contact_type",{{"CUSTOMER_CONTACT_ID", Int64.Type}, {"customer_contact_customer_id", Int64.Type}, {"customer_contact_social_number", type text}, {"customer_contact_email", type text}, {"customer_contact_phone_mobile", type text}, {"customer_contact_phone_work", type text}, {"customer_contact_last_name", type text}, {"customer_contact_first_name", type text}, {"customer_contact_primary", type logical}, {"customer_contact_newsletter", type logical}, {"customer_contact_type", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetCustomers = () as table =>
let
   source = fQBISRestAPI("GetCustomers"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"CUSTOMER_ID", Int64.Type}, {"customer_active", type logical}, {"customer_name", type text}, {"customer_code", type text}, {"customer_abbreviated", type text}}),
    #"Expanded customer_address" = fExpand(#"Changed Type", "customer_address"),
    #"Expanded customer_zip_code" =  fExpand(#"Expanded customer_address", "customer_zip_code"),
    #"Expanded customer_city" = fExpand(#"Expanded customer_zip_code", "customer_city"),
    #"Expanded customer_country" = fExpand(#"Expanded customer_city", "customer_country"),    
    #"Expanded customer_phone" = fExpand(#"Expanded customer_country", "customer_phone"),
    #"Expanded customer_fax" = fExpand(#"Expanded customer_phone" , "customer_fax"),
    #"Expanded customer_email" = fExpand(#"Expanded customer_fax", "customer_email"),
    #"Expanded customer_homepage" = fExpand(#"Expanded customer_email", "customer_homepage"),    
    #"Expanded customer_organization_number" = fExpand(#"Expanded customer_homepage", "customer_organization_number"),
    #"Expanded customer_notes" = fExpand(#"Expanded customer_organization_number", "customer_notes"),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded customer_notes",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetDepartments = () as table =>
let
   source = fQBISRestAPI("GetDepartments"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"DEPARTMENT_ID", Int64.Type}, {"department_code", type text}, {"department_name", type text}, {"department_manager_id", Int64.Type}, {"department_manager_first_name", type text}, {"department_manager_last_name", type text}, {"department_active", type logical}, {"department_unit_type_id", Int64.Type}, {"department_parent_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetEmployees = () as table =>
let
   source = fQBISRestAPI("GetEmployees"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded employee_address" = fExpand(Table6, "employee_address"),
    #"Expanded employee_zip_code" = fExpand(#"Expanded employee_address", "employee_zip_code"),
    #"Expanded employee_city" = fExpand(#"Expanded employee_zip_code", "employee_city"),
    #"Expanded employee_country" = fExpand(#"Expanded employee_city", "employee_country"),
    #"Expanded employee_phone_home" = fExpand(#"Expanded employee_country", "employee_phone_home"),
    #"Expanded employee_phone_work" = fExpand(#"Expanded employee_phone_home" , "employee_phone_work"),
    #"Expanded employee_phone_mobile" = fExpand(#"Expanded employee_phone_work", "employee_phone_mobile"),
    #"Expanded employee_email" = fExpand(#"Expanded employee_phone_mobile", "employee_email"),
    #"Expanded employee_social_number" = fExpand( #"Expanded employee_email", "employee_social_number"),
    #"Expanded employee_abbreviated" = fExpand(#"Expanded employee_social_number", "employee_abbreviated"),
    #"Expanded employee_notes" = fExpand(#"Expanded employee_abbreviated", "employee_notes"),
    #"Expanded employee_position" = fExpand(#"Expanded employee_notes", "employee_position"),
    #"Expanded employee_pin" = fExpand(#"Expanded employee_position" , "employee_pin"),
    #"Expanded employee_salary_code" = fExpand(#"Expanded employee_pin", "employee_salary_code"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded employee_salary_code",{{"EMPLOYEE_ID", Int64.Type}, {"employee_active", type logical}, {"employee_date_join", type datetimezone}, {"employee_first_name", type text}, {"employee_last_name", type text}, {"employee_code", Int64.Type}, {"employee_username", type text}, {"employee_salary", type text}, {"employee_manager_id", Int64.Type}, {"employee_manager", type text}, {"employee_department", type text}, {"employee_use_company_schedule", Int64.Type}, {"employee_monday_schedule", Int64.Type}, {"employee_tuesday_schedule", Int64.Type}, {"employee_wednesday_schedule", Int64.Type}, {"employee_thursday_schedule", Int64.Type}, {"employee_friday_schedule", Int64.Type}, {"employee_saturday_schedule", Int64.Type}, {"employee_sunday_schedule", Int64.Type}, {"employee_language", Int64.Type}, {"employee_pin_active", Int64.Type}, {"employee_pin_active1", Int64.Type}, {"employee_price_per_hour", type text}, {"employee_cost_per_hour", type text}, {"employee_department_id", Int64.Type}, {"employee_job_id", Int64.Type}, {"employee_date_leave", type datetimezone}, {"employee_address", type text}, {"employee_zip_code", type text}, {"employee_city", type text}, {"employee_country", type text}, {"employee_phone_home", type text}, {"employee_phone_work", type text}, {"employee_phone_mobile", type text}, {"employee_email", type text}, {"employee_social_number", type text}, {"employee_abbreviated", type text}, {"employee_salary_code", type text}, {"employee_notes", type text}, {"employee_position", type text}, {"employee_pin", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"employee_salary", "employee_price_per_hour", "employee_cost_per_hour"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"employee_salary", Currency.Type}, {"employee_price_per_hour", Currency.Type}, {"employee_cost_per_hour", Currency.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetExpenseCodes = () as table =>
let
   source = fQBISRestAPI("GetExpenseCodes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded expense_code_code" = fExpand(Table6, "expense_code_code"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded expense_code_code",{{"EXPENSE_CODE_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetExpenseItems = () as table =>
let
   source = fQBISRestAPI("GetExpenseItems"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded expense_item_invoice_notes" =  fExpand(Table6, "expense_item_invoice_notes"),
    #"Expanded expense_item_notes" = fExpand(#"Expanded expense_item_invoice_notes", "expense_item_notes"),
    #"Expanded expense_item_material_invoice_notes" = fExpand(#"Expanded expense_item_notes", "expense_item_material_invoice_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded expense_item_material_invoice_notes",{{"EXPENSE_ITEM_ID", Int64.Type}, {"expense_item_expense_sheet_id", Int64.Type}, {"expense_item_date", type datetimezone}, {"expense_item_project_id", Int64.Type}, {"expense_item_expense_code_id", Int64.Type}, {"expense_item_notes", type text}, {"expense_item_number", type text}, {"expense_item_net", type text}, {"expense_item_vat", type text}, {"expense_item_gross", type text}, {"expense_item_invoice_notes", type text}, {"expense_item_name", type text}, {"expense_item_material_name", type text}, {"expense_item_material_price", type text}, {"expense_item_material_invoice_notes", type text}, {"expense_item_currency_rate_id", Int64.Type}, {"expense_item_material_invoice", Int64.Type}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"expense_item_net", "expense_item_vat", "expense_item_gross", "expense_item_material_price"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"expense_item_net", Currency.Type}, {"expense_item_vat", Currency.Type}, {"expense_item_gross", Currency.Type}, {"expense_item_material_price", Currency.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetExpenseSheets = () as table =>
let
   source = fQBISRestAPI("GetExpenseSheets"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded expense_sheet_subject" =  fExpand(Table6, "expense_sheet_subject"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded expense_sheet_subject",{{"EXPENSE_SHEET_ID", Int64.Type}, {"expense_sheet_employee_id", Int64.Type}, {"expense_sheet_date", type datetimezone}, {"expense_sheet_status", Int64.Type}, {"expense_sheet_paid", Int64.Type}, {"expense_sheet_paid_date", type datetimezone}, {"expense_sheet_subject", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetJobs = () as table =>
let
   source = fQBISRestAPI("GetJobs"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table], 
    #"Expanded job_requirements" = fExpand(Table6, "job_requirements"),
    #"Expanded job_skills" = fExpand(#"Expanded job_requirements", "job_skills"),
    #"Expanded job_notes" = fExpand(#"Expanded job_skills" , "job_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded job_notes",{{"JOB_ID", Int64.Type}, {"job_code", type text}, {"job_name", type text}, {"job_active", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
 in 
  #"Removed Columns";

GetOpportunities = () as table =>
let
   source = fQBISRestAPI("GetOpportunities"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"OPPORTUNITY_ID", Int64.Type}, {"opportunity_customer_id", Int64.Type}, {"opportunity_customer_contact_id", Int64.Type}, {"opportunity_createdby_id", Int64.Type}, {"opportunity_salesperson_id", Int64.Type}, {"opportunity_name", type text}, {"opportunity_start_date", type datetimezone}, {"opportunity_end_date", type datetimezone}, {"opportunity_percentage_value", Int64.Type}, {"opportunity_stage_name", type text}, {"opportunity_currency_name", type text}, {"opportunity_currency_symbol", type text}, {"opportunity_close_date", type datetimezone}, {"opportunity_created_date", type datetimezone}, {"opportunity_sales_project_id", Int64.Type}}),
    #"Expanded opportunity_notes" =  fExpand(#"Changed Type", "opportunity_notes"),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded opportunity_notes",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetLodgingItems = () as table =>
let
   source = fQBISRestAPI("GetLodgingItems"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table], 
    #"Expanded lodging_item_purpose" = fExpand(Table6, "lodging_item_purpose"),
    #"Expanded lodging_item_invoice_notes" = fExpand(#"Expanded lodging_item_purpose", "lodging_item_invoice_notes"),
    #"Expanded lodging_item_material_invoice_notes" = fExpand(#"Expanded lodging_item_invoice_notes", "lodging_item_material_invoice_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded lodging_item_material_invoice_notes",{{"LODGING_ITEM_ID", Int64.Type}, {"lodging_item_expense_sheet_id", Int64.Type}, {"lodging_item_from_date", type datetimezone}, {"lodging_item_to_date", type datetimezone}, {"lodging_item_lodging_type_id", Int64.Type}, {"lodging_item_project_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetLodgingTypes = () as table =>
let
   source = fQBISRestAPI("GetLodgingTypes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table], 
    #"Expanded lodging_type_gl_code" = fExpand(Table6, "lodging_type_gl_code"),
    #"Expanded lodging_type_reduction_01_gl_code" = fExpand(#"Expanded lodging_type_gl_code", "lodging_type_reduction_01_gl_code"),
    #"Expanded lodging_type_reduction_02_gl_code" = fExpand(#"Expanded lodging_type_reduction_01_gl_code", "lodging_type_reduction_02_gl_code"),
    #"Expanded lodging_type_reduction_03_gl_code" = fExpand(#"Expanded lodging_type_reduction_02_gl_code", "lodging_type_reduction_03_gl_code"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded lodging_type_reduction_03_gl_code",{{"LODGING_TYPE_ID", Int64.Type}, {"lodging_type_amount", Currency.Type}, {"lodging_type_active", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetModules = () as table =>
let
   source = fQBISRestAPI("GetModules"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table], 
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"MODULE_ID", Int64.Type}, {"module_licensed", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetPermissions = () as table =>
let
   source = fQBISRestAPI("GetPermissions"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table], 
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"permission_employee", Int64.Type}, {"permission_module", Int64.Type}, {"permission_role", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetPositionHistories = () as table =>
let
   source = fQBISRestAPI("GetPositionHistories"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"POSITION_HISTORY_ID", Int64.Type}, {"position_history_position_id", Int64.Type}, {"position_history_from_date", type datetimezone}, {"position_history_to_date", type datetimezone}, {"position_history_employee_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"}) 
in
    #"Removed Columns";

GetPositions = () as table =>
let
   source = fQBISRestAPI("GetPositions"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded position_notes" = fExpand(Table6, "position_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded position_notes",{{"POSITION_ID", Int64.Type}, {"position_job_id", Int64.Type}, {"position_department_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectActivities = () as table =>
let
   source = fQBISRestAPI("GetProjectActivities"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded project_activity_invoice_name" = fExpand(Table6, "project_activity_invoice_name"),
    #"Expanded project_activity_gl_code" = fExpand(#"Expanded project_activity_invoice_name", "project_activity_gl_code"),
    #"Expanded project_activity_category" = fExpand(#"Expanded project_activity_gl_code", "project_activity_category"),
    #"Expanded project_activity_notes" = fExpand(#"Expanded project_activity_category", "project_activity_notes"),
    #"Expanded project_activity_phase_name" = fExpand(#"Expanded project_activity_notes", "project_activity_phase_name"),
    #"Expanded project_activity_phase_code" = fExpand(#"Expanded project_activity_phase_name", "project_activity_phase_code"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded project_activity_phase_code",{{"PROJECT_ACTIVITY_ID", type number}, {"project_activity_active", type logical}, {"project_activity_complete", type logical}, {"project_activity_project_id", Int64.Type}, {"project_activity_start_date", type datetimezone}, {"project_activity_end_date", type datetimezone}, {"project_activity_chargeable", type logical}, {"project_activity_lock", type logical}, {"project_activity_warning", type logical}, {"project_activity_group_budget", type logical}, {"project_activity_phase_id", Int64.Type}, {"project_activity_chargeable_date", type datetimezone}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectMaterials = () as table =>
let
   source = fQBISRestAPI("GetProjectMaterials"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded project_material_unit" = fExpand(Table6, "project_material_unit"),
    #"Expanded project_material_invoice_notes" = fExpand(#"Expanded project_material_unit", "project_material_invoice_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded project_material_invoice_notes",{{"PROJECT_MATERIAL_ID", type number}, {"project_material_date", type datetimezone}, {"project_material_project_id", Int64.Type}, {"project_material_phase_id", Int64.Type}, {"project_material_activity_id", Int64.Type}, {"project_material_article_id", Int64.Type}, {"project_material_expense_code_id", Int64.Type}, {"project_material_supplier_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectMembers = () as table =>
let
   source = fQBISRestAPI("GetProjectMembers"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"project_member_employee_id", Int64.Type}, {"project_member_project_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectMilestones = () as table =>
let
   source = fQBISRestAPI("GetProjectMilestones"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded project_milestone_notes" = fExpand(Table6, "project_milestone_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded project_milestone_notes",{{"PROJECT_MILESTONE_ID", type number}, {"project_milestone_project_id", type number}, {"project_milestone_start_date", type datetimezone}, {"project_milestone_end_date", type datetimezone}, {"project_milestone_complete", type logical}, {"project_milestone_phase_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectPhases = () as table =>
let
   source = fQBISRestAPI("GetProjectPhases"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"PROJECT_PHASE_ID", Int64.Type}, {"project_phase_project_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjects = () as table =>
let
   source = fQBISRestAPI("GetProjects"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded project_abbreviated" =  fExpand(Table6, "project_abbreviated"),
    #"Expanded project_created_by" =  fExpand(#"Expanded project_abbreviated", "project_created_by"),
    #"Expanded project_gl_code" =  fExpand(#"Expanded project_created_by", "project_gl_code"),
    #"Expanded project_category" =  fExpand(#"Expanded project_gl_code", "project_category"),
    #"Expanded project_notes" =  fExpand(#"Expanded project_category", "project_notes"),
    #"Expanded project_order_number" =  fExpand(#"Expanded project_notes", "project_order_number"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded project_order_number" ,{{"PROJECT_ID", Int64.Type}, {"project_active", type logical}, {"project_public", type logical}, {"project_customer_id", Int64.Type}, {"project_manager_id", Int64.Type}, {"project_created_date", type datetimezone}, {"project_start_date", type datetimezone}, {"project_end_date", type datetimezone}, {"project_complete", type logical}, {"project_complete_date", type datetimezone}, {"project_chargeable", type logical}, {"project_chargeable_date", type datetimezone}, {"project_currency_rate_id", Int64.Type}, {"project_customer_contact_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetSuppliers = () as table =>
let
   source = fQBISRestAPI("GetSuppliers"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded supplier_address" =  fExpand(Table6, "supplier_address"),
    #"Expanded supplier_zip_code" =  fExpand(#"Expanded supplier_address", "supplier_zip_code"),
    #"Expanded supplier_city" =  fExpand(#"Expanded supplier_zip_code", "supplier_city"),
    #"Expanded supplier_phone" =  fExpand(#"Expanded supplier_city", "supplier_phone"),
    #"Expanded supplier_fax" =  fExpand(#"Expanded supplier_phone", "supplier_fax"),
    #"Expanded supplier_email" =  fExpand(#"Expanded supplier_fax", "supplier_email"),
    #"Expanded supplier_homepage" =  fExpand(#"Expanded supplier_email", "supplier_homepage"),
    #"Expanded supplier_country" =  fExpand(#"Expanded supplier_homepage", "supplier_country"),
    #"Expanded supplier_organization_number" =  fExpand(#"Expanded supplier_country", "supplier_organization_number"),
    #"Expanded supplier_notes" =  fExpand(#"Expanded supplier_organization_number", "supplier_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded supplier_notes",{{"SUPPLIER_ID", Int64.Type}, {"supplier_active", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetSalaryActivities = () as table =>
    let
   source = fQBISRestAPI("GetSalaryActivities"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded salary_activity_gl_code" =  fExpand(Table6, "salary_activity_gl_code"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded salary_activity_gl_code",{{"SALARY_ACTIVITY_ID", Int64.Type}, {"salary_activity_active", type logical}, {"salary_activity_default", type logical}, {"salary_activity_allow_negative", type logical}, {"salary_activity_allow_positive", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
#"Removed Columns";

GetSalaryTime = () as table =>
let
   source = fQBISRestAPI("GetSalaryTime"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded salary_time_notes" = fExpand(Table6, "salary_time_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded salary_time_notes",{{"SALARY_TIME_ID", Int64.Type}, {"salary_time_employee_id", Int64.Type}, {"salary_time_activity_id", Int64.Type}, {"salary_time_date", type datetimezone}, {"salary_time_locked", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleAgreementActivities = () as table =>
    let
   source = fQBISRestAPI("GetScheduleAgreementActivities"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"schedule_agreement_activity_agreement_id", Int64.Type}, {"schedule_agreement_activity_salary_activity_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleAgreementSpecialDays = () as table =>
    let
   source = fQBISRestAPI("GetScheduleAgreementSpecialDays"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"schedule_agreement_special_days_agreement_id", Int64.Type}, {"schedule_agreement_special_days_date", type datetimezone}, {"schedule_agreement_special_days_salary_activity_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleAgreements = () as table =>
    let
   source = fQBISRestAPI("GetScheduleAgreements"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded schedule_agreement_notes" = fExpand(Table6, "schedule_agreement_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded schedule_agreement_notes",{{"SCHEDULE_AGREEMENT_ID", Int64.Type}, {"schedule_agreement_active", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleBreaks = () as table =>
    let
   source = fQBISRestAPI("GetScheduleBreaks"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"schedule_break_week_id", Int64.Type}, {"schedule_break_interval_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleDays = () as table =>
    let
   source = fQBISRestAPI("GetScheduleDays"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"SCHEDULE_DAY_ID", Int64.Type}, {"schedule_day_arrive_time", type time}, {"schedule_day_leave_time", type time}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleEmployees = () as table =>
    let
   source = fQBISRestAPI("GetScheduleEmployees"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"schedule_employee_id", Int64.Type}, {"schedule_employee_schedule_id", Int64.Type}, {"schedule_employee_from_date", type datetimezone}, {"schedule_employee_to_date", type datetimezone}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleIntervals = () as table =>
    let
   source = fQBISRestAPI("GetScheduleIntervals"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"SCHEDULE_INTERVAL_ID", Int64.Type}, {"schedule_interval_from_time", type time}, {"schedule_interval_to_time", type time}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleSalaryActivities = () as table =>
    let
   source = fQBISRestAPI("GetScheduleSalaryActivities"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"SCHEDULE_SALARY_ACTIVITY_ID", Int64.Type}, {"schedule_salary_activity_activity_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleSalaryActivityDays = () as table =>
    let
   source = fQBISRestAPI("GetScheduleSalaryActivityDays"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"schedule_salary_activity_day_activity_id", Int64.Type}, {"schedule_salary_activity_day_interval_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleSalaryActivityWeeks = () as table =>
    let
   source = fQBISRestAPI("GetScheduleSalaryActivityWeeks"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"schedule_salary_activity_week_id", Int64.Type}, {"schedule_salary_activity_week_activity_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleTimes = () as table =>
let
   source = fQBISRestAPI("GetScheduleTimes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"schedule_time_week_id", Int64.Type}, {"schedule_time_day_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleVersions = () as table =>
    let
   source = fQBISRestAPI("GetScheduleVersions"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"SCHEDULE_VERSION_ID", Int64.Type}, {"schedule_version_schedule_id", Int64.Type}, {"schedule_version_from_date", type datetimezone}, {"schedule_version_to_date", type datetimezone}, {"schedule_version_recurring", type logical}, {"schedule_version_changed_by_id", Int64.Type}, {"schedule_version_changed_date", type datetimezone}, {"schedule_version_lunches_specified", type logical}, {"schedule_version_override_public_holidays", type logical}, {"schedule_version_agreement_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetScheduleWeeks = () as table =>
    let
   source = fQBISRestAPI("GetScheduleWeeks"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"SCHEDULE_WEEK_ID", Int64.Type}, {"schedule_week_version_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetSchedules = () as table =>
    let
   source = fQBISRestAPI("GetSchedules"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded schedule_notes" = fExpand(Table6, "schedule_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded schedule_notes",{{"SCHEDULE_ID", Int64.Type}, {"schedule_active", type logical}, {"schedule_deleted", type logical}, {"schedule_prefill_time_active", type logical}, {"schedule_from_date", type datetimezone}, {"schedule_to_date", type datetimezone}, {"schedule_created_by_id", Int64.Type}, {"schedule_created_date", type datetimezone}, {"schedule_changed_by_id", Int64.Type}, {"schedule_changed_date", type datetimezone}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetTasks = () as table =>
let
   source = fQBISRestAPI("GetTasks"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded task_notes" =  fExpand(Table6, "task_notes"),
    #"Expanded task_company_category" =  fExpand( #"Expanded task_notes", "task_company_category"),
    #"Expanded task_opportunity_name" =  fExpand( #"Expanded task_company_category", "task_opportunity_name"),
    #"Expanded task_opportunity_stage_name" =  fExpand(#"Expanded task_opportunity_name" , "task_opportunity_stage_name"),
    #"Expanded task_reminder_time" =  fExpand(#"Expanded task_opportunity_stage_name", "task_reminder_time"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded task_reminder_time",{{"TASK_ID", Int64.Type}, {"task_date", type datetimezone}, {"task_start_time", type time}, {"task_end_time", type time}, {"task_complete", type logical}, {"task_show_in_calendar", type logical}, {"task_reminder_date", type datetimezone}, {"task_reminder_time", type time}, {"task_company_id", Int64.Type}, {"task_assigned_to_id", Int64.Type}, {"task_opportunity_id", Int64.Type}, {"task_created_by_id", Int64.Type}, {"task_contact_way_id", Int64.Type}, {"task_contact_person_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetTravelTypes = () as table =>
let
   source = fQBISRestAPI("GetTravelTypes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"TRAVEL_TYPE_ID", Int64.Type}, {"travel_type_active", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetTravelItems = () as table =>
let
   source = fQBISRestAPI("GetTravelItems"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Expanded travel_item_purpose" = fExpand(Table6, "travel_item_purpose"),
    #"Expanded travel_item_invoice_notes" = fExpand(#"Expanded travel_item_purpose", "travel_item_invoice_notes"),
    #"Expanded travel_item_material_invoice_notes" = fExpand(#"Expanded travel_item_invoice_notes", "travel_item_material_invoice_notes"),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded travel_item_material_invoice_notes",{{"TRAVEL_ITEM_ID", Int64.Type}, {"travel_item_expense_sheet_id", Int64.Type}, {"travel_item_from_date", type datetimezone}, {"travel_item_to_date", type datetimezone}, {"travel_item_project_id", Int64.Type}, {"travel_item_travel_type_id", Int64.Type}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"travel_item_travel_type_amount"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"travel_item_travel_type_amount", type number}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type1",".",",",Replacer.ReplaceText,{"travel_item_material_price"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Replaced Value1",{{"travel_item_material_price", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type2",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetUnitTypes = () as table =>
let
   source = fQBISRestAPI("GetUnitTypes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"UNIT_TYPE_ID", Int64.Type}, {"unit_type_parent_id", Int64.Type}, {"unit_type_system", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetWorkingTime = () as table =>
let
   source = fQBISRestAPI("GetWorkingTime"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"WORKING_TIME_ID", Int64.Type}, {"working_time_arrived", type datetimezone}, {"working_time_left", type datetimezone}, {"working_time_employee_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

//Mobile
GeTeams = () as table =>
let
    source = fQBISMobileRestAPI("GeTeams"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"team_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GeTeamsMembers = () as table =>
let
    source = fQBISMobileRestAPI("GeTeamsMembers"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"teammember_member_id", Int64.Type}, {"teammember_team_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetCRMCategories = () as table =>
let
    source = fQBISMobileRestAPI("GetCRMCategories"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"CATEGORY_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
    //#"Expanded lodging_item_purpose" = fExpand(Table6, "lodging_item_purpose")
in
    #"Removed Columns";

GetContactTypes = () as table =>
let
    source = fQBISMobileRestAPI("GetContactTypes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"CONTACT_TYPE_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetPipelinePercentages = () as table =>
let
    source = fQBISMobileRestAPI("GetPipelinePercentages"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"PIPELINE_PERCENTAGE_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetPipelineStages = () as table =>
let
    source = fQBISMobileRestAPI("GetPipelineStages"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"PIPELINE_STAGE_ID", Int64.Type}, {"pipeline_stage_isoffer", type logical}, {"pipeline_stage_isclose", type logical}, {"pipeline_stage_islock", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetPriorityLevels = () as table =>
let
    source = fQBISMobileRestAPI("GetPriorityLevels"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"priority_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetServiceDeskRequestCategories = () as table =>
let
    source = fQBISMobileRestAPI("GetServiceDeskRequestCategories"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"SDRC_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetServiceDeskSchedue = () as table =>
let
    source = fQBISMobileRestAPI("GetServiceDeskSchedue"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"Schedule_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetServiceDeskSettings = () as table =>
let
    source = fQBISMobileRestAPI("GetServiceDeskSettings"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"sds_materials", type logical}, {"sds_applycustomerfilterstoadministrator", type logical}, {"sds_usecustomerteamcategoryfilter", type logical}, {"sds_usecustomerteamcontractfilter", type logical}, {"sds_canassignuser", type logical}, {"isProject", type logical}, {"isCRM", type logical}, {"isServiceDesk", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetStatusTypes = () as table =>
let
    source = fQBISMobileRestAPI("GetStatusTypes"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"statustype_id", Int64.Type}, {"statustype_systemtype", type logical}, {"statustype_sendmessage", type logical}, {"statustype_remindremained", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetTeamsCategories = () as table =>
    let
    source = fQBISMobileRestAPI("GetTeamsCategories"),
    Table = source{0}[Table],
    Table1 = Table{0}[Table],
    Table2 = Table1{1}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{0}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table6,{{"teamcategory_team_id", Int64.Type}, {"teamcategory_category_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

//
// Helper Functions
//

Value.IfNull = (a, b) => if a <> null then a else b;

GetScopeString = (scopes as list, optional scopePrefix as text) as text =>
	let
		prefix = Value.IfNull(scopePrefix, ""),
		addPrefix = List.Transform(scopes, each prefix & _),
		asText = Text.Combine(addPrefix, " ")
	in
		asText;

Table.ToNavigationTable = (
	table as table,
	keyColumns as list,
	nameColumn as text,
	dataColumn as text,
	itemKindColumn as text,
	itemNameColumn as text,
	isLeafColumn as text
) as table =>
	let
		tableType = Value.Type(table),
		newTableType = Type.AddTableKey(tableType, keyColumns, true) meta 
		[
			NavigationTable.NameColumn = nameColumn, 
			NavigationTable.DataColumn = dataColumn,
			NavigationTable.ItemKindColumn = itemKindColumn, 
			Preview.DelayColumn = itemNameColumn, 
			NavigationTable.IsLeafColumn = isLeafColumn
		],
		navigationTable = Value.ReplaceType(table, newTableType)
	in
		navigationTable;

fQBISRestAPI = (
    QBisTable as text
) as table =>
   let
      body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
                            <soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
                               <soap12:Body>
                                  <" & QBisTable & " xmlns='https://bin.qbis.se/WS/QBIS_DS'>
                                     <Company>" & Extension.CurrentCredential()[Username] & "</Company>
                                     <Password>" & Extension.CurrentCredential()[Password] & "</Password>
                                  </" & QBisTable & ">
                               </soap12:Body>
                            </soap12:Envelope>"),
      actualUrl = QbisActualURL,
      options = [
         Headers = [
                    #"Content-type"="application/soap+xml",
                    #"Charset"="UTF-8"
         ],
         Content = body
      ],
      result = Xml.Tables(Web.Contents(actualUrl, options)),
      Table = result{0}[Table],
      Table1 = Table{0}[Table],
      CheckEmpty1 = if Table.IsEmpty(Table1) then null else Table1
    in
       CheckEmpty1;

fQBISMobileRestAPI = (
    QBisTable as text
) as table =>
   let
      body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
                            <soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
                               <soap12:Body>
                                  <" & QBisTable & " xmlns='https://bin.qbis.se/WS/QBIS_MOBILEAPP'>
                                     <Company>" & Extension.CurrentCredential()[Username] & "</Company>
                                     <Password>" & Extension.CurrentCredential()[Password] & "</Password>
                                  </" & QBisTable & ">
                               </soap12:Body>
                            </soap12:Envelope>"),
      actualUrl = QbisMobileActualURL,
      options = [
         Headers = [
                    #"Content-type"="application/soap+xml",
                    #"Charset"="UTF-8"
         ],
         Content = body
      ],
      result = Xml.Tables(Web.Contents(actualUrl, options)),
      Table = result{0}[Table],
      Table1 = Table{0}[Table],
      CheckEmpty1 = if Table.IsEmpty(Table1) then null else Table1
    in
       CheckEmpty1;

fExpand = (
    Source as table, 
    Column_name as any
) as table =>
    let
       #"Lists from Column" = Table.TransformColumns(Source,{{Column_name, each if _ is table then Table.ToList(_) else {_}}}),
       #"Expanded Column" = Table.ExpandListColumn(#"Lists from Column" , Column_name)
    in
    #"Expanded Column";