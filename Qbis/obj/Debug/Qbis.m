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
//             GetPositionHistories 
//             GetPositions 
            { Extension.LoadString("GetProjectActivities"), GetProjectActivities(), "Table", "Table", true },
            { Extension.LoadString("GetProjectMaterials"), GetProjectMaterials(), "Table", "Table", true },
            { Extension.LoadString("GetProjectMembers"), GetProjectMembers(), "Table", "Table", true },
            { Extension.LoadString("GetProjectMilestones"), GetProjectMilestones(), "Table", "Table", true },            
//             GetProjectPhases 
            { Extension.LoadString("GetProjects"), GetProjects(), "Table", "Table", true },          
//             GetResourceAvailability 
//             GetSalaryActivities 
//             GetSalaryTime 
//             GetScheduleAgreementActivities 
//             GetScheduleAgreementSpecialDays
//             GetScheduleAgreements 
//             GetScheduleBreaks 
//             GetScheduleDays 
//             GetScheduleDepartments 
//             GetScheduleEmployees 
//             GetScheduleIntervals 
//             GetSchedulePositions 
//             GetScheduleSalaryActivities 
//             GetScheduleSalaryActivityDays 
//             GetScheduleSalaryActivityWeeks 
//             GetScheduleTimes 
//             GetScheduleVersions 
//             GetScheduleWeeks 
//             GetSchedules 
            { Extension.LoadString("GetSuppliers"), GetSuppliers(), "Table", "Table", true },          
            { Extension.LoadString("GetTasks"), GetTasks(), "Table", "Table", true },          
            { Extension.LoadString("GetTravelItems"), GetTravelItems(), "Table", "Table", true },          
            { Extension.LoadString("GetTravelTypes"), GetTravelTypes(), "Table", "Table", true },          
            { Extension.LoadString("GetUnitTypes"), GetUnitTypes(), "Table", "Table", true },          
            { Extension.LoadString("GetWorkingTime"), GetWorkingTime(), "Table", "Table", true }         

            //Mobile
//             GeTeams 
//             GeTeamsMembers 
//             GetCRMCategories 
//             GetCRMCompanies 
//             GetCRMContacts 
//             GetCompanyCRMCategories 
//             GetCompanySDCategories 
//             GetContactTypes
//             GetPipelinePercentages 
//             GetPipelineStages 
//             GetPriorityLevels 
//             GetProducts 
//             GetSalesProjects 
//             GetServiceDeskContacts 
//             GetServiceDeskContracts 
//             GetServiceDeskRequestCategories 
//             GetServiceDeskRequests 
//             GetServiceDeskSchedue 
//             GetServiceDeskSettings 
//             GetServiceDeskStatusLog 
//             GetServiceRequestsActions 
//             GetServiceRequestsMaterial 
//             GetServicedeskCompanies 
//             GetStatusTypes 
//             GetTeamsCategories 
 		}),
		navTable = Table.ToNavigationTable(source, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
	in
		navTable;

GetActivityMembers = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetActivityMembers xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetActivityMembers>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"activity_member_activity_id", Int64.Type}, {"activity_member_employee_id", Int64.Type}, {"activity_member_allocated_minutes", Int64.Type}, {"activity_member_completed_by_member", type logical}, {"activity_member_completed_by_manager", type logical}, {"activity_member_cost_per_hour", type text}, {"activity_member_price_per_hour", type text}}),
	#"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"activity_member_cost_per_hour", "activity_member_price_per_hour"}),
	#"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"activity_member_cost_per_hour", Currency.Type}, {"activity_member_price_per_hour", Currency.Type}})
in
	#"Changed Type1";

GetActivityTime = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetActivityTime xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetActivityTime>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Expanded activity_time_notes_internal" = Table.ExpandTableColumn(Table8, "activity_time_notes_internal", {"Element:Text"}, {"activity_time_notes_internal"}),
	#"Expanded activity_time_notes_external" = Table.ExpandTableColumn(#"Expanded activity_time_notes_internal", "activity_time_notes_external", {"Element:Text"}, {"activity_time_notes_external"}),
	#"Changed Type" = Table.TransformColumnTypes(#"Expanded activity_time_notes_external",{{"ACTIVITY_TIME_ID", Int64.Type}, {"activity_time_activity_id", Int64.Type}, {"activity_time_employee_id", Int64.Type}, {"activity_time_minutes", Int64.Type}, {"activity_time_notes_external", type text}, {"activity_time_notes_internal", type text}, {"activity_time_factor", type text}, {"activity_time_date", type datetimezone}}),
	#"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"activity_time_factor"}),
	#"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"activity_time_factor", type number}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetArticles = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetArticles xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetArticles>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Expanded article_unit" = Table.ExpandTableColumn(Table8, "article_unit", {"Element:Text"}, {"article_unit"}),
	#"Expanded article_notes" = Table.ExpandTableColumn(#"Expanded article_unit", "article_notes", {"Element:Text"}, {"article_notes"}),
	#"Expanded article_gl_code" = Table.ExpandTableColumn(#"Expanded article_notes", "article_gl_code", {"Element:Text"}, {"article_gl_code"}),
	#"Changed Type" = Table.TransformColumnTypes(#"Expanded article_gl_code",{{"ARTICLE_ID", Int64.Type}, {"article_code", type text}, {"article_name", type text}, {"article_unit", type text}, {"article_cost", type text}, {"article_price", type text}, {"article_supplier_id", Int64.Type}, {"article_notes", type text}}),
	#"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"article_cost", "article_price"}),
	#"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"article_cost", Currency.Type}, {"article_price", Currency.Type}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetCRMTypes = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetCRMTypes xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetCRMTypes>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"CRM_TYPE_ID", Int64.Type}, {"crm_type_name", type text}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetCompanyInformation = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetCompanyInformation xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetCompanyInformation>
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
	  Table2 = Table1{0}[Table],
	  Table3 = Table2{0}[Table],
	  Table4 = Table3{1}[Table],
	  Table5 = Table4{0}[Table],
	  Table6 = Table5{0}[Table],
	  Table7 = Table6{0}[Table],
	  Table8 = Table7{0}[Table],
	  #"Changed Type" = Table.TransformColumnTypes(Table8,{{"company_name", type text}, {"company_organisation_number", type text}, {"company_minimum_lunch", Int64.Type}, {"company_maximum_lunch", Int64.Type}, {"company_use_lunch_rule", type logical}, {"company_monday_schedule", Int64.Type}, {"company_tuesday_schedule", Int64.Type}, {"company_wednesday_schedule", Int64.Type}, {"company_thursday_schedule", Int64.Type}, {"company_friday_schedule", Int64.Type}, {"company_saturday_schedule", Int64.Type}, {"company_sunday_schedule", Int64.Type}, {"company_allow_timeclock", type logical}, {"company_enforce_lunch_timeclock", type logical}, {"company_enforce_lunch_min_timeclock", Int64.Type}, {"company_enforce_lunch_from_timeclock", type datetimezone}, {"company_show_inout", type logical}}),
	  #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
  in
	  #"Removed Columns";

GetContactWays = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetContactWays xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetContactWays>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"CONTACT_WAY_ID", Int64.Type}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetCurrencies = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetCurrencies xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetCurrencies>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"CURRENCY_ID", Int64.Type}, {"currency_name", type text}, {"currency_region", type text}, {"currency_default", type logical}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetCurrencyRates = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetCurrencyRates xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetCurrencyRates>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"CURRENCY_RATE_ID", Int64.Type}, {"currency_rate_currency_id", Int64.Type}, {"currency_rate_rate", type text}, {"currency_rate_date", type datetimezone}}),
	#"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"currency_rate_rate"}),
	#"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"currency_rate_rate", Currency.Type}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetCustomerContacts = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetCustomerContacts xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetCustomerContacts>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Expanded customer_contact_title" = Table.ExpandTableColumn(Table8, "customer_contact_title", {"Element:Text"}, {"customer_contact_title.Element:Text"}),
	#"Expanded customer_contact_phone_work" = Table.ExpandTableColumn(#"Expanded customer_contact_title", "customer_contact_phone_work", {"Element:Text"}, {"customer_contact_phone_work.Element:Text"}),
	#"Expanded customer_contact_phone_mobile" = Table.ExpandTableColumn(#"Expanded customer_contact_phone_work", "customer_contact_phone_mobile", {"Element:Text"}, {"customer_contact_phone_mobile.Element:Text"}),
	#"Expanded customer_contact_email" = Table.ExpandTableColumn(#"Expanded customer_contact_phone_mobile", "customer_contact_email", {"Element:Text"}, {"customer_contact_email.Element:Text"}),
	#"Expanded customer_contact_social_number" = Table.ExpandTableColumn(#"Expanded customer_contact_email", "customer_contact_social_number", {"Element:Text"}, {"customer_contact_social_number.Element:Text"}),
	#"Expanded customer_contact_type" = Table.ExpandTableColumn(#"Expanded customer_contact_social_number", "customer_contact_type", {"Element:Text"}, {"customer_contact_type.Element:Text"}),
	#"Renamed Columns" = Table.RenameColumns(#"Expanded customer_contact_type",{{"customer_contact_type.Element:Text", "customer_contact_type"}, {"customer_contact_social_number.Element:Text", "customer_contact_social_number"}, {"customer_contact_email.Element:Text", "customer_contact_email"}, {"customer_contact_phone_mobile.Element:Text", "customer_contact_phone_mobile"}, {"customer_contact_phone_work.Element:Text", "customer_contact_phone_work"}, {"customer_contact_title.Element:Text", "customer_contact_title"}}),
	#"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"CUSTOMER_CONTACT_ID", Int64.Type}, {"customer_contact_customer_id", Int64.Type}, {"customer_contact_social_number", type text}, {"customer_contact_email", type text}, {"customer_contact_phone_mobile", type text}, {"customer_contact_phone_work", type text}, {"customer_contact_last_name", type text}, {"customer_contact_first_name", type text}, {"customer_contact_primary", type logical}, {"customer_contact_newsletter", type logical}, {"customer_contact_type", type text}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetCustomers = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetCustomers xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetCustomers>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"CUSTOMER_ID", Int64.Type}, {"customer_active", type logical}, {"customer_name", type text}, {"customer_code", type text}, {"customer_abbreviated", type text}}),
	#"Expanded customer_address" = Table.ExpandTableColumn(#"Changed Type", "customer_address", {"Element:Text"}, {"customer_address"}),
	#"Expanded customer_zip_code" = Table.ExpandTableColumn(#"Expanded customer_address", "customer_zip_code", {"Element:Text"}, {"customer_zip_code"}),
	#"Expanded customer_city" = Table.ExpandTableColumn(#"Expanded customer_zip_code", "customer_city", {"Element:Text"}, {"customer_city"}),
	#"Expanded customer_country" = Table.ExpandTableColumn(#"Expanded customer_city", "customer_country", {"Element:Text"}, {"customer_country"}),
	#"Expanded customer_phone" = Table.ExpandTableColumn(#"Expanded customer_country", "customer_phone", {"Element:Text"}, {"customer_phone"}),
	#"Expanded customer_fax" = Table.ExpandTableColumn(#"Expanded customer_phone", "customer_fax", {"Element:Text"}, {"customer_fax"}),
	#"Expanded customer_email" = Table.ExpandTableColumn(#"Expanded customer_fax", "customer_email", {"Element:Text"}, {"customer_email"}),
	#"Expanded customer_homepage" = Table.ExpandTableColumn(#"Expanded customer_email", "customer_homepage", {"Element:Text"}, {"customer_homepage"}),
	#"Expanded customer_organization_number" = Table.ExpandTableColumn(#"Expanded customer_homepage", "customer_organization_number", {"Element:Text"}, {"customer_organization_number"}),
	#"Expanded customer_notes" = Table.ExpandTableColumn(#"Expanded customer_organization_number", "customer_notes", {"Element:Text"}, {"customer_notes"}),
	#"Removed Columns" = Table.RemoveColumns(#"Expanded customer_notes",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetDepartments = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetDepartments xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetDepartments>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"DEPARTMENT_ID", Int64.Type}, {"department_code", type text}, {"department_name", type text}, {"department_manager_id", Int64.Type}, {"department_manager_first_name", type text}, {"department_manager_last_name", type text}, {"department_active", type logical}, {"department_unit_type_id", Int64.Type}, {"department_parent_id", Int64.Type}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetEmployees = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetEmployees xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetEmployees>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Expanded employee_address" = Table.ExpandTableColumn(Table8, "employee_address", {"Element:Text"}, {"employee_address"}),
	#"Expanded employee_zip_code" = Table.ExpandTableColumn(#"Expanded employee_address", "employee_zip_code", {"Element:Text"}, {"employee_zip_code"}),
	#"Expanded employee_city" = Table.ExpandTableColumn(#"Expanded employee_zip_code", "employee_city", {"Element:Text"}, {"employee_city"}),
	#"Expanded employee_country" = Table.ExpandTableColumn(#"Expanded employee_city", "employee_country", {"Element:Text"}, {"employee_country"}),
	#"Expanded employee_phone_home" = Table.ExpandTableColumn(#"Expanded employee_country", "employee_phone_home", {"Element:Text"}, {"employee_phone_home"}),
	#"Expanded employee_phone_work" = Table.ExpandTableColumn(#"Expanded employee_phone_home", "employee_phone_work", {"Element:Text"}, {"employee_phone_work"}),
	#"Expanded employee_phone_mobile" = Table.ExpandTableColumn(#"Expanded employee_phone_work", "employee_phone_mobile", {"Element:Text"}, {"employee_phone_mobile"}),
	#"Expanded employee_email" = Table.ExpandTableColumn(#"Expanded employee_phone_mobile", "employee_email", {"Element:Text"}, {"employee_email"}),
	#"Expanded employee_social_number" = Table.ExpandTableColumn(#"Expanded employee_email", "employee_social_number", {"Element:Text"}, {"employee_social_number"}),
	#"Expanded employee_abbreviated" = Table.ExpandTableColumn(#"Expanded employee_social_number", "employee_abbreviated", {"Element:Text"}, {"employee_abbreviated"}),
	#"Expanded employee_notes" = Table.ExpandTableColumn(#"Expanded employee_abbreviated", "employee_notes", {"Element:Text"}, {"employee_notes"}),
	#"Expanded employee_position" = Table.ExpandTableColumn(#"Expanded employee_notes", "employee_position", {"Element:Text"}, {"employee_position"}),
	#"Expanded employee_pin" = Table.ExpandTableColumn(#"Expanded employee_position", "employee_pin", {"Element:Text"}, {"employee_pin"}),
	#"Expanded employee_salary_code" = Table.ExpandTableColumn(#"Expanded employee_pin", "employee_salary_code", {"Element:Text"}, {"employee_salary_code"}),
	#"Changed Type" = Table.TransformColumnTypes(#"Expanded employee_salary_code",{{"EMPLOYEE_ID", Int64.Type}, {"employee_active", type logical}, {"employee_date_join", type datetimezone}, {"employee_first_name", type text}, {"employee_last_name", type text}, {"employee_code", Int64.Type}, {"employee_username", type text}, {"employee_salary", type text}, {"employee_manager_id", Int64.Type}, {"employee_manager", type text}, {"employee_department", type text}, {"employee_use_company_schedule", Int64.Type}, {"employee_monday_schedule", Int64.Type}, {"employee_tuesday_schedule", Int64.Type}, {"employee_wednesday_schedule", Int64.Type}, {"employee_thursday_schedule", Int64.Type}, {"employee_friday_schedule", Int64.Type}, {"employee_saturday_schedule", Int64.Type}, {"employee_sunday_schedule", Int64.Type}, {"employee_language", Int64.Type}, {"employee_pin_active", Int64.Type}, {"employee_pin_active1", Int64.Type}, {"employee_price_per_hour", type text}, {"employee_cost_per_hour", type text}, {"employee_department_id", Int64.Type}, {"employee_job_id", Int64.Type}, {"employee_date_leave", type datetimezone}, {"employee_address", type text}, {"employee_zip_code", type text}, {"employee_city", type text}, {"employee_country", type text}, {"employee_phone_home", type text}, {"employee_phone_work", type text}, {"employee_phone_mobile", type text}, {"employee_email", type text}, {"employee_social_number", type text}, {"employee_abbreviated", type text}, {"employee_salary_code", type text}, {"employee_notes", type text}, {"employee_position", type text}, {"employee_pin", type text}}),
	#"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"employee_salary", "employee_price_per_hour", "employee_cost_per_hour"}),
	#"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"employee_salary", Currency.Type}, {"employee_price_per_hour", Currency.Type}, {"employee_cost_per_hour", Currency.Type}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetExpenseCodes = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetExpenseCodes xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetExpenseCodes>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Expanded expense_code_gl_code" = Table.ExpandTableColumn(Table8, "expense_code_gl_code", {"Element:Text"}, {"expense_code_gl_code.Element:Text"}),
	#"Renamed Columns" = Table.RenameColumns(#"Expanded expense_code_gl_code",{{"expense_code_gl_code.Element:Text", "expense_code_gl_code"}}),
	#"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"EXPENSE_CODE_ID", Int64.Type}, {"expense_code_name", type text}, {"expense_code_code", type text}, {"expense_code_gl_code", Int64.Type}, {"expense_code_vat_rate", type number}, {"expense_code_active", Int64.Type}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetExpenseItems = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetExpenseItems xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetExpenseItems>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Expanded expense_item_invoice_notes" = Table.ExpandTableColumn(Table8, "expense_item_invoice_notes", {"Element:Text"}, {"expense_item_invoice_notes"}),
	#"Expanded expense_item_notes" = Table.ExpandTableColumn(#"Expanded expense_item_invoice_notes", "expense_item_notes", {"Element:Text"}, {"expense_item_notes"}),
	#"Expanded expense_item_material_invoice_notes" = Table.ExpandTableColumn(#"Expanded expense_item_notes", "expense_item_material_invoice_notes", {"Element:Text"}, {"expense_item_material_invoice_notes"}),
	#"Changed Type" = Table.TransformColumnTypes(#"Expanded expense_item_material_invoice_notes",{{"EXPENSE_ITEM_ID", Int64.Type}, {"expense_item_expense_sheet_id", Int64.Type}, {"expense_item_date", type datetimezone}, {"expense_item_project_id", Int64.Type}, {"expense_item_expense_code_id", Int64.Type}, {"expense_item_notes", type text}, {"expense_item_number", type text}, {"expense_item_net", type text}, {"expense_item_vat", type text}, {"expense_item_gross", type text}, {"expense_item_invoice_notes", type text}, {"expense_item_name", type text}, {"expense_item_material_name", type text}, {"expense_item_material_price", type text}, {"expense_item_material_invoice_notes", type text}, {"expense_item_currency_rate_id", Int64.Type}, {"expense_item_material_invoice", Int64.Type}}),
	#"Replaced Value" = Table.ReplaceValue(#"Changed Type",".",",",Replacer.ReplaceText,{"expense_item_net", "expense_item_vat", "expense_item_gross", "expense_item_material_price"}),
	#"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"expense_item_net", Currency.Type}, {"expense_item_vat", Currency.Type}, {"expense_item_gross", Currency.Type}, {"expense_item_material_price", Currency.Type}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetExpenseSheets = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetExpenseSheets xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetExpenseSheets>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Expanded expense_sheet_subject" = Table.ExpandTableColumn(Table8, "expense_sheet_subject", {"Element:Text"}, {"expense_sheet_subject"}),
	#"Changed Type" = Table.TransformColumnTypes(#"Expanded expense_sheet_subject",{{"EXPENSE_SHEET_ID", Int64.Type}, {"expense_sheet_employee_id", Int64.Type}, {"expense_sheet_date", type datetimezone}, {"expense_sheet_status", Int64.Type}, {"expense_sheet_paid", Int64.Type}, {"expense_sheet_paid_date", type datetimezone}, {"expense_sheet_subject", type text}}),
	#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetJobs = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetJobs xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetJobs>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded job_requirements" = Table.ExpandTableColumn(#"Table8", "job_requirements", {"Element:Text"}, {"job_requirements"}),
        #"Expanded job_skills" = Table.ExpandTableColumn(#"Expanded job_requirements", "job_skills", {"Element:Text"}, {"job_skills"}),
            #"Expanded job_notes" = Table.ExpandTableColumn(#"Expanded job_skills", "job_notes", {"Element:Text"}, {"job_notes"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded job_notes",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetLodgingItems = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetLodgingItems xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetLodgingItems>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "lodging_item_purpose", {"Element:Text"}, {"lodging_item_purpose"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "lodging_item_invoice_notes", {"Element:Text"}, {"lodging_item_invoice_notes"}),
    #"Expanded {0}2" = Table.ExpandTableColumn(#"Expanded {0}1", "lodging_item_city", {"Element:Text"}, {"lodging_item_city"}),
    #"Expanded {0}3" = Table.ExpandTableColumn(#"Expanded {0}2", "lodging_item_material_invoice_notes", {"Element:Text"}, {"lodging_item_material_invoice_notes"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}3",{{"LODGING_ITEM_ID", Int64.Type}, {"lodging_item_expense_sheet_id", Int64.Type}, {"lodging_item_from_date", type datetimezone}, {"lodging_item_to_date", type datetimezone}, {"lodging_item_lodging_type_id", Int64.Type}, {"lodging_item_project_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetLodgingTypes = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetLodgingTypes xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetLodgingTypes>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "lodging_type_gl_code", {"Element:Text"}, {"lodging_type_gl_code.Element:Text"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "lodging_type_reduction_01_gl_code", {"Element:Text"}, {"lodging_type_reduction_01_gl_code"}),
    #"Expanded {0}2" = Table.ExpandTableColumn(#"Expanded {0}1", "lodging_type_reduction_02_gl_code", {"Element:Text"}, {"lodging_type_reduction_02_gl_code"}),
    #"Expanded {0}3" = Table.ExpandTableColumn(#"Expanded {0}2", "lodging_type_reduction_03_gl_code", {"Element:Text"}, {"lodging_type_reduction_03_gl_code"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}3",{{"LODGING_TYPE_ID", Int64.Type}, {"lodging_type_amount", Currency.Type}, {"lodging_type_active", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetModules = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetModules xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetModules>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table8,{{"MODULE_ID", Int64.Type}, {"module_licensed", type logical}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetOpportunities = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetOpportunities xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetOpportunities>
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
	Table2 = Table1{0}[Table],
	Table3 = Table2{0}[Table],
	Table4 = Table3{1}[Table],
	Table5 = Table4{0}[Table],
	Table6 = Table5{0}[Table],
	Table7 = Table6{0}[Table],
	Table8 = Table7{0}[Table],
	#"Changed Type" = Table.TransformColumnTypes(Table8,{{"OPPORTUNITY_ID", Int64.Type}, {"opportunity_customer_id", Int64.Type}, {"opportunity_customer_contact_id", Int64.Type}, {"opportunity_createdby_id", Int64.Type}, {"opportunity_salesperson_id", Int64.Type}, {"opportunity_name", type text}, {"opportunity_start_date", type datetimezone}, {"opportunity_end_date", type datetimezone}, {"opportunity_percentage_value", Int64.Type}, {"opportunity_stage_name", type text}, {"opportunity_currency_name", type text}, {"opportunity_currency_symbol", type text}, {"opportunity_close_date", type datetimezone}, {"opportunity_created_date", type datetimezone}, {"opportunity_sales_project_id", Int64.Type}}),
	#"Expanded opportunity_notes" = Table.ExpandTableColumn(#"Changed Type", "opportunity_notes", {"Element:Text"}, {"opportunity_notes.Element:Text"}),
	#"Removed Columns" = Table.RemoveColumns(#"Expanded opportunity_notes",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
	#"Removed Columns";

GetPermissions = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetPermissions xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetPermissions>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table8,{{"permission_employee", Int64.Type}, {"permission_module", Int64.Type}, {"permission_role", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectActivities = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetProjectActivities xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetProjectActivities>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "project_activity_name", {"Element:Text"}, {"project_activity_name"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "project_activity_code", {"Element:Text"}, {"project_activity_code"}),
    #"Expanded {0}2" = Table.ExpandTableColumn(#"Expanded {0}1", "project_activity_invoice_name", {"Element:Text"}, {"project_activity_invoice_name"}),
    #"Expanded {0}3" = Table.ExpandTableColumn(#"Expanded {0}2", "project_activity_gl_code", {"Element:Text"}, {"project_activity_gl_code"}),
    #"Expanded {0}4" = Table.ExpandTableColumn(#"Expanded {0}3", "project_activity_category", {"Element:Text"}, {"project_activity_category"}),
    #"Expanded {0}5" = Table.ExpandTableColumn(#"Expanded {0}4", "project_activity_notes", {"Element:Text"}, {"project_activity_notes"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}5",{{"PROJECT_ACTIVITY_ID", type number}, {"project_activity_active", type logical}, {"project_activity_complete", type logical}, {"project_activity_project_id", Int64.Type}, {"project_activity_start_date", type datetimezone}, {"project_activity_end_date", type datetimezone}, {"project_activity_chargeable", type logical}, {"project_activity_lock", type logical}, {"project_activity_warning", type logical}, {"project_activity_group_budget", type logical}, {"project_activity_phase_id", Int64.Type}, {"project_activity_chargeable_date", type datetimezone}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectMaterials = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetProjectMaterials xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetProjectMaterials>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table8,{{"PROJECT_MATERIAL_ID", type number}, {"project_material_date", type datetimezone}, {"project_material_project_id", Int64.Type}, {"project_material_phase_id", Int64.Type}, {"project_material_activity_id", Int64.Type}, {"project_material_article_id", Int64.Type}, {"project_material_expense_code_id", Int64.Type}, {"project_material_supplier_id", Int64.Type}}),
    #"Expanded {0}" = Table.ExpandTableColumn(#"Changed Type", "project_material_unit", {"Element:Text"}, {"project_material_unit.Element:Text"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "project_material_invoice_notes", {"Element:Text"}, {"project_material_invoice_notes.Element:Text"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded {0}1",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectMembers = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetProjectMembers xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetProjectMembers>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table8,{{"project_member_employee_id", Int64.Type}, {"project_member_project_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjectMilestones = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetProjectMilestones xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetProjectMilestones>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table8,{{"PROJECT_MILESTONE_ID", type number}, {"project_milestone_project_id", type number}, {"project_milestone_start_date", type datetimezone}, {"project_milestone_end_date", type datetimezone}, {"project_milestone_complete", type logical}, {"project_milestone_phase_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetProjects = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetProjects xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetProjects>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "project_abbreviated", {"Element:Text"}, {"project_abbreviated"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "project_created_by", {"Element:Text"}, {"project_created_by"}),
    #"Expanded {0}2" = Table.ExpandTableColumn(#"Expanded {0}1", "project_gl_code", {"Element:Text"}, {"project_gl_code"}),
    #"Expanded {0}3" = Table.ExpandTableColumn(#"Expanded {0}2", "project_category", {"Element:Text"}, {"project_category"}),
    #"Expanded {0}4" = Table.ExpandTableColumn(#"Expanded {0}3", "project_notes", {"Element:Text"}, {"project_notes"}),
    #"Expanded {0}5" = Table.ExpandTableColumn(#"Expanded {0}4", "project_order_number", {"Element:Text"}, {"project_order_number"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}5",{{"PROJECT_ID", Int64.Type}, {"project_active", type logical}, {"project_public", type logical}, {"project_customer_id", Int64.Type}, {"project_manager_id", Int64.Type}, {"project_created_date", type datetimezone}, {"project_start_date", type datetimezone}, {"project_end_date", type datetimezone}, {"project_complete", type logical}, {"project_complete_date", type datetimezone}, {"project_chargeable", type logical}, {"project_chargeable_date", type datetimezone}, {"project_currency_rate_id", Int64.Type}, {"project_customer_contact_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetSuppliers = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetSuppliers xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetSuppliers>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "supplier_address", {"Element:Text"}, {"supplier_address"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "supplier_zip_code", {"Element:Text"}, {"supplier_zip_code"}),
    #"Expanded {0}2" = Table.ExpandTableColumn(#"Expanded {0}1", "supplier_city", {"Element:Text"}, {"supplier_city"}),
    #"Expanded {0}3" = Table.ExpandTableColumn(#"Expanded {0}2", "supplier_phone", {"Element:Text"}, {"supplier_phone"}),
    #"Expanded {0}4" = Table.ExpandTableColumn(#"Expanded {0}3", "supplier_fax", {"Element:Text"}, {"supplier_fax"}),
    #"Expanded {0}5" = Table.ExpandTableColumn(#"Expanded {0}4", "supplier_email", {"Element:Text"}, {"supplier_email"}),
    #"Expanded {0}6" = Table.ExpandTableColumn(#"Expanded {0}5", "supplier_homepage", {"Element:Text"}, {"supplier_homepage"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}6",{{"SUPPLIER_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetTasks = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetTasks xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetTasks>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "task_notes", {"Element:Text"}, {"task_notes"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "task_company_category", {"Element:Text"}, {"task_company_category"}),
    #"Expanded {0}2" = Table.ExpandTableColumn(#"Expanded {0}1", "task_opportunity_name", {"Element:Text"}, {"task_opportunity_name"}),
    #"Expanded {0}3" = Table.ExpandTableColumn(#"Expanded {0}2", "task_opportunity_stage_name", {"Element:Text"}, {"task_opportunity_stage_name"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}3",{{"TASK_ID", Int64.Type}, {"task_date", type datetimezone}, {"task_start_time", type time}, {"task_end_time", type time}, {"task_reminder_date", type datetimezone}, {"task_reminder_time", type time}, {"task_company_id", Int64.Type}, {"task_assigned_to_id", Int64.Type}, {"task_opportunity_id", Int64.Type}, {"task_created_by_id", Int64.Type}, {"task_contact_way_id", Int64.Type}, {"task_contact_person_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetTravelTypes = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetTravelTypes xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetTravelTypes>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "travel_type_gl_code", {"Element:Text"}, {"travel_type_gl_code"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}",{{"TRAVEL_TYPE_ID", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetTravelItems = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetTravelItems xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetTravelItems>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Expanded {0}" = Table.ExpandTableColumn(Table8, "travel_item_purpose", {"Element:Text"}, {"travel_item_purpose"}),
    #"Expanded {0}1" = Table.ExpandTableColumn(#"Expanded {0}", "travel_item_invoice_notes", {"Element:Text"}, {"travel_item_invoice_notes"}),
    #"Expanded {0}2" = Table.ExpandTableColumn(#"Expanded {0}1", "travel_item_material_invoice_notes", {"Element:Text"}, {"travel_item_material_invoice_notes"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}2",{{"TRAVEL_ITEM_ID", Int64.Type}, {"travel_item_expense_sheet_id", Int64.Type}, {"travel_item_from_date", type datetimezone}, {"travel_item_to_date", type datetimezone}, {"travel_item_project_id", Int64.Type}, {"travel_item_travel_type_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetUnitTypes = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetUnitTypes xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetUnitTypes>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table8,{{"UNIT_TYPE_ID", Int64.Type}, {"unit_type_parent_id", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"UNIT_TYPE_ID", "unit_type_parent_id", "urn:schemas-microsoft-com:xml-diffgram-v1", "urn:schemas-microsoft-com:xml-msdata"})
in
    #"Removed Columns";

GetWorkingTime = () as table =>
let
   body = Text.ToBinary("<?xml version='1.0' encoding='utf-8'?>
<soap12:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap12='http://www.w3.org/2003/05/soap-envelope'>
  <soap12:Body>
	<GetWorkingTime xmlns='https://bin.qbis.se/WS/QBIS_DS'>
	  <Company>" & Extension.CurrentCredential()[Username] & "</Company>
	  <Password>" & Extension.CurrentCredential()[Password] & "</Password>
	</GetWorkingTime>
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
    Table2 = Table1{0}[Table],
    Table3 = Table2{0}[Table],
    Table4 = Table3{1}[Table],
    Table5 = Table4{0}[Table],
    Table6 = Table5{0}[Table],
    Table7 = Table6{0}[Table],
    Table8 = Table7{0}[Table],
    #"Changed Type" = Table.TransformColumnTypes(Table8,{{"WORKING_TIME_ID", Int64.Type}, {"working_time_arrived", type datetimezone}, {"working_time_left", type datetimezone}}),
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