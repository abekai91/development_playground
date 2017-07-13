Class Analysis
   Private rack_name
    Private rack_no
    Private analysis_detail
	Private analysis_fill_oms
	Private GUID
	Private uid
	
	Private batch_id
	Private prod_id
	Private cyl_id
	Private cyl_to_analys
	
	Private analyzer_res_position
	Private analyzer_res 
	
'Description : Class Analysis Constructor
'---------------------------------------------------------------
    Public Default Function Init(parameters)
         Select Case UBound(parameters)
             Case 5 
             	Set Init = InitFiveParam(parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5))
             Case Else
                Set Init = Me
         End Select
    End Function
    
    Private Function InitFiveParam(parameter1, parameter2, parameter3, parameter4, parameter5, parameter6)
    	rack_name 		 	= parameter1
        rack_no 		 	= parameter2
        analysis_detail		= parameter3
		analysis_fill_oms	= parameter4
        GUID				= parameter5
		uid					= parameter6
		
        Call analysis_filter
        Set InitFiveParam = Me
    End Function	

'Description : Filter the Analysis Information
'---------------------------------------------------------------	
	Public Function analysis_filter()
	On Error Resume Next
	Dim analysis_str , cnt
		
		analysis_str = Split(analysis_detail,"|")
		cnt 		 = ubound(analysis_str)
		
		For j = 0 to cnt
			cyl_id		  = analysis_str(0)  
			batch_id      = analysis_str(1)
			prod_id       = analysis_str(2)
			cyl_to_analys = analysis_str(3)
		Next
		If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function analysis_filter is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
		End If
	End Function
	
'Description : Start OMS Filling 
'---------------------------------------------------------------
	
	Public Function fill_by_oms()
	On Error Resume Next
	Dim Tags 
		Tags = Analysis_DB_Tag(rack_no,"Batch_ID") 
		HMIRuntime.Tags(Tags).Write batch_id
		
		Tags = Analysis_DB_Tag(rack_no,"Prod_Name") 
		HMIRuntime.Tags(Tags).Write prod_id
		
		Tags = Analysis_DB_Tag(rack_no,"Cyl_ID") 
		HMIRuntime.Tags(Tags).Write cyl_id
		
		Tags = Analysis_DB_Tag(rack_no,"Cyl_To_Check") 
		HMIRuntime.Tags(Tags).Write cyl_to_analys
		
		Call GF_LogToFile_("Execute", " Sent Analysis Recipe : " & rack_no , "Analysis")
		Call GF_LogToFile_("Execute", "[" & rack_no & "|" & batch_id & "|" & prod_id & "|" & cyl_id & "|" & cyl_to_analys & "(Count)]"  , "Analysis")
		
		If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function fill_by_oms is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
		End If
		Call fill_by_oms_deactivate(0)
		If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function fill_by_oms_deactivate is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
		End If
	End Function
	
	Public Function fill_by_oms_deactivate(stat)
		On Error Resume Next
		If stat = 0 Then
			Call Mysql_Non_Query("Update analysis_rack Set oms_fill = 0 Where rack_id = "& rack_no & "")
			Call GF_LogToFile_("Execute", "Set Database oms_fill = 0"  , "Analysis")
		
		End If
		
		If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function fill_by_oms_deactivate is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
		End If
	End Function
	
	'Description : Deactivate Analysis
'---------------------------------------------------------------	
		Public Function deactivate_analysis_1()
			On Error Resume Next
			start_time = Time
			start_date = Date
			Call Mysql_Non_Query("Update analysis_rack Set cp1 = 0 , cp2 = 1 ,cp3 = 0 , cp4 = 0 Where rack_id = "& rack_no & "")
			
			Call GF_LogToFile_("Execute", "Deactivate Analysis CP 1 : " & rack_no  , "Analysis")
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function deactivate_analysis_1 is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
		End Function
		
		Public Function deactivate_analysis_2()
			On Error Resume Next
			end_time = Time
			end_date = Date
			Call Mysql_Non_Query("Update analysis_rack Set cp1 = 0 , cp2 = 0 , cp3 = 1 , cp4 = 0 Where rack_id = "& rack_no & "")
			Call GF_LogToFile_("Execute", "Deactivate CP 2 : " & rack_no ,"Analysis")
		
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function deactivate_analysis_2 is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
		End Function
		
		Public Function deactivate_analysis_3()
			On Error Resume Next
			Call Mysql_Non_Query("Update analysis_rack Set cp1 = 0 , cp2 = 0 , cp3 = 0 , cp4 = 1 , prod_id = '',cyl_id = 0  Where rack_id = "& rack_no & "")
			Call GF_LogToFile_("Execute", "Deactivate CP 3 : " & rack_no ,"Analysis")
		
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function deactivate_analysis_3 is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
		End Function
		
		Public Function deactivate_analysis_4()
			On Error Resume Next
			Call Mysql_Non_Query("Update analysis_qc_history SET status = 2 Where rack_name = '"& rack_name &"' and  oms_batch = '"& batch_id &"' ")
			Call Mysql_Non_Query("Update analysis_rack Set occupied = 0 , cp1 = 0 , cp2 = 0 , cp3 = 0 , cp4 = 0 , cyl_id = 0 , batch_id = '' , prod_id = '' ,cyl_to_analyse = 0 ,user_id = 0 Where rack_id = "& rack_no & "")
			
			Call GF_LogToFile_("Execute", "Deactivate CP 4 : " & rack_no ,"Analysis")
		
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function deactivate_analysis_4 is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
		End Function
		
'Description : Let Property for Analysis
'---------------------------------------------------------------
Dim Tags 
	Public Property Let start_time(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"start_time") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let start_date(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"start_date") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let end_time(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"end_time") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let end_date(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"end_date") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_1_val(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_1_val") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_1_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_1_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_2_val(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_2_val") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_2_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_2_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_3_val(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_3_val") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_3_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_3_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_4_val(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_4_val") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_4_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_4_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_5_val(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_5_val") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_5_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_5_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_6_val(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_6_val") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let analyser_6_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"analyser_6_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let pressure_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"pressure_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let temperature_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"temperature_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let sticker_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"sticker_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let pressure_cyl_result(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"pressure_cyl_result") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let filling_class(ByVal value)
		Tags = Analysis_DB_Tag(rack_no,"filling_class") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
'Description : Get Property for Analysis
'---------------------------------------------------------------
	Public Property Get pressure_cyl_result
		Tags = Analysis_DB_Tag(rack_no,"pressure_cyl_result") 
		pressure_cyl_result= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get rack_name_get
		rack_name_get = rack_name
	End property
	
	Public Property Get prod_id_get
		prod_id_get = prod_id
	End Property	
		
	public Property Get fill_oms
		fill_oms = analysis_fill_oms
	End Property
	
	Public Property Get batch_id_get
		Tags = Analysis_DB_Tag(rack_no,"Batch_ID") 
		batch_id_get = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get prod_name
		Tags = Analysis_DB_Tag(rack_no,"Prod_Name") 
		prod_name = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get cylinder_id
		Tags = Analysis_DB_Tag(rack_no,"Cyl_ID") 
		cylinder_id = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get cylinder_to_analys
		Tags = Analysis_DB_Tag(rack_no,"Cyl_To_Check") 
		cylinder_to_analys = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get cyl_counter
		Tags = Analysis_DB_Tag(rack_no,"Cyl_Counter") 
		cyl_counter = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get start
		Tags = Analysis_DB_Tag(rack_no,"Start") 
		start = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get print_st 
		Tags = Analysis_DB_Tag(rack_no,"Print_St") 
		print_st = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get capture 
		Tags = Analysis_DB_Tag(rack_no,"Capture") 
		capture = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get complete 
		Tags = Analysis_DB_Tag(rack_no,"Complete") 
		complete = HMIRuntime.Tags(Tags).Read
	End Property 
	
	Public Property Get pressure
		Tags = Analysis_DB_Tag(rack_no,"Pressure") 
		pressure = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get temperature
		Tags = Analysis_DB_Tag(rack_no,"Temperature") 
		temperature = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_CH4
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace_CH4"
		AYM_CH4 = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_O2_100_PERCENT
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Purity1_O2"
		AYM_O2_100_PERCENT = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_O2_21_PERCENT
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_21_Percent_O2"
		AYM_O2_21_PERCENT = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_O2_50_PERCENT
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_50_Percent_O2"
		AYM_O2_50_PERCENT = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_CO
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace_CO"
		AYM_CO  = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_CO2
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace_CO2"
		AYM_CO2 = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_H2O
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace1_H2O"
		AYM_H2O = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_N2O
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace_Nox"
		AYM_N2O = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_SPARE1
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Standby1"
		AYM_SPARE1 = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYM_SPARE2
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Standby2"
		AYM_SPARE2 = HMIRuntime.Tags(Tags).Read
	End Property
	
'Industry Tags	
	Public Property Get AYI_O2_PPM
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Purity2_O2"
		AYI_O2_PPM = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_O2_IN_N2
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace_O2_in_N2"
		AYI_O2_IN_N2 = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_O2_IN_AR
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace_O2_in_Ar"
		AYI_O2_IN_AR = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_H2O
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Trace2_H2O"
		AYI_H2O = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_CO_PERCENT
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_CO_Percent"
		AYI_CO_PERCENT = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_N2_PERCENT
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_N2_Percent"
		AYI_N2_PERCENT = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_O2_PERCENT
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_O2_Percent"
		AYI_O2_PERCENT = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_HE_IN_N2
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_He_in_N2"
		AYI_HE_IN_N2 = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_HE_IN_AR
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_He_in_Ar"
		AYI_HE_IN_AR = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get AYI_CO2
		Tags	= "PLC_01/DB_PLC_AY_VAL.Analysis_Purity_CO2"
		AYI_CO2 = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get sticker_result
		Tags = Analysis_DB_Tag(rack_no,"sticker_result") 
		sticker_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get pressure_result
		Tags = Analysis_DB_Tag(rack_no,"pressure_result") 
		pressure_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get temperature_result
		Tags = Analysis_DB_Tag(rack_no,"temperature_result") 
		temperature_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get start_time
		Tags = Analysis_DB_Tag(rack_no,"start_time") 
		start_time = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get start_date
		Tags = Analysis_DB_Tag(rack_no,"start_date") 
		start_date = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get end_time
		Tags = Analysis_DB_Tag(rack_no,"end_time") 
		end_time = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get end_date
		Tags = Analysis_DB_Tag(rack_no,"end_date") 
		end_date = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_1_val
		Tags = Analysis_DB_Tag(rack_no,"analyser_1_val") 
		analyser_1_val = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_1_result
		Tags = Analysis_DB_Tag(rack_no,"analyser_1_result") 
		analyser_1_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_2_val
		Tags = Analysis_DB_Tag(rack_no,"analyser_2_val") 
		analyser_2_val = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_2_result
		Tags = Analysis_DB_Tag(rack_no,"analyser_2_result") 
		analyser_2_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_3_val
		Tags = Analysis_DB_Tag(rack_no,"analyser_3_val") 
		analyser_3_val = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_3_result
		Tags = Analysis_DB_Tag(rack_no,"analyser_3_result") 
		analyser_3_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_4_val
		Tags = Analysis_DB_Tag(rack_no,"analyser_4_val") 
		analyser_4_val = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_4_result
		Tags = Analysis_DB_Tag(rack_no,"analyser_4_result") 
		analyser_4_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_5_val
		Tags = Analysis_DB_Tag(rack_no,"analyser_5_val") 
		analyser_5_val = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_5_result
		Tags = Analysis_DB_Tag(rack_no,"analyser_5_result") 
		analyser_5_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_6_val
		Tags = Analysis_DB_Tag(rack_no,"analyser_6_val") 
		analyser_6_val = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get analyser_6_result
		Tags = Analysis_DB_Tag(rack_no,"analyser_6_result") 
		analyser_6_result = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get filling_class
		Tags = Analysis_DB_Tag(rack_no,"filling_class") 
		filling_class = HMIRuntime.Tags(Tags).Read
	End Property

	'Description : Generate Reports
'---------------------------------------------------------------
	Public Function generate_reports()
	On Error Resume Next
	Dim masterlist
		
		Set masterlist = New MasterList_Reports 
		'Set masterlist.Individual  = New Individual
		Set masterlist.Loose_Palletize = New Loose_Palletize
		
		'Run Loose and Palletize scripts
		masterlist.Loose_Palletize.Batch_Id = batch_id
		masterlist.Loose_Palletize.GUID = GUID
		masterlist.Loose_Palletize.Generate_Report
		
		Call GF_LogToFile_("Execute", " : Generate Reports Excel" ,"Analysis")	
	
		If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function generate_reports is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
		End If
	End Function

	'Description : Analysis Rules
'---------------------------------------------------------------		
	Public Function analysis_rules()
	On Error Resume Next
	Dim parking,masterlist
	
		parking = Left(rack_name,3)
		
		Set masterlist = New MasterList_Analysis 
		Set masterlist.medical  = New Medical
		Set masterlist.industry = New Industry
			
		If parking = "AYM" Then
			
			masterlist.medical.product =  prod_id
			masterlist.medical.pressure = pressure
			masterlist.medical.temperature = temperature
			pressure_cyl_result = masterlist.medical.pressure_result
			filling_class = masterlist.medical.filling_clasification
			HMIRuntime.Trace("pressure result : " & pressure_result)
			
			Call keep_into_local_tag(masterlist.medical.analyzer_name,masterlist.medical.analyzer_val,masterlist.medical.analyzer_number)
			
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function keep_into_local_tag is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
		
		Elseif parking = "AYI" Then
		
			masterlist.industry.product = prod_id
			masterlist.industry.pressure = pressure
			masterlist.industry.temperature = temperature
			pressure_cyl_result = masterlist.industry.pressure_result
			filling_class = masterlist.industry.filling_clasification
			
			Call keep_into_local_tag(masterlist.industry.analyzer_name,masterlist.industry.analyzer_val,masterlist.industry.analyzer_number)
			
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function keep_into_local_tag is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
		End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function analysis_rules is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
	End Function
	
	Public Function keep_into_local_tag(Byval analyzer_name, Byval analyzer_val, Byval analyzer_number)
	On Error Resume Next
	Dim als_name(10)
	Dim als_val(10)
	Dim als_num(10)
	Dim als_val_min(10)
	Dim als_val_max(10)
	
	'Testing Purpose
	HMIRuntime.Trace("Analyzer Value : " & analyzer_val & vbCrlf)
	HMIRuntime.Trace("Analyzer Name : " & analyzer_name & vbCrlf)
	HMIRuntime.Trace("Analyzer Number : " & analyzer_number & vbCrlf)
	
	HMIRuntime.Trace("Analyzer AYM_N2O : " & AYM_N2O & vbCrlf)
	HMIRuntime.Trace("Analyzer AYM_N2O(Convert) : " & CDbl(AYM_N2O) & vbCrlf)
	
	HMIRuntime.Trace("Analyzer AYM_O2_21_PERCENT : " & AYM_O2_21_PERCENT & vbCrlf)
	HMIRuntime.Trace("Analyzer AYM_O2_21_PERCENT(Convert) : " & CDbl(AYM_O2_21_PERCENT) & vbCrlf)
	
	HMIRuntime.Trace("Analyzer AYM_O2_50_PERCENT : " & AYM_O2_50_PERCENT & vbCrlf)
	HMIRuntime.Trace("Analyzer AYM_O2_50_PERCENT(Convert) : " & CDbl(AYM_O2_50_PERCENT) & vbCrlf)
	
	HMIRuntime.Trace("Analyzer AYM_O2_100_PERCENT : " & AYM_O2_100_PERCENT & vbCrlf)
	HMIRuntime.Trace("Analyzer AYM_O2_100_PERCENT(Convert) : " & CDbl(AYM_O2_100_PERCENT) & vbCrlf)
	
	HMIRuntime.Trace("Analyzer AYM_H2O : " & AYM_H2O & vbCrlf)
	HMIRuntime.Trace("Analyzer AYM_H2O(Convert) : " & CDbl(AYM_H2O) & vbCrlf)
	
	HMIRuntime.Trace("Analyzer AYM_CO2 : " & AYM_CO2 & vbCrlf)
	HMIRuntime.Trace("Analyzer AYM_CO2(Convert) : " & CDbl(AYM_CO2) & vbCrlf)
	
	HMIRuntime.Trace("Analyzer AYM_CO : " & AYM_CO & vbCrlf)
	HMIRuntime.Trace("Analyzer AYM_CO(Convert) : " & CDbl(AYM_CO) & vbCrlf)
	'End Testing Purpose
	
	Dim filter_name	, filter_val , filter_num, cnt_1, i ,qty_1
	Dim cnt_2, qty_2
	Dim filter_val_2 , qty_3
		
		analyzer_name = Replace(analyzer_name,",NULL","")
		analyzer_res_position = Replace(analyzer_number,",NULL","")
	
	'filter the analyzer name 
		filter_name = Split(analyzer_name,",")
		cnt_1 = ubound(filter_name)
		qty_1 = cnt_1
		
		For i = 0 to cnt_1
			als_name(i) = filter_name(i)
		Next
		
	'filter the analyzer number
		filter_num = Split(analyzer_number,",")
		cnt_1 = ubound(filter_num)
		
		For i = 0 to cnt_1
			als_num(i) = filter_num(i)
		Next
	
	'filter the analyzer value
		filter_val = Split(analyzer_val,",")
		cnt_2 = ubound(filter_val)
		qty_2 = cnt_2
		
		For i = 0 to qty_2
			als_val(i) = filter_val(i)
		Next
		
		For i = 0 To qty_2
			filter_val_2 = Split(als_val(i),"|")
			qty_3 = ubound(filter_val_2)

			For j = 0 to qty_3
				als_val_min(i) = filter_val_2(0)
				als_val_max(i) = filter_val_2(1)
			Next
		Next
		
	'store all of the information into internal tags	
		For i=0 to qty_1
			Select Case als_name(i)
				Case ""
						If als_num(i) =  1 Then
							analyser_1_val 	  = 0
							analyser_1_result = "PASS"
						End If
						If als_num(i) =  2 Then
							analyser_2_val 	  = 0
							analyser_2_result = "PASS"
						End If
						If als_num(i) =  3 Then
							analyser_3_val 	  = 0
							analyser_3_result = "PASS"
						End If
						If als_num(i) =  4 Then
							analyser_4_val 	  = 0
							analyser_4_result = "PASS"
						End If
						If als_num(i) =  5 Then
							analyser_5_val 	  = 0
							analyser_5_result = "PASS"
						End If
						If als_num(i) =  6 Then
							analyser_6_val 	  = 0
							analyser_6_result = "PASS"
						End If
						sticker_result = "PASS"
						
				'Medical		
				Case "AYM_N2O"
						If CDbl(AYM_N2O) >= CDbl(als_val_min(i)) And CDbl(AYM_N2O) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_N2O
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_N2O
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_N2O
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_N2O
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_N2O
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_N2O
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_N2O
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_N2O
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_N2O
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_N2O
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_N2O
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_N2O
								analyser_6_result = "FAIL"
							End If 
							
							sticker_result = "FAIL"
							
						End IF
				Case "AYM_O2_21_PERCENT"
						If CDbl(AYM_O2_21_PERCENT) >= CDbl(als_val_min(i)) And CDbl(AYM_O2_21_PERCENT) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_O2_21_PERCENT
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_O2_21_PERCENT
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_O2_21_PERCENT
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_O2_21_PERCENT
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_O2_21_PERCENT
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_O2_21_PERCENT
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_O2_21_PERCENT
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_O2_21_PERCENT
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_O2_21_PERCENT
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_O2_21_PERCENT
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_O2_21_PERCENT
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_O2_21_PERCENT
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYM_O2_50_PERCENT"
						If CDbl(AYM_O2_50_PERCENT) >= CDbl(als_val_min(i)) And CDbl(AYM_O2_50_PERCENT) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_O2_50_PERCENT
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_O2_50_PERCENT
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_O2_50_PERCENT
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_O2_50_PERCENT
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_O2_50_PERCENT
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_O2_50_PERCENT
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_O2_50_PERCENT
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_O2_50_PERCENT
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_O2_50_PERCENT
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_O2_50_PERCENT
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_O2_50_PERCENT
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_O2_50_PERCENT
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYM_O2_100_PERCENT"
						If CDbl(AYM_O2_100_PERCENT) >= CDbl(als_val_min(i)) And CDbl(AYM_O2_100_PERCENT) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_O2_100_PERCENT
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_O2_100_PERCENT
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_O2_100_PERCENT
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_O2_100_PERCENT
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_O2_100_PERCENT
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_O2_100_PERCENT
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_O2_100_PERCENT
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_O2_100_PERCENT
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_O2_100_PERCENT
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_O2_100_PERCENT
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_O2_100_PERCENT
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_O2_100_PERCENT
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYM_H2O"
						If CDbl(AYM_H2O) >= CDbl(als_val_min(i)) And CDbl(AYM_H2O) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_H2O
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_H2O
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_H2O
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_H2O
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_H2O
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_H2O
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_H2O
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_H2O
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_H2O
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_H2O
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_H2O
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_H2O
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYM_CO2"
						If CDbl(AYM_CO2) >= CDbl(als_val_min(i)) And CDbl(AYM_CO2) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_CO2
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_CO2
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_CO2
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_CO2
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_CO2
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_CO2
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_CO2
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_CO2
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_CO2
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_CO2
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_CO2
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_CO2
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYM_CO"
						If CDbl(AYM_CO) >= CDbl(als_val_min(i)) And CDbl(AYM_CO) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_CO
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_CO
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_CO
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_CO
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_CO
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_CO
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYM_CO
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYM_CO
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYM_CO
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYM_CO
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYM_CO
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYM_CO
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				'Industry
				Case "AYI_CO2_PERCENT"
						If CDbl(AYI_CO2) >= CDbl(als_val_min(i)) And CDbl(AYI_CO2) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_CO2
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_CO2
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_CO2
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_CO2
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_CO2
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_CO2
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_CO2
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_CO2
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_CO2
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_CO2
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_CO2
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_CO2
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYI_O2_PERCENT"
						If CDbl(AYI_O2_PERCENT) >= CDbl(als_val_min(i)) And CDbl(AYI_O2_PERCENT) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_O2_PERCENT
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_O2_PERCENT
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_O2_PERCENT
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_O2_PERCENT
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_O2_PERCENT
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_O2_PERCENT
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_O2_PERCENT
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_O2_PERCENT
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_O2_PERCENT
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_O2_PERCENT
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_O2_PERCENT
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_O2_PERCENT
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYI_O2_PPM"
						If CDbl(AYI_O2_PPM) >= CDbl(als_val_min(i)) And CDbl(AYI_O2_PPM) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_O2_PPM
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_O2_PPM
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_O2_PPM
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_O2_PPM
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_O2_PPM
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_O2_PPM
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_O2_PPM
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_O2_PPM
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_O2_PPM
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_O2_PPM
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_O2_PPM
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_O2_PPM
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYI_HE_IN_N2"
						If CDbl(AYI_HE_IN_N2) >= CDbl(als_val_min(i)) And CDbl(AYI_HE_IN_N2) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_HE_IN_N2
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_HE_IN_N2
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_HE_IN_N2
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_HE_IN_N2
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_HE_IN_N2
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_HE_IN_N2
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_HE_IN_N2
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_HE_IN_N2
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_HE_IN_N2
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_HE_IN_N2
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_HE_IN_N2
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_HE_IN_N2
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYI_N2_PERCENT"
						If CDbl(AYI_N2_PERCENT) >= CDbl(als_val_min(i)) And CDbl(AYI_N2_PERCENT) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_N2_PERCENT
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_N2_PERCENT
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_N2_PERCENT
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_N2_PERCENT
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_N2_PERCENT
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_N2_PERCENT
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_N2_PERCENT
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_N2_PERCENT
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_N2_PERCENT
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_N2_PERCENT
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_N2_PERCENT
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_N2_PERCENT
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYI_HE_IN_AR"
						If CDbl(AYI_HE_IN_AR) >= CDbl(als_val_min(i)) And CDbl(AYI_HE_IN_AR) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_HE_IN_AR
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_HE_IN_AR
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_HE_IN_AR
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_HE_IN_AR
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_HE_IN_AR
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_HE_IN_AR
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_HE_IN_AR
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_HE_IN_AR
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_HE_IN_AR
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_HE_IN_AR
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_HE_IN_AR
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_HE_IN_AR
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYI_H2O"
						If CDbl(AYI_H2O) >= CDbl(als_val_min(i)) And CDbl(AYI_H2O) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_H2O
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_H2O
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_H2O
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_H2O
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_H2O
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_H2O
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_H2O
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_H2O
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_H2O
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_H2O
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_H2O
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_H2O
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case "AYI_CO2"
						If CDbl(AYI_CO2) >= CDbl(als_val_min(i)) And CDbl(AYI_CO2) <= CDbl(als_val_max(i)) Then
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_CO2
								analyser_1_result = "PASS"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_CO2
								analyser_2_result = "PASS"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_CO2
								analyser_3_result = "PASS"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_CO2
								analyser_4_result = "PASS"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_CO2
								analyser_5_result = "PASS"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_CO2
								analyser_6_result = "PASS"
							End If
							
							sticker_result = "PASS"
						Else
							If als_num(i) =  1 Then
								analyser_1_val 	  = AYI_CO2
								analyser_1_result = "FAIL"
							End If
							If als_num(i) =  2 Then
								analyser_2_val 	  = AYI_CO2
								analyser_2_result = "FAIL"
							End If
							If als_num(i) =  3 Then
								analyser_3_val 	  = AYI_CO2
								analyser_3_result = "FAIL"
							End If
							If als_num(i) =  4 Then
								analyser_4_val 	  = AYI_CO2
								analyser_4_result = "FAIL"
							End If
							If als_num(i) =  5 Then
								analyser_5_val 	  = AYI_CO2
								analyser_5_result = "FAIL"
							End If
							If als_num(i) =  6 Then
								analyser_6_val 	  = AYI_CO2
								analyser_6_result = "FAIL"
							End If
							
							sticker_result = "FAIL"
						End IF
				Case Else
					HMIRuntime.trace(Now & " CheckPoint(2) Keep Local Tag Not Match ("& analyzer_name & ")" & vbCrlf )
			End Select
		Next
		
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function keep_into_local_tag is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
			
	End Function
	
'Description : Set Pressure and Temperature Value
'---------------------------------------------------------------
	Public Function analysis_pressure_temperature()
		On Error Resume Next
			pressure_result = pressure 
			temperature_result = temperature
			
			Call GF_LogToFile_("Execute", "Pressure : " & pressure & "| Temperature : " & temperature ,"Analysis")	
			
			If Err.Number <> 0 Then
		    	Call GF_LogError("Error", "Analysis.bmo - Function analysis_pressure_temperature is not Workings [" & Err.Description & "]","Analysis")
		    	Err.Clear
			End If
		
	End Function

'Description : Print Stickers
'---------------------------------------------------------------
	Public Function print_sticker()
	On Error Resume Next
	Dim stick_res
		Dim res : res = analyser_1_result & "|" & analyser_2_result & "|" & analyser_3_result & "|" & analyser_4_result & "|" & analyser_5_result & "|" & analyser_6_result
		
		If res = "PASS|PASS|PASS|PASS|PASS|PASS" Then
			stick_res = "PASS"
		Else
			stick_res = "FAIL"
		End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function print_sticker is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
		Call generate_sticker(stick_res)
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function generate_sticker is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
	'Description : Dataset for Analysis DB Tag
'---------------------------------------------------------------	
	Private Function Analysis_DB_Tag(Byval index, Byval DB_name)
	On Error Resume Next 
		If DB_name = "Batch_ID" Then	
				Dim Batch_ID
					Batch_ID 		= Array("RESERVED SPACE" , _
											"PLC_01/DB_OMS_AYM_1.Batch_ID" , _
											"PLC_01/DB_OMS_AYM_2-3.Batch_ID" , _
											"PLC_01/DB_OMS_AYM_4-5.Batch_ID" , _
											"PLC_01/DB_OMS_AYI_1-2.Batch_ID" , _
											"PLC_01/DB_OMS_AYI_3-4-5.Batch_ID" , _
											"PLC_01/DB_OMS_AYI_6.Batch_ID" ) 
					Analysis_DB_Tag = Batch_ID(index)	
					
		ElseIf DB_name = "Prod_Name" Then		
				Dim Prod_Name
					Prod_Name 			= Array("RESERVED SPACE" , _
												"PLC_01/DB_OMS_AYM_1.Prod_Name" , _
												"PLC_01/DB_OMS_AYM_2-3.Prod_Name" , _
												"PLC_01/DB_OMS_AYM_4-5.Prod_Name" , _
												"PLC_01/DB_OMS_AYI_1-2.Prod_Name" , _
												"PLC_01/DB_OMS_AYI_3-4-5.Prod_Name" , _
												"PLC_01/DB_OMS_AYI_6.Prod_Name")
					Analysis_DB_Tag = Prod_Name(index)	
					
		ElseIf DB_name = "Cyl_ID" Then
				Dim Cyl_ID							
					Cyl_ID 			= Array("RESERVED SPACE" , _
											"PLC_01/DB_OMS_AYM_1.Cyl_ID" ,  _
											"PLC_01/DB_OMS_AYM_2-3.Cyl_ID" ,  _
											"PLC_01/DB_OMS_AYM_4-5.Cyl_ID" ,  _
											"PLC_01/DB_OMS_AYI_1-2.Cyl_ID" ,  _
											"PLC_01/DB_OMS_AYI_3-4-5.Cyl_ID" ,  _
											"PLC_01/DB_OMS_AYI_6.Cyl_ID" )
					Analysis_DB_Tag = Cyl_ID(index)
					
		ElseIf DB_name = "Cyl_To_Check" Then
				Dim Cyl_To_Check 
					Cyl_To_Check 	= Array("RESERVED SPACE" , _
											"PLC_01/DB_OMS_AYM_1.Cyl_to_check" , _
											"PLC_01/DB_OMS_AYM_2-3.Cyl_to_check" , _
											"PLC_01/DB_OMS_AYM_4-5.Cyl_to_check" , _
											"PLC_01/DB_OMS_AYI_1-2.Cyl_to_check" , _
											"PLC_01/DB_OMS_AYI_3-4-5.Cyl_to_check" , _
											"PLC_01/DB_OMS_AYI_6.Cyl_to_check")
					Analysis_DB_Tag = Cyl_To_Check(index)
					
		ElseIf DB_name = "Cyl_Counter" Then	
				Dim Cyl_Counter
					Cyl_Counter 	= Array("RESERVED SPACE" , _
											"PLC_01/DB_PLC_AYM_1.Cyl_Counter" ,  _
											"PLC_01/DB_PLC_AYM_2-3.Cyl_Counter" ,  _
											"PLC_01/DB_PLC_AYM_4-5.Cyl_Counter" ,  _
											"PLC_01/DB_PLC_AYI_1-2.Cyl_Counter" ,  _
											"PLC_01/DB_PLC_AYI_3-4-5.Cyl_Counter" ,  _
											"PLC_01/DB_PLC_AYI_6.Cyl_Counter")
					Analysis_DB_Tag = Cyl_Counter(index)								

		ElseIf DB_name = "Start" Then												
				Dim Start	
					Start 			= Array("RESERVED SPACE" , _
											"PLC_01/DB_PLC_AYM_1.Start" , _
											"PLC_01/DB_PLC_AYM_2-3.Start" , _
											"PLC_01/DB_PLC_AYM_4-5.Start" , _
											"PLC_01/DB_PLC_AYI_1-2.Start" , _
											"PLC_01/DB_PLC_AYI_3-4-5.Start" , _
											"PLC_01/DB_PLC_AYI_6.Start")
					Analysis_DB_Tag = Start(index)	
					
		ElseIf DB_name = "Print_St" Then	'Datablock 303											
				Dim Print_St	
					Print_St 			= Array("RESERVED SPACE" , _
												"PLC_01/DB_PLC_AYM_1.Print" , _
												"PLC_01/DB_PLC_AYM_2-3.Print" , _
												"PLC_01/DB_PLC_AYM_4-5.Print" , _
												"PLC_01/DB_PLC_AYI_1-2.Print" , _
												"PLC_01/DB_PLC_AYI_3-4-5.Print" , _
												"PLC_01/DB_PLC_AYI_6.Print")
					Analysis_DB_Tag = Print_St(index)
					
		ElseIf DB_name = "Capture" Then												
				Dim Capture	
					Capture 			= Array("RESERVED SPACE" , _
												"PLC_01/DB_PLC_AYM_1.Capture" ,  _
												"PLC_01/DB_PLC_AYM_2-3.Capture" ,  _
												"PLC_01/DB_PLC_AYM_4-5.Capture" ,  _
												"PLC_01/DB_PLC_AYI_1-2.Capture" ,  _
												"PLC_01/DB_PLC_AYI_3-4-5.Capture" ,  _
												"PLC_01/DB_PLC_AYI_6.Capture")
					Analysis_DB_Tag = Capture(index)
					
		ElseIf DB_name = "Complete" Then												
				Dim Complete	
					Complete 			= Array("RESERVED SPACE" , _
												"PLC_01/DB_PLC_AYM_1.Complete" ,  _
												"PLC_01/DB_PLC_AYM_2-3.Complete" ,  _
												"PLC_01/DB_PLC_AYM_4-5.Complete" ,  _
												"PLC_01/DB_PLC_AYI_1-2.Complete" ,  _
												"PLC_01/DB_PLC_AYI_3-4-5.Complete" ,  _
												"PLC_01/DB_PLC_AYI_6.Complete")
					Analysis_DB_Tag = Complete(index)
		
		ElseIf DB_name = "Pressure" Then										
				Dim Pressure	
					Pressure 			= Array("RESERVED SPACE" , _
												"PLC_01/DB_PLC_AY_VAL.PT04" ,  _
												"PLC_01/DB_PLC_AY_VAL.PT05" ,  _
												"PLC_01/DB_PLC_AY_VAL.PT06" ,  _
												"PLC_01/DB_PLC_AY_VAL.PT01" ,  _
												"PLC_01/DB_PLC_AY_VAL.PT02" ,  _
												"PLC_01/DB_PLC_AY_VAL.PT03")
					Analysis_DB_Tag = Pressure(index)
		
		ElseIf DB_name = "Temperature" Then										
				Dim Temperature	
					Temperature 			= Array("RESERVED SPACE" , _
												"PLC_01/DB_PLC_AY_VAL.TT01" ,  _
												"PLC_01/DB_PLC_AY_VAL.TT02" ,  _
												"PLC_01/DB_PLC_AY_VAL.TT03" ,  _
												"PLC_01/DB_PLC_AY_VAL.TT04" ,  _
												"PLC_01/DB_PLC_AY_VAL.TT05" ,  _
												"PLC_01/DB_PLC_AY_VAL.TT06")
					Analysis_DB_Tag = Temperature(index)
	
	
	'Internal Tags Wincc
	
		ElseIf DB_name = "start_time" Then										
				Dim start_time	
					start_time 			= Array("RESERVED SPACE" , _
												"start_time_analysis_1" ,  _
												"start_time_analysis_2" ,  _
												"start_time_analysis_3" ,  _
												"start_time_analysis_4" ,  _
												"start_time_analysis_5" ,  _
												"start_time_analysis_6")
					Analysis_DB_Tag = start_time(index)
		
		ElseIf DB_name = "start_date" Then										
				Dim start_date	
					start_date 			= Array("RESERVED SPACE" , _
												"start_date_analysis_1" ,  _
												"start_date_analysis_2" ,  _
												"start_date_analysis_3" ,  _
												"start_date_analysis_4" ,  _
												"start_date_analysis_5" ,  _
												"start_date_analysis_6")
					Analysis_DB_Tag = start_date(index)
		
		ElseIf DB_name = "end_time" Then										
				Dim end_time	
					end_time 			= Array("RESERVED SPACE" , _
												"end_time_analysis_1" ,  _
												"end_time_analysis_2" ,  _
												"end_time_analysis_3" ,  _
												"end_time_analysis_4" ,  _
												"end_time_analysis_5" ,  _
												"end_time_analysis_6")
					Analysis_DB_Tag = end_time(index)
		
		ElseIf DB_name = "end_date" Then										
				Dim end_date	
					end_date 			= Array("RESERVED SPACE" , _
												"end_date_analysis_1" ,  _
												"end_date_analysis_2" ,  _
												"end_date_analysis_3" ,  _
												"end_date_analysis_4" ,  _
												"end_date_analysis_5" ,  _
												"end_date_analysis_6")
					Analysis_DB_Tag = end_date(index)
		
		ElseIf DB_name = "analyser_1_val" Then										
				Dim analyser_1_val	
					analyser_1_val 			= Array("RESERVED SPACE" , _
												"analysis_val_1_1" ,  _
												"analysis_val_1_2" ,  _
												"analysis_val_1_3" ,  _
												"analysis_val_1_4" ,  _
												"analysis_val_1_5" ,  _
												"analysis_val_1_6")
					Analysis_DB_Tag = analyser_1_val(index)	
		
		ElseIf DB_name = "analyser_1_result" Then										
				Dim analyser_1_result	
					analyser_1_result 			= Array("RESERVED SPACE" , _
												"analysis_res_1_1" ,  _
												"analysis_res_1_2" ,  _
												"analysis_res_1_3" ,  _
												"analysis_res_1_4" ,  _
												"analysis_res_1_5" ,  _
												"analysis_res_1_6")
					Analysis_DB_Tag = analyser_1_result(index)	
					
		ElseIf DB_name = "analyser_2_val" Then										
				Dim analyser_2_val	
					analyser_2_val 			= Array("RESERVED SPACE" , _
												"analysis_val_2_1" ,  _
												"analysis_val_2_2" ,  _
												"analysis_val_2_3" ,  _
												"analysis_val_2_4" ,  _
												"analysis_val_2_5" ,  _
												"analysis_val_2_6")
					Analysis_DB_Tag = analyser_2_val(index)	
		
		ElseIf DB_name = "analyser_2_result" Then										
				Dim analyser_2_result	
					analyser_2_result 			= Array("RESERVED SPACE" , _
												"analysis_res_2_1" ,  _
												"analysis_res_2_2" ,  _
												"analysis_res_2_3" ,  _
												"analysis_res_2_4" ,  _
												"analysis_res_2_5" ,  _
												"analysis_res_2_6")
					Analysis_DB_Tag = analyser_2_result(index)	
					
		ElseIf DB_name = "analyser_3_val" Then										
				Dim analyser_3_val	
					analyser_3_val 			= Array("RESERVED SPACE" , _
												"analysis_val_3_1" ,  _
												"analysis_val_3_2" ,  _
												"analysis_val_3_3" ,  _
												"analysis_val_3_4" ,  _
												"analysis_val_3_5" ,  _
												"analysis_val_3_6")
					Analysis_DB_Tag = analyser_3_val(index)	
		
		ElseIf DB_name = "analyser_3_result" Then										
				Dim analyser_3_result	
					analyser_3_result 			= Array("RESERVED SPACE" , _
												"analysis_res_3_1" ,  _
												"analysis_res_3_2" ,  _
												"analysis_res_3_3" ,  _
												"analysis_res_3_4" ,  _
												"analysis_res_3_5" ,  _
												"analysis_res_3_6")
					Analysis_DB_Tag = analyser_3_result(index)	
					
		ElseIf DB_name = "analyser_4_val" Then										
				Dim analyser_4_val	
					analyser_4_val 			= Array("RESERVED SPACE" , _
												"analysis_val_4_1" ,  _
												"analysis_val_4_2" ,  _
												"analysis_val_4_3" ,  _
												"analysis_val_4_4" ,  _
												"analysis_val_4_5" ,  _
												"analysis_val_4_6")
					Analysis_DB_Tag = analyser_4_val(index)	
		
		ElseIf DB_name = "analyser_4_result" Then										
				Dim analyser_4_result	
					analyser_4_result 			= Array("RESERVED SPACE" , _
												"analysis_res_4_1" ,  _
												"analysis_res_4_2" ,  _
												"analysis_res_4_3" ,  _
												"analysis_res_4_4" ,  _
												"analysis_res_4_5" ,  _
												"analysis_res_4_6")
					Analysis_DB_Tag = analyser_4_result(index)	
					
		ElseIf DB_name = "analyser_5_val" Then										
				Dim analyser_5_val	
					analyser_5_val 			= Array("RESERVED SPACE" , _
												"analysis_val_5_1" ,  _
												"analysis_val_5_2" ,  _
												"analysis_val_5_3" ,  _
												"analysis_val_5_4" ,  _
												"analysis_val_5_5" ,  _
												"analysis_val_5_6")
					Analysis_DB_Tag = analyser_5_val(index)	
		
		ElseIf DB_name = "analyser_5_result" Then										
				Dim analyser_5_result	
					analyser_5_result 			= Array("RESERVED SPACE" , _
												"analysis_res_5_1" ,  _
												"analysis_res_5_2" ,  _
												"analysis_res_5_3" ,  _
												"analysis_res_5_4" ,  _
												"analysis_res_5_5" ,  _
												"analysis_res_5_6")
					Analysis_DB_Tag = analyser_5_result(index)	
					
		ElseIf DB_name = "analyser_6_val" Then										
				Dim analyser_6_val	
					analyser_6_val 			= Array("RESERVED SPACE" , _
												"analysis_val_6_1" ,  _
												"analysis_val_6_2" ,  _
												"analysis_val_6_3" ,  _
												"analysis_val_6_4" ,  _
												"analysis_val_6_5" ,  _
												"analysis_val_6_6")
					Analysis_DB_Tag = analyser_6_val(index)	
		
		ElseIf DB_name = "analyser_6_result" Then										
				Dim analyser_6_result	
					analyser_6_result 			= Array("RESERVED SPACE" , _
												"analysis_res_6_1" ,  _
												"analysis_res_6_2" ,  _
												"analysis_res_6_3" ,  _
												"analysis_res_6_4" ,  _
												"analysis_res_6_5" ,  _
												"analysis_res_6_6")
					Analysis_DB_Tag = analyser_6_result(index)	
					
		ElseIf DB_name = "pressure_result" Then										
				Dim pressure_result	
					pressure_result 			= Array("RESERVED SPACE" , _
												"pressure_res_6_1" ,  _
												"pressure_res_6_2" ,  _
												"pressure_res_6_3" ,  _
												"pressure_res_6_4" ,  _
												"pressure_res_6_5" ,  _
												"pressure_res_6_6")
					Analysis_DB_Tag = pressure_result(index)

		ElseIf DB_name = "temperature_result" Then										
				Dim temperature_result	
					temperature_result 			= Array("RESERVED SPACE" , _
												"temperature_res_6_1" ,  _
												"temperature_res_6_2" ,  _
												"temperature_res_6_3" ,  _
												"temperature_res_6_4" ,  _
												"temperature_res_6_5" ,  _
												"temperature_res_6_6")
					Analysis_DB_Tag = temperature_result(index)
					
		ElseIf DB_name = "sticker_result" Then										
				Dim sticker_result	
					sticker_result 			= Array("RESERVED SPACE" , _
												"sticker_res_1" ,  _
												"sticker_res_2" ,  _
												"sticker_res_3" ,  _
												"sticker_res_4" ,  _
												"sticker_res_5" ,  _
												"sticker_res_6")
					Analysis_DB_Tag = sticker_result(index)
		
		ElseIf DB_name = "pressure_cyl_result" Then												
				Dim pressure_cyl_result	
					pressure_cyl_result 			= Array("RESERVED SPACE" , _
													"pressure_cyl_result_1" ,  _
													"pressure_cyl_result_2" ,  _
													"pressure_cyl_result_3" ,  _
													"pressure_cyl_result_4" ,  _
													"pressure_cyl_result_5" ,  _
													"pressure_cyl_result_6")
					Analysis_DB_Tag = pressure_cyl_result(index)
					
		Elseif DB_name = "filling_class" Then										
				Dim filling_class	
					filling_class 			= Array("RESERVED SPACE" , _
												"filling_class_1" ,  _
												"filling_class_2" ,  _
												"filling_class_3" ,  _
												"filling_class_4" ,  _
												"filling_class_5" ,  _
												"filling_class_6")
					Analysis_DB_Tag = filling_class(index)
								
		End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function Analysis_DB_Tag is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
'Description : Generate Sticker Inside DB
'---------------------------------------------------------------	
	Public Function generate_sticker(Byval stick_res)
	On Error Resume Next
	Dim date_now : date_now = Replace(FormatDateTime(date,2),"/","-")
	Dim	date_exp : date_exp = Replace(FormatDateTime(DateAdd("yyyy",1,Date),2),"/","-")
	Dim masterlist
		
		Call Mysql_Non_Query("Insert analysis_sticker_cyl Set batch_no = '"& batch_id &"' , analyze_date = '"& date_now &"' , expired_date = '"& date_exp &"' , analysing_result = '"& stick_res &"' , cyl_id = "& cyl_id &" , sticker_trigger = 2 ")
	
	Call GF_LogToFile_("Execute", "Insert Database analysis_sticker_cyl [" & date_now &"|"& date_exp&"]" ,"Analysis")
		
'Sticker is been generate inside excel
'-------------------------------------------------
		Set masterlist = New MasterList_Sticker
		Set masterlist.Cyl_Sticker = New Cyl_Sticker
		masterlist.Cyl_Sticker.Batch_Id 	= batch_id
		masterlist.Cyl_Sticker.Cyl_Id   	= cylinder_id
		masterlist.Cyl_Sticker.Generate_Sticker
	
	Call GF_LogToFile_("Execute", "Generate Excel Analysis [Cylinders]" ,"Analysis")
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function generate_sticker is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
	End Function

	Public Function generate_sticker_report()
	On Error Resume Next
	Dim stick_res, pressure
		Dim res : res = analyser_1_result & "|" & analyser_2_result & "|" & analyser_3_result & "|" & analyser_4_result & "|" & analyser_5_result & "|" & analyser_6_result
		
		If res = "PASS|PASS|PASS|PASS|PASS|PASS" Then
			stick_res = "PASS"
		Else
			stick_res = "FAIL"
		End If
		
		Dim date_now : date_now = Replace(FormatDateTime(date,2),"/","-")
		Dim	date_exp : date_exp = Replace(FormatDateTime(DateAdd("yyyy",1,Date),2),"/","-")
		
		pressure = pressure_guid_result(GUID)
		Call Mysql_Non_Query("Insert analysis_sticker_report Set batch_no = '"& batch_id &"' , analyze_date = '"& date_now &"' , expired_date = '"& date_exp &"' , analysing_result = '"& stick_res &"' , pressure_result = '"& pressure &"' , sticker_trigger = 1 ")
		
		Call GF_LogToFile_("Execute", "Insert Database analysis_sticker_report [" & batch_id &"|"& date_now & "|" & date_exp & "|" & stick_res &"|" & pressure &"]" ,"Analysis")
		
'Sticker is been generate inside excel
'------------------------------------------------
		Set masterlist = New MasterList_Sticker
		Set masterlist.Report_Sticker = New Report_Sticker
		masterlist.Report_Sticker.Batch_Id = batch_id
		masterlist.Report_Sticker.Generate_Sticker
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function masterlist.Report_Sticker.Generate_Sticker is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
		Call GF_LogToFile_("Execute", "Generate Excel Analysis [Reports]" ,"Analysis")
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function generate_sticker_report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function

'Description : Generate GUID 
'---------------------------------------------------------------	
	Public Function generate_guid()
	On Error Resume Next
	Dim TypeLib , myGuid
	
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		myGuid = TypeLib.Guid
		myGuid = Left(myGuid, Len(myGuid)-2)
		Call Mysql_Non_Query("Update analysis_rack Set GUID = '"& myGuid &"' Where rack_id = "& rack_no & "")
		Call GF_LogToFile_("Execute", "Generate New GUID : " & myGuid & "| racks: " & rack_no ,"Analysis")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function generate_guid is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
'Description : store data into DB
'---------------------------------------------------------------
	
	Public Function StoreToDB_Analysis()
	On Error Resume Next
	Dim res : res = analyser_1_result & "|" & analyser_2_result & "|" & analyser_3_result & "|" & analyser_4_result & "|" & analyser_5_result & "|" & analyser_6_result
	Dim ana_result
	
	If res = "PASS|PASS|PASS|PASS|PASS|PASS" Then
		ana_result = "PASS"
	Else
		ana_result = "FAIL"
	End If
			Call Mysql_Non_Query("Insert analysis_reports Set user_id = "& uid &", GUID = '"& GUID &"' , start_time = '"& start_time &"' , start_date = '"& start_date &"' , end_time = '"& end_time &"' , end_date = '"& end_date &"' ,end_date_guid='',end_time_guid='', filling_clasification = '"& filling_class &"',oms_batch = '"& batch_id &"' , " & _
							 "filling_batch = '????' , pressure_result = '"& pressure_cyl_result &"' , vac_pressure = '0' , fill_pressure = 0 , fill_temp = 0 , production_result = '????' , " & _
							 "cyl_prod_name = '"& prod_id &"' , cyl_id = "& cyl_id &" , " & _ 
							 "ana1 = "& analyser_1_val &" , ana2 = "& analyser_2_val &" , ana3 = "& analyser_3_val &" , ana4 = "& analyser_4_val &" , ana5 = "& analyser_5_val &" , ana6 = "& analyser_6_val &"  ," & _ 
							 "ana_1_fill_press = "& pressure_result &" , ana_2_fill_press = "& pressure_result &" , ana_3_fill_press = "& pressure_result &" , ana_4_fill_press = "& pressure_result &" , ana_5_fill_press = "& pressure_result &" , ana_6_fill_press = "& pressure_result &"  ," & _ 
							 "ana_1_fill_temp = "& temperature_result &" , ana_2_fill_temp = "& temperature_result &" , ana_3_fill_temp = "& temperature_result &" , ana_4_fill_temp = "& temperature_result &" , ana_5_fill_temp = "& temperature_result &" , ana_6_fill_temp = "& temperature_result &"  ," & _ 
							 "ana_1_result = '"& analyser_1_result &"' , ana_2_result = '"& analyser_2_result &"' , ana_3_result = '"& analyser_3_result &"' , ana_4_result = '"& analyser_4_result &"' , ana_5_result = '"& analyser_5_result &"' , ana_6_result = '"& analyser_6_result &"' , ana_result = '"& ana_result &"'")
	
	Call GF_LogToFile_("Execute", "Insert Database analysis_reports" ,"Analysis")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function StoreToDB_Analysis is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
	Public Function Update_end_time()
		On Error Resume Next
		
		Call Mysql_Non_Query("Update analysis_reports Set end_date_guid='"& date &"' , end_time_guid ='"& time &"' Where GUID = '"& GUID &"'")
		Call GF_LogToFile_("Execute", "Update end_time DB analysis_reports : " & GUID ,"Analysis")
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function Update_end_time is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
'Last Point Stop Here Ahmad
'---------------------------------------------------------------
'New Developement

	Public Function ClearAnalysisInternalTags
		On Error Resume Next
		
		start_time = ""   
		start_date = "" 
		end_time = "" 
		end_date = "" 
		analyser_1_val = "" 
		analyser_1_result = "" 	
		analyser_2_val = "" 	
		analyser_2_result = "" 	
		analyser_3_val = "" 
		analyser_3_result = "" 
		analyser_4_val = "" 
		analyser_4_result = "" 	
		analyser_5_val = "" 	
		analyser_5_result = "" 	
		analyser_6_val = "" 
		analyser_6_result = "" 
		pressure_result = "" 
		temperature_result = "" 
		sticker_result = "" 
		pressure_cyl_result = "" 
		filling_class = "" 
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Analysis.bmo - Function ClearAnalysisInternalTags is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function

End Class