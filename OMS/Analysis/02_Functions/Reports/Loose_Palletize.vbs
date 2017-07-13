Class Loose_Palletize
	'General Variable
	Private m_GUID, m_Batch, m_Count
	Private Dict_LP
	'Database
	Private objConn,objRS
	Private DB, UID, PASS, SVR
	'Excel
	Private xlApp, xlBook, xlSht 
	Private filename, Savefile 
	Private objFSO, path, Template, Today_date_time, SaveTofile 
	Private classification,DestinationFile,SourceFile,MasterFiles
	'Refrence Key
	Private prod_name
	
	
'**********************************************************************************************************************************************************	
' Description  : Class Initialize
' Types        : Function
'**********************************************************************************************************************************************************			
	Private Sub Class_Initialize()
		On Error Resume Next
		path            = "C:\SRT\REPORT\Analysis_Report_PL\" & Replace(Date,"/","_") & "\"
		Template        = "C:\SRT\REPORT\Analysis_Report_Master\Analysis_PL_Master.xlsx"
		Today_date_time = Right("0" & Hour(Time),2) & Right("0" & Minute(Time),2) & Right("0" & Second(Time),2) 'Replace(Date,"/","") & "_" & Replace(Time,":","")
		SaveTofile      = path & "Linde_Analysis_PL_"& Today_date_time &"_"& classification &".xlsx"
		DestinationFile = "C:\SRT\REPORT\Analysis_Report_PL\Analysis_PL_Master.xlsx"
        SourceFile      =  Template 
        MasterFiles     = "C:\SRT\REPORT\Analysis_Report_PL\"
		
		Set objConn = CreateObject("ADODB.Connection")
		Set objRS   = CreateObject("ADODB.Recordset")
		Set Dict_LP = CreateObject("Scripting.Dictionary")
		Set objFSO  = CreateObject("Scripting.FileSystemObject")
		Set xlApp   = CreateObject("Excel.Application")
		
        If objFSO.FileExists(DestinationFile) Then
	        If Not objFSO.GetFile(DestinationFile).Attributes And 1 Then 
	            objFSO.CopyFile SourceFile, MasterFiles, True
	        Else 
	            objFSO.GetFile(DestinationFile).Attributes = objFSO.GetFile(DestinationFile).Attributes - 1
	            objFSO.CopyFile SourceFile, MasterFiles, True
	            objFSO.GetFile(DestinationFile).Attributes = objFSO.GetFile(DestinationFile).Attributes + 1
	        End If
        Else
           objFSO.CopyFile SourceFile, MasterFiles, True
        End If
        
		set xlBook = xlApp.WorkBooks.Open(DestinationFile)
		Set xlSht = xlApp.activesheet
		
		SVR  = HMIRuntime.Tags("Server").Read
		DB   = HMIRuntime.Tags("Database").Read
		UID  = HMIRuntime.Tags("UID").Read
		PASS = HMIRuntime.Tags("PASS").Read
	
		Call Mysql_Open_Conn(objConn,objRS,SVR,DB,UID,PASS)
		
		If Not objFSO.FolderExists(path) Then
		   objFSO.CreateFolder(path)
		End If
		
		xlApp.DisplayAlerts = False
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Loose_Palletize.bmo - Function Class_Initialize is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Sub
	
'**********************************************************************************************************************************************************	
' Description  : Declare property
' Types        : Object property
'**********************************************************************************************************************************************************		
	Public Property Get Batch_Id
		Batch_Id = m_Batch
	End Property

	Public Property Let Batch_Id(ByVal value)
		m_Batch = value
	End Property
	
	Public Property Get report_type
		report_type = classification
	End Property

	Public Property Let report_type(ByVal value)
		classification = value
	End Property
	
		
	Public Property Get GUID
		GUID = m_GUID
	End Property

	Public Property Let GUID(ByVal value)
		m_GUID = value
	End Property
	
'**********************************************************************************************************************************************************	
' Description  : Interface/flows of Class
' Types        : Function
'**********************************************************************************************************************************************************			
	Public Function Generate_Report
		On Error Resume Next
		Call Write_General_Information_LP_Ana_Report
		Call Write_Cyl_ID_LP_Ana_Report
		Call Write_Analysis_LP_Ana_Report
		Call Write_Analysis_LP_AnalyzerName
		Call Write_Analysis_LP_AnalyzerPassFail
		
		xlBook.SaveAs SaveTofile, 51 
		xlBook.Close SaveChanges=True
		xlApp.Quit
		
		'always deallocate after use...
		set xlSht = Nothing 
		Set xlBook = Nothing
		Set xlApp = Nothing
		
		Call Autoprint(SaveTofile, 1, 0, 1)
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Loose_Palletize.bmo - Function Generate_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
'**********************************************************************************************************************************************************	
' Description  : Writing General information into Excel Columns
' Types        : Function
'**********************************************************************************************************************************************************	
	Public Function Write_General_Information_LP_Ana_Report 
		On Error Resume Next
		Dim Filler,QC
		Dim Filler_Mula_Date , Filler_Mula_Time
		Dim Filler_Henti_Date, Filler_Henti_Time
		Dim QC_Mula_Date , QC_Mula_Time
		Dim QC_Henti_Date, QC_Henti_Time
		Dim Filler_Loc , QC_Loc
		Dim OMS_Batch , CRYOSTART_Batch , ANA_Results
		Dim Category, Vacc_P, Fill_weigth , Fill_Press , Fill_Temp , Prod_Res
		Dim Total_Cyl
		Dim filling_clasification
		
'Description : Get the QC operator name , start datetime & end datetime
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'Comment : remove GUID because no longer used for refrences !
		'Set objRS = objConn.execute("Select user_id,start_time,start_date,end_date_guid,end_time_guid ,filling_clasification " & _
		'							" from analysis_reports where oms_batch = '" & m_Batch & "' and GUID = '"& m_GUID &"' and filling_clasification in (2,3,4) Order by time(start_time) ASC LIMIT 1")
		Set objRS = objConn.execute("Select user_id,start_time,start_date,end_date_guid,end_time_guid ,filling_clasification " & _
									" from analysis_reports where oms_batch = '" & m_Batch & "' and filling_clasification in (2,3,4) Order by time(start_time) ASC LIMIT 1")
		If objRS.EOF  Then
				QC = 0
				QC_Mula_Time = "--:--"
				QC_Mula_Date = "-/-/-"
				QC_Henti_Date = "-/-/-"
				QC_Henti_Time = "--:--"
		Else
				QC = objRS(0).value
				QC_Mula_Time = objRS(1).value
				QC_Mula_Date = objRS(2).value
				QC_Henti_Date = objRS(3).value
				QC_Henti_Time = objRS(4).value
				filling_clasification  = objRS(5).value
		End If
	
'Description : Get the Filling Location rack name
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Set objRS = objConn.execute("SELECT GROUP_CONCAT(rack_name SEPARATOR ', ') FROM `analysis_filling_history` WHERE oms_batch = '" & m_Batch &"' and status = 2 GROUP BY oms_batch")
		If objRS.EOF  Then
			Filler_Loc  = "Unknown"
		Else
			Filler_Loc  = objRS(0).value
		End If

'Description : Get the QC Location rack name
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Set objRS = objConn.execute("SELECT GROUP_CONCAT(rack_name SEPARATOR ', ') FROM `analysis_qc_history` WHERE oms_batch = '" & m_Batch &"' and status in (1,2) GROUP BY oms_batch")
		If objRS.EOF  Then
			Filler_Loc  = ""
		Else
			Filler_Loc  = objRS(0).value
		End If
		
'Description : Get the Filler user id , Filling Time , Filling Pressure , Filing Vacuum
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Select Case filling_clasification
		   Case 2
		        report_type = "Loose"
			 	Set objRS = objConn.execute("Select user_id, user_entry_start_date, user_entry_end_date,vacuum_pressure,filling_pressure1,	filling_temperature1,result " & _
				" from sub_fill_loose where oms_batch = '" & m_Batch & "' Order by user_entry_start_date ASC LIMIT 1")
		   Case 3
		        report_type = "Mcp"
			 	Set objRS = objConn.execute("Select user_id, user_entry_start_date, user_entry_end_date,vacuum_pressure,filling_pressure1,	filling_temperature1,result " & _
				" from sub_fill_mcp where oms_batch = '" & m_Batch & "' Order by user_entry_start_date ASC LIMIT 1")
		   Case 4
		        report_type = "Cdp"
			 	Set objRS = objConn.execute("Select user_id, user_entry_start_date, user_entry_end_date,vacuum_pressure,filling_pressure1,	filling_temperature1,result " & _
				" from sub_fill_cdp where oms_batch = '" & m_Batch & "' Order by user_entry_start_date ASC LIMIT 1")
		   case else
			 HMIRuntime.Trace(Now & " Generate Report Filling Mode(" & filling_clasification & ") Not Found!"  & vbCrlf)
		End select
		
		If objRS.EOF  Then
				Filler 			 	= 0
				Filler_Mula_Time 	= "-/-/- " & " --:-- " 
				Filler_Henti_Time 	= "-/-/- " & " --:--"
				Vacc_P 				= "UNKNOWN"
				Fill_Temp 			= "UNKNOWN"
				Prod_Res			= "UNKNOWN"
		Else
				Filler 			 	= objRS(0).value
				Filler_Mula_Time 	= objRS(1).value
				Filler_Henti_Time 	= objRS(2).value
				Vacc_P 				= objRS(3).value
				Fill_Press 			= objRS(4).value
				Fill_Temp 			= objRS(5).value
			
			If objRS(6).value = 0 Then
					Prod_Res = "PASS"
			Elseif  objRS(6).value = 1 Then
					Prod_Res = "FAIL"
			Else
					Prod_Res = "UNKNOWN"
			End If
		
		End If
		
'Description : Get the Username for QC
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Set objRS = objConn.execute("SELECT GROUP_CONCAT(user_name SEPARATOR ', ') FROM `analysis_qc_history` WHERE oms_batch = '" & m_Batch &"' and status in (1,2) GROUP BY oms_batch")
		'Set objRS = objConn.execute("Select user_name " & _
		'							" from operator_table where user_id = " & QC & " ")
		If objRS.EOF  Then
				QC = "UNKNOWN"
		Else
				QC = objRS(0).value
		End If
								
'Description : Get the Username for Filler
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Set objRS = objConn.execute("SELECT GROUP_CONCAT(user_name SEPARATOR ', ') FROM `analysis_filling_history` WHERE oms_batch = '" & m_Batch &"' and status = 2 GROUP BY oms_batch")
		'Set objRS = objConn.execute("Select user_name " & _
		'							" from operator_table where user_id = " & Filler & " ")							
		If objRS.EOF  Then
			Filler = "UNKNOWN"
		Else
			Filler = objRS(0).value
		End If
		
'Description : Get the total cylinder quantity for 1 batch
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Set objRS = objConn.execute("Select count(*) " & _
									" from pallet_table where oms_batch = '" & m_Batch & "' ")							
		
		If objRS.EOF  Then
			Total_Cyl = "UNKNOWN"
		Else
			Total_Cyl = objRS(0).value
		End If		
		
'Description : Get the analysis result pass/fail
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'Comment : remove guid because no longer used !
		'Set objRS = objConn.execute("Select ana_result " & _
		'							" from analysis_reports where oms_batch = '" & m_Batch & "' and GUID = '"& m_GUID &"' and filling_clasification in (2,3,4) Order by time(start_time) ASC LIMIT 1")
		
		Set objRS = objConn.execute("Select ana_result " & _
									" from analysis_reports where oms_batch = '" & m_Batch & "' and filling_clasification in (2,3,4) Order by time(start_time) ASC LIMIT 1")
		
		If objRS.EOF  Then
				ANA_Results = "UNKNOWN"
		Else
			Do While Not objRS.EOF
				If objRS(0).value = "FAIL" Then
					ANA_Results = "FAIL"
					Exit Do
				End If
				ANA_Results = objRS(0).value
			objRS.MoveNext    
			Loop
		End If
									
'Description : Write data into the spreadsheet
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		xlSht.Range("E4").Value = Filler  'Filler
		xlSht.Range("E5").Value = QC	 'QC
		
		xlSht.Range("F7").Value = Filler_Mula_Time   'Filler Mula
		xlSht.Range("F8").Value = QC_Mula_Date & ", " & QC_Mula_Time   'QC Mula
		
		xlSht.Range("K7").Value = Filler_Henti_Time  	'Filler Henti
		xlSht.Range("K8").Value = QC_Henti_Date & ", " & QC_Henti_Time	'QC Henti
		
		xlSht.Range("E9").Value = Filler_Loc 	'Filler Location
		xlSht.Range("E10").Value = QC_Loc	 'QC Location
		
		xlSht.Range("D13").Value = m_Batch   		'OMS Batch#
		xlSht.Range("D14").Value = "-"	 	 		'CRYOSTAR BATCH#
		xlSht.Range("D15").Value = ANA_Results	 	'ANALYSIS RESULTS
		
		xlSht.Range("Q13").Value = "-"   		'CATEGORY
		xlSht.Range("Q14").Value = Vacc_P	 	    'VACUUM PRESSURE
		xlSht.Range("Q15").Value = "-"	 	'FILL WEIGHT
		xlSht.Range("Y13").Value = Fill_Press	    'FILL PRESSURE
		xlSht.Range("Y14").Value = Fill_Temp	 	'FILL TEMPERATURE
		xlSht.Range("Y15").Value = Prod_Res	 	    'PRODUCTION RESULT
		
		xlSht.Range("N16").Value = Total_Cyl	 	 'TOTAL CYLINDER
		xlSht.Range("A50").Value = "GUID : " & m_GUID	 'REPORT GUID
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Loose_Palletize.bmo - Function Write_General_Information_LP_Ana_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function



'**********************************************************************************************************************************************************	
' Description  : Writing Cylinder & Product into Excel Columns
' Types        : Function
'**********************************************************************************************************************************************************	
	Public Function Write_Cyl_ID_LP_Ana_Report
	On Error Resume Next
'Description : Write Cylinder ID into Excel Columns
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Dim i : i = 1	
		Set objRS = objConn.execute("Select cyl_id " & _
									" from pallet_table where oms_batch = '" & m_Batch & "' order by cyl_id asc")
		
		If objRS.EOF  Then
		Else
			Do While Not objRS.EOF
			  
			If i < 25 Then '1-24
				
				xlSht.Range("D"&(i + 23)).Value = objRS(0).value 'CYL_Left
				Dict_LP.Add objRS(0).value,"L-"&(i + 23)
				
			Else '25-48
				
				xlSht.Range("Q"&(i - 1)).Value = objRS(0).value 'CYL_Right
				Dict_LP.Add objRS(0).value,"R-"&(i - 1)
				
			End If
			
			i = i + 1
			objRS.MoveNext    
			Loop
		End If
		
'Description : Write product into Excel Columns
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------		
		i = 1
		Set objRS = objConn.execute("Select g.cyl_prod_name " & _
									" FROM pallet_table p JOIN gas_id_table g on (g.cyl_prod_id = p.cyl_prod_id) WHERE oms_batch = '" & m_Batch & "' order by p.cyl_id asc")
		If objRS.EOF  Then
		
		Else
			Do While Not objRS.EOF
			If i < 25 Then '1-24
			xlSht.Range("B"&(i + 23) ).Value = objRS(0).value   	'PROD_Left 
			Else
			
			xlSht.Range("O"&(i - 1)).Value = objRS(0).value   		'PROD_Right
			End If
			
			prod_name = objRS(0).value 
			
			i = i + 1
			objRS.MoveNext    
			Loop
		End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Loose_Palletize.bmo - Function Write_Cyl_ID_LP_Ana_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
	End Function

'**********************************************************************************************************************************************************	
' Description  : Writing General information into Excel Columns
' Types        : Function
'**********************************************************************************************************************************************************		
	Public Function Write_Analysis_LP_Ana_Report
	On Error Resume Next
	Dim position,index_excel,res
		'Comment : remove guid because no longer used !
		'Set objRS = objConn.execute("Select a.cyl_id,a.ana1,a.ana2,a.ana3,a.ana4,a.ana5,a.ana_6_fill_press,a.ana_result,p.post_fill_state " & _
		'							" from analysis_reports a join pallet_table p on (a.cyl_id = p.cyl_id) where a.oms_batch = '" & m_Batch & "' and a.GUID = '"& m_GUID &"' ")
		Set objRS = objConn.execute("Select a.cyl_id,a.ana1,a.ana2,a.ana3,a.ana4,a.ana5,a.ana_6_fill_press,a.ana_result,p.post_fill_state " & _
									" from analysis_reports a join pallet_table p on (a.cyl_id = p.cyl_id) where a.oms_batch = '" & m_Batch & "' ")
		
		If objRS.EOF  Then
		Else
			Do While Not objRS.EOF
			If Dict_LP.Exists(objRS(0).value) Then
				position = Dict_LP.Item(objRS(0).value)
				index_excel = Right(position, Len(position) - 2)
				position = (Left(position,1))
				
				
				If position = "L" Then
					xlSht.Range("F"& index_excel).Value = objRS(1).value  
					xlSht.Range("G"& index_excel).Value = objRS(2).value  
					xlSht.Range("H"& index_excel).Value = objRS(3).value  
					xlSht.Range("I"& index_excel).Value = objRS(4).value 
					xlSht.Range("J"& index_excel).Value = objRS(5).value  
					xlSht.Range("K"& index_excel).Value = objRS(6).value 
					xlSht.Range("L"& index_excel).Value = objRS(8).value '"-" 
					xlSht.Range("M"& index_excel).Value = objRS(7).value    		
					   		

				Elseif position = "R" Then
					xlSht.Range("T"& index_excel).Value = objRS(1).value
					xlSht.Range("U"& index_excel).Value = objRS(2).value
					xlSht.Range("W"& index_excel).Value = objRS(3).value
					xlSht.Range("Y"& index_excel).Value = objRS(4).value
					xlSht.Range("AA"& index_excel).Value = objRS(5).value
					xlSht.Range("AC"& index_excel).Value = objRS(6).value
					xlSht.Range("AD"& index_excel).Value = objRS(8).value '"-"
					xlSht.Range("AE"& index_excel).Value = objRS(7).value
				End If
				
			End If
			objRS.MoveNext    
			Loop
		End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Loose_Palletize.bmo - Function Write_Analysis_LP_Ana_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
	End Function
	
	
'**********************************************************************************************************************************************************	
' Description  : Writing Analyzer Name into Excel Columns
' Types        : Function
'**********************************************************************************************************************************************************		
	Public Function Write_Analysis_LP_AnalyzerName
	On Error Resume Next
		
		Set objRS = objConn.execute("Select analyzer_1,analyzer_2, " & _
									" analyzer_3,analyzer_4, " & _
									" analyzer_5,analyzer_6, " & _
									" analyzer_1_unit,analyzer_2_unit, " & _
									" analyzer_3_unit,analyzer_4_unit, " & _
									" analyzer_5_unit,analyzer_6_unit " & _
									" from masterlist_medical where product_name = '" & prod_name & "'")
		If objRS.EOF  Then
		Else
			Do While Not objRS.EOF
					'Left Side
					xlSht.Range("F"& index_excel).Value = objRS(0).value  & "(" & objRS(6).value & ")"
					xlSht.Range("G"& index_excel).Value = objRS(1).value  & "(" & objRS(7).value & ")"
					xlSht.Range("H"& index_excel).Value = objRS(2).value  & "(" & objRS(8).value & ")"
					xlSht.Range("I"& index_excel).Value = objRS(3).value  & "(" & objRS(9).value & ")"
					xlSht.Range("J"& index_excel).Value = objRS(4).value  & "(" & objRS(10).value & ")"
					xlSht.Range("K"& index_excel).Value = objRS(5).value  & "(" & objRS(11).value & ")"		
					'Right Side   		
					xlSht.Range("T"& index_excel).Value = objRS(0).value  & "(" & objRS(6).value & ")"
					xlSht.Range("U"& index_excel).Value = objRS(1).value  & "(" & objRS(7).value & ")"
					xlSht.Range("W"& index_excel).Value = objRS(2).value  & "(" & objRS(8).value & ")"
					xlSht.Range("Y"& index_excel).Value = objRS(3).value  & "(" & objRS(9).value & ")"
					xlSht.Range("AA"& index_excel).Value = objRS(4).value & "(" & objRS(10).value & ")"
					xlSht.Range("AC"& index_excel).Value = objRS(5).value & "(" & objRS(11).value & ")"
				
			objRS.MoveNext    
			Loop
		End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Loose_Palletize.bmo - Function Write_Analysis_LP_AnalyzerName is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
	End Function
	
'**********************************************************************************************************************************************************	
' Description  : Writing Analyzer Result into Excel Columns (for marking)
' Types        : Function
'**********************************************************************************************************************************************************		
	Public Function Write_Analysis_LP_AnalyzerPassFail
	On Error Resume Next
	Dim position,index_excel,res
		Set objRS = objConn.execute("Select a.cyl_id,a.ana_1_result,a.ana_2_result,a.ana_3_result,a.ana_4_result,a.ana_5_result,a.ana_6_result " & _
									" from analysis_reports a join pallet_table p on (a.cyl_id = p.cyl_id) where a.oms_batch = '" & m_Batch & "' ")
		
		If objRS.EOF  Then
		Else
			Do While Not objRS.EOF
			If Dict_LP.Exists(objRS(0).value) Then
				position = Dict_LP.Item(objRS(0).value)
				index_excel = Right(position, Len(position) - 2)
				position = (Left(position,1))
				
				
				If position = "L" Then
					xlSht.Range("AF"& index_excel).Value = objRS(1).value  
					xlSht.Range("AG"& index_excel).Value = objRS(2).value  
					xlSht.Range("AH"& index_excel).Value = objRS(3).value  
					xlSht.Range("AI"& index_excel).Value = objRS(4).value 
					xlSht.Range("AJ"& index_excel).Value = objRS(5).value  
					xlSht.Range("AK"& index_excel).Value = objRS(6).value   		
					   		

				Elseif position = "R" Then
					xlSht.Range("AN"& index_excel).Value = objRS(1).value
					xlSht.Range("AO"& index_excel).Value = objRS(2).value
					xlSht.Range("AP"& index_excel).Value = objRS(3).value
					xlSht.Range("AQ"& index_excel).Value = objRS(4).value
					xlSht.Range("AR"& index_excel).Value = objRS(5).value
					xlSht.Range("AS"& index_excel).Value = objRS(6).value
				End If
				
			End If
			objRS.MoveNext    
			Loop
		End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Loose_Palletize.bmo - Function Write_Analysis_LP_AnalyzerPassFail is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
	End Function
End Class