Class Individual

'---------------------------------------------------------------------------------------------------------------------------
' Descriptions : Variable of Individual 
'---------------------------------------------------------------------------------------------------------------------------
	'General Variable
	Private m_GUID, m_Batch, m_Count, m_Cyl
	Private Dict_LP
	
	'Timestamp for shift
	Private morning_trigger , noon_trigger , evening_trigger , time_h
	
	'Database
	Private objConn, objRS
	Private DB, UID, PASS, SVR
	
	'Excel
	Private xlApp, xlBook, xlSht 
	Private filename, Savefile 
	Private objFSO
	
	Private path 
	Private Template 
	Private Today_date_time
	Private SaveTofile 
	
	Private DestinationFile,SourceFile,MasterFiles
'---------------------------------------------------------------------------------------------------------------------------
' Descriptions : Constructor of Individual
'---------------------------------------------------------------------------------------------------------------------------
	
	Private Sub Class_Initialize()
		On Error Resume Next
		path = "C:\SRT\REPORT\Analysis_Report_I\" & Replace(Date,"/","_") & "\"
		Template = "C:\SRT\REPORT\Analysis_Report_Master\Analysis_I_Master.xlsx"
		Today_date_time = Replace(Date,"/","") & "_" & Replace(Time,":","")
		SaveTofile = path & "Analysis_Report_I_"& Today_date_time &"_"&".xlsx"
		
		Set objConn = CreateObject("ADODB.Connection")
		Set objRS   = CreateObject("ADODB.Recordset")
		Set Dict_LP = CreateObject("Scripting.Dictionary")
		Set objFSO  = CreateObject("Scripting.FileSystemObject")
		Set xlApp = CreateObject("Excel.Application")
		
		DestinationFile = "C:\SRT\REPORT\Analysis_Report_I\Analysis_I_Master.xlsx"
        SourceFile =  Template 
        MasterFiles = "C:\SRT\REPORT\Analysis_Report_I\"
        
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
		    Call GF_LogError("Error", "Individual.bmo - Function Class_Initialize is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Sub
	
'---------------------------------------------------------------------------------------------------------------------------
' Descriptions : Set And Get 
'---------------------------------------------------------------------------------------------------------------------------

	'Batch ID variable (set:get)
	Public Property Get Batch_Id
		Batch_Id = m_Batch
	End Property

	Public Property Let Batch_Id(ByVal value)
		m_Batch = value
	End Property
	
	'GUID variable (set:get)
	Public Property Get GUID
		GUID = m_GUID
	End Property

	Public Property Let GUID(ByVal value)
		m_GUID = value
	End Property
	
'---------------------------------------------------------------------------------------------------------------------------
' Descriptions : Initial Function to Trigger Reports
'---------------------------------------------------------------------------------------------------------------------------
	
	'Generate Report 	
	Public Function Generate_Report
		On Error Resume Next
		time_h = Hour(Now())
		
		'Read shift trigger time
		morning_trigger = HMIRuntime.Tags("individual_report_morning").Read
		noon_trigger = HMIRuntime.Tags("individual_report_noon").Read
		evening_trigger = HMIRuntime.Tags("individual_report_evening").Read
		
		If morning_trigger = "" Then
			HMIRuntime.Tags("individual_report_morning").Write 0
			morning_trigger = 0
		Elseif noon_trigger = "" Then
			HMIRuntime.Tags("individual_report_noon").Write 0
			noon_trigger = 0
		Elseif evening_trigger = "" Then
			HMIRuntime.Tags("individual_report_evening").Write 0
			evening_trigger = 0
		End If
		
		'check end of shift start from midnight to morning
		If time_h = 7 Then
			HMIRuntime.Tags("individual_report_morning").Write 0
			HMIRuntime.Tags("individual_report_noon").Write 0
			If HMIRuntime.Tags("individual_report_evening").Read = 0 Then
				Call Report_trigger
				HMIRuntime.Tags("individual_report_evening").Write 1
			End If
			
		Elseif time_h = 15 Then
			HMIRuntime.Tags("individual_report_noon").Write 0
			HMIRuntime.Tags("individual_report_evening").Write 0
			If HMIRuntime.Tags("individual_report_morning").Read = 0 Then
				Call Report_trigger
				HMIRuntime.Tags("individual_report_morning").Write 1
			End If
			
			
		Elseif time_h = 22 Then
			HMIRuntime.Tags("individual_report_morning").Write 0
			HMIRuntime.Tags("individual_report_evening").Write 0
			If HMIRuntime.Tags("individual_report_noon").Read = 0 Then
				Call Report_trigger
				HMIRuntime.Tags("individual_report_noon").Write 1
			End If
			
		End If
		
		xlApp.Quit
		'always deallocate after use...
		set xlSht = Nothing 
		Set xlBook = Nothing
		Set xlApp = Nothing
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Individual.bmo - Function Generate_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
'---------------------------------------------------------------------------------------------------------------------------
' Descriptions : Functions to fetch all report information
'---------------------------------------------------------------------------------------------------------------------------
	
	Public Function Report_trigger
		On Error Resume Next
		Call Write_General_Information_I_Ana_Report
		Call Write_Cyl_ID_I_Ana_Report
		Call Write_Analysis_I_Ana_Report
		
		xlBook.SaveAs SaveTofile, 51 
		xlBook.Close SaveChanges=True
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Individual.bmo - Function Report_trigger is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
		Call Autoprint(SaveTofile, 1, 0, 1)
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Individual.bmo - Function AutoPrint is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function

'---------------------------------------------------------------------------------------------------------------------------
' Descriptions : General Functions
'---------------------------------------------------------------------------------------------------------------------------
	
	'Write general information into excel files
	Public Function Write_General_Information_I_Ana_Report
		On Error Resume Next
		Dim Start_Time_Report , End_Time_Report
		Dim Start_Date_Report , End_Date_Report
		
		If time_h = 7 Then
			Start_Time_Report = "22:00"
			Start_Date_Report = (Date() - 1)
			End_Date_Report = Date()
			End_Time_Report = "07:00"
		Elseif time_h = 15 Then
			Start_Time_Report = "07:00"
			Start_Date_Report = Date()
			End_Date_Report = Date()
			End_Time_Report = "15:00"
		Elseif time_h = 22 Then
			Start_Time_Report = "15:00"
			Start_Date_Report = Date()
			End_Time_Report = Date()
			End_Time_Report = "22:00"
		End If
									
		'Write data into the spreadsheet
		xlSht.Range("E6").Value = Start_Date_Report & ", " & Start_Time_Report  'Start time and date
		xlSht.Range("L6").Value = End_Date_Report & ", " & End_Time_Report	 'End time and date
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Individual.bmo - Function Write_General_Information_I_Ana_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
	End Function
	
	Public Function Write_Cyl_ID_I_Ana_Report
	On Error Resume Next
	Dim i : i = 17	
	Dim re : Set re = New RegExp
		re.Global = True
		re.Pattern = "[a-zA-Z_]"
	Dim long_str , mid_str , low_str, prod_recipe , prod_cyl
		
		
	
		If time_h = 7 Then
			'Set objRS = objConn.execute("Select cylinder_id,prod_id " & _
			'							" from filling_individual_complete where timestamp BETWEEN concat(SUBDATE(CURDATE(), INTERVAL 1 DAY),' 22:00:00') And NoW() ")
			
			Set objRS = objConn.execute("SELECT p.cyl_id, g.cyl_prod_name ,f.weight1 ,f.weight2 ,f.weight3 ,f.weight1_starttime , f.weight1_endtime , f.weight2_starttime ," & _
										"f.weight2_endtime , f.weight3_starttime ,f.weight3_endtime, f.final_result " & _
										"FROM pallet_table p " & _
										"	INNER JOIN gas_id_table g ON p.cyl_prod_id = g.cyl_prod_id " & _
    									"	INNER JoIN sub_fill_individual f On p.cyl_id = f.cylinder_id " & _
    									"WHERE p.id = f.pallet_table_id and f.time_created BETWEEN concat(SUBDATE(CURDATE(), INTERVAL 1 DAY),' 22:00:00') And NoW()") 
		
		Elseif time_h = 15 Then	
			
			Set objRS = objConn.execute("SELECT p.cyl_id, g.cyl_prod_name ,f.weight1 ,f.weight2 ,f.weight3 ,f.weight1_starttime , f.weight1_endtime , f.weight2_starttime ," & _
										"f.weight2_endtime , f.weight3_starttime ,f.weight3_endtime, f.final_result " & _
										"FROM pallet_table p " & _
										"	INNER JOIN gas_id_table g ON p.cyl_prod_id = g.cyl_prod_id " & _
    									"	INNER JoIN sub_fill_individual f On p.cyl_id = f.cylinder_id " & _
    									"WHERE p.id = f.pallet_table_id and f.time_created BETWEEN concat(CURDATE(),' 07:00:00') And NoW()") 
    									
			'Set objRS = objConn.execute("Select p.cyl_id,g.cyl_prod_name " & _
			'							" from pallet_table p join gas_id_table g on (g.cyl_prod_id = p.cyl_prod_id) where   and timestamp BETWEEN concat(CURDATE(),' 07:00:00') And NoW() ")
		
		Elseif time_h = 22 Then
		
			'Set objRS = objConn.execute("Select p.cyl_id,g.cyl_prod_name " & _
			'							" from pallet_table p join gas_id_table g on (g.cyl_prod_id = p.cyl_prod_id) where analysis_mode = 1 and fill_mode = 1  and timestamp BETWEEN concat(CURDATE(),' 15:00:00') And NoW() ")
			
			Set objRS = objConn.execute("SELECT p.cyl_id, g.cyl_prod_name ,f.weight1 ,f.weight2 ,f.weight3 ,f.weight1_starttime , f.weight1_endtime , f.weight2_starttime ," & _
										"f.weight2_endtime , f.weight3_starttime ,f.weight3_endtime, f.final_result " & _
										"FROM pallet_table p " & _
										"	INNER JOIN gas_id_table g ON p.cyl_prod_id = g.cyl_prod_id " & _
    									"	INNER JoIN sub_fill_individual f On p.cyl_id = f.cylinder_id " & _
    									"WHERE p.id = f.pallet_table_id and f.time_created BETWEEN concat(CURDATE(),' 15:00:00') And NoW()")
    									
		End If
			
			
			If objRS.EOF  Then
			Else
				Do While Not objRS.EOF
				
					'low_str = Split(objRS(1).value, "-")
					'del_3 = ubound(low_str)
		
					'For j = 0 To del_3
					'	If del_3 = 2 Then
					'		prod_recipe = low_str(0) 
					'		prod_recipe = re.Replace(prod_recipe,"")
					'		prod_cyl = low_str(2) & "-" & low_str(1) 
					'	Else
					'		prod_recipe = low_str(0)
					'		prod_recipe = re.Replace(prod_recipe,"")
					'		prod_cyl = low_str(1)	
					'	End If	
					'Next
				
				   xlSht.Range("D"&(i)).Value = objRS(0).value 
				   xlSht.Range("B"&(i)).Value = objRS(1).value 
				   
				   xlSht.Range("F"&(i)).Value = objRS(2).value
				   xlSht.Range("G"&(i)).Value = objRS(5).value
				   xlSht.Range("H"&(i)).Value = objRS(6).value
				   
				   xlSht.Range("I"&(i)).Value = objRS(3).value
				   xlSht.Range("J"&(i)).Value = objRS(7).value
				   xlSht.Range("K"&(i)).Value = objRS(8).value
				   
				   xlSht.Range("L"&(i)).Value = objRS(4).value
				   xlSht.Range("M"&(i)).Value = objRS(9).value
				   xlSht.Range("P"&(i)).Value = objRS(10).value
				   
				   If objRS(11).value = 0 Then
						xlSht.Range("AB"&(i)).Value = "PASS"
				   Else
						 xlSht.Range("AB"&(i)).Value = "FAIL"
				   End If
						 
					'If prod_recipe = "211" Then
					'	 xlSht.Range("I"&(i)).Value = objRS(2).value
					'	 xlSht.Range("J"&(i)).Value = objRS(4).value
					'	 xlSht.Range("K"&(i)).Value = objRS(5).value
					'	 xlSht.Range("L"&(i)).Value = objRS(3).value
					'	 xlSht.Range("M"&(i)).Value = objRS(6).value
					'	 xlSht.Range("P"&(i)).Value = objRS(7).value
					'	 If objRS(8).value = 0 Then
					'	 	xlSht.Range("AB"&(i)).Value = "PASS"
					'	 Else
					'	 	xlSht.Range("AB"&(i)).Value = "FAIL"
					'	 End If
						 
					'Else
					'	 xlSht.Range("F"&(i)).Value = objRS(2).value
					'	 xlSht.Range("G"&(i)).Value = objRS(4).value
					'	 xlSht.Range("H"&(i)).Value = objRS(5).value
					'	 If objRS(8).value = 0 Then
					'	 	xlSht.Range("AB"&(i)).Value = "PASS"
					'	 Else
					'	 	xlSht.Range("AB"&(i)).Value = "FAIL"
					'	 End If
					'End If
			
				   Dict_LP.Add objRS(0).value,i
				   i = i + 1
				   objRS.MoveNext    
				Loop
			End If
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Individual.bmo - Function Write_Cyl_ID_I_Ana_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
	
	Public Function Write_Analysis_I_Ana_Report
	On Error Resume Next
	Dim position,index_excel,res
		
		If time_h = 7 Then
			Set objRS = objConn.execute("Select cyl_id,ana1,ana2,ana3,ana4,ana5,ana_6_fill_press,ana_result " & _
									" from analysis_reports where timestamp BETWEEN concat(SUBDATE(CURDATE(), INTERVAL 1 DAY),' 22:00:00') And NoW()  ")
		Elseif time_h = 15 Then
			Set objRS = objConn.execute("Select cyl_id,ana1,ana2,ana3,ana4,ana5,ana_6_fill_press,ana_result " & _
									" from analysis_reports where timestamp BETWEEN concat(CURDATE(),' 07:00:00') And NoW()  ")
		Elseif time_h = 22 Then
			Set objRS = objConn.execute("Select cyl_id,ana1,ana2,ana3,ana4,ana5,ana_6_fill_press,ana_result " & _
									" from analysis_reports where timestamp BETWEEN concat(CURDATE(),' 15:00:00') And NoW()  ")
		End If
		
		
		If objRS.EOF  Then
		Else
			Do While Not objRS.EOF
			If Dict_LP.Exists(objRS(0).value) Then
				position = Dict_LP.Item(objRS(0).value)
				
					xlSht.Range("S"& position).Value = objRS(1).value  
					xlSht.Range("T"& position).Value = objRS(2).value  
					xlSht.Range("V"& position).Value = objRS(3).value  
					xlSht.Range("X"& position).Value = objRS(4).value 
					xlSht.Range("Z"& position).Value = objRS(5).value
					xlSht.Range("AC"& position).Value = objRS(7).value    		
					  
			End If
			objRS.MoveNext    
			Loop
		End If
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Individual.bmo - Function Write_Analysis_I_Ana_Report is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
	
End Class