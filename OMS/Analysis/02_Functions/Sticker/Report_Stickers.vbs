Class Report_Sticker
	'General Variable
	Private m_Batch
	
	'Database
	Private objConn,objRS
	Private DB, UID, PASS, SVR
	
	'Excel
	Private xlApp, xlBook, xlSht 
	Private filename, Savefile 
	Private objFSO
	
	Private path 
	Private Template 
	Private Today_date_time
	Private SaveTofile
	Private Cylinder_Count
	
	Private DestinationFile,SourceFile,MasterFiles
	
	Private Sub Class_Initialize()
		On Error Resume Next
		path = "C:\SRT\STICKER\REPORT\STICKER\" & Replace(Date,"/","_") & "\"
		Template = "C:\SRT\STICKER\REPORT\MASTER\REPORT_STICKER.xls"
		Today_date_time = Replace(Date,"/","") & "_" & Replace(Time,":","")
		SaveTofile = path & "Report_Sticker_"& Today_date_time &"_"&".xlsx"
		
		Set objConn = CreateObject("ADODB.Connection")
		Set objRS   = CreateObject("ADODB.Recordset")
		Set Dict_LP = CreateObject("Scripting.Dictionary")
		Set objFSO  = CreateObject("Scripting.FileSystemObject")
		
		DestinationFile = "C:\SRT\STICKER\REPORT\STICKER\REPORT_STICKER.xls"
        SourceFile =  Template 
        MasterFiles = "C:\SRT\STICKER\REPORT\STICKER\"
        
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
		
		Set xlApp = CreateObject("Excel.Application")
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
		
		Call GF_LogToFile_("Checking", "Function Class_Initialize Is Workings","Analysis")
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Report_Sticker.bmo - Function Class_Initialize is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Sub
	
	Public Property Get Batch_Id
		Batch_Id = m_Batch
	End Property

	Public Property Let Batch_Id(ByVal value)
		m_Batch = value
	End Property
	
	Public Function Generate_Sticker
		On Error Resume Next
		Call Write_Cylinder
		
		xlBook.SaveAs SaveTofile, 51 
		xlBook.Close SaveChanges=True
		xlApp.Quit
		
		'always deallocate after use...
		set xlSht = Nothing 
		Set xlBook = Nothing
		Set xlApp = Nothing
		
		Call GF_LogToFile_("Checking", "Function Generate Sticker Is Workings (Create excel)","Analysis")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Report_Sticker.bmo - Function Generate_Sticker is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
		
		Call Autoprint(SaveTofile, 0, 1, Cylinder_Count)
		Call GF_LogToFile_("Checking", "Function AutoPrint Is Workings","Analysis")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Report_Sticker.bmo - AutoPrint Problems","Analysis")
		    Err.Clear
		End If
	End Function
	
	Public Function Write_Cylinder
		On Error Resume Next
		Dim OMS_Batch,Analyzer_Date,Analyzer_Date_Exp,Ana_Result,Press_Result
		
		'Check table analysis	
		Set objRS = objConn.execute("Select batch_no,analyze_date,expired_date,analysing_result,pressure_result " & _
									" from analysis_sticker_report where batch_no = '" & m_Batch & "' order by timestamp desc LIMIT 1")
		If objRS.EOF  Then
			OMS_Batch = "UNKNOWN"
			Analyzer_Date = "-/-/-"
			Analyzer_Date_Exp = "-/-/-"
			Ana_Result = "FAIL"
			Press_Result = "FAIL"
			Cylinder_Count = 1				
		Else
			OMS_Batch = objRS(0).value
			Analyzer_Date = objRS(1).value
			Analyzer_Date_Exp = objRS(2).value
			Ana_Result = objRS(3).value
			Press_Result = objRS(4).value

			Set objRS = objConn.execute("Select count(id) from pallet_table where oms_batch = '" & m_Batch & "' And cyl_state=1 LIMIT 1")
				
			If objRS.EOF Then
				Cylinder_Count = 1
			Else
				Cylinder_Count = CInt(objRS(0).value) + 1
			End If
		End If
		
		'Write data into the spreadsheet
		xlSht.Range("C10").Value = OMS_Batch  		 'Batch ID
		xlSht.Range("C11").Value = Analyzer_Date	 'Analyze Date
		xlSht.Range("C12").Value = Analyzer_Date_Exp 'Analyze Date Exp
		xlSht.Range("C13").Value = Ana_Result	 	 'Analysis Results
		xlSht.Range("C14").Value = Press_Result	     'Pressure Results
		
		Call GF_LogToFile_("Checking", "Function Write Cylinder Is Workings (" & Cyl_ID & ")","Analysis")
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Report_Sticker.bmo - Function Write_Cylinder is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Function
End Class