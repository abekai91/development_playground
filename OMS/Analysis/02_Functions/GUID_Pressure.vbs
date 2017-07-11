' Description : 
'--------------------------------------------------------------------------
Function pressure_guid_result(getGuid)
	On Error Resume Next
	Dim objConn,objRS
	Dim SVR,DB,UID,PASS
	Dim ret
	
		Set objConn = CreateObject("ADODB.Connection")
		Set objRS   = CreateObject("ADODB.Recordset")
		
		SVR  = HMIRuntime.Tags("Server").Read
		DB   = HMIRuntime.Tags("Database").Read
		UID  = HMIRuntime.Tags("UID").Read
		PASS = HMIRuntime.Tags("PASS").Read
	
		Call Mysql_Open_Conn(objConn,objRS,SVR,DB,UID,PASS)
		
		Set objRS = objConn.execute("Select pressure_result " & _
									" from analysis_reports where GUID = '" & getGuid & "'")
		
		Do While Not objRS.EOF
			 If objRS(0).value = "FAIL" Then
			 	ret = "FAIL"
			    Exit Do   
			  End If
			
			ret = "PASS"
			objRS.MoveNext	
		Loop
	
	pressure_guid_result = ret
	
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "GUID_Pressure.bmo - Function pressure_guid_result is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
End Function