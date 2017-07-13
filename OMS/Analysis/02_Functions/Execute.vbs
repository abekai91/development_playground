' Description : Execute Checkpoint 1
'--------------------------------------------------------------------------
Function execute_cp_1(par1)
On Error Resume Next
Dim fill_status : fill_status = par1.fill_oms
Dim check_start : check_start = par1.start
	 
	 If fill_status = 1 Then
		par1.fill_by_oms
	 End If
	 
	 If check_start = 1 Then
		par1.deactivate_analysis_1
	 End If
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function execute_cp_1 is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
End Function

' Description : Execute Checkpoint 2
'--------------------------------------------------------------------------
Function execute_cp_2(par1)
On Error Resume Next
Dim Check_Capture : Check_Capture = par1.capture
Dim exit_cp2 	  : exit_cp2 = 0
Dim Info_Msg : Info_Msg = " : Analysis CP2 ("& par1.rack_name_get & "|" & par1.prod_id_get & "|" & Check_Capture & ")"
	 
	Select Case CINT(Check_Capture)
		   Case 0
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Analysis")
		   Case 1
				HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
				par1.analysis_pressure_temperature
				par1.analysis_rules
		   		exit_cp2 = 1 
		   		Call GF_LogToFile_("Execute", Info_Msg ,"Analysis")	
		   Case Else
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Analysis")
	End Select
	
	If exit_cp2 = 1 Then
		par1.deactivate_analysis_2
	 End If
	
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function execute_cp_2 is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
End Function


' Description : Execute Checkpoint 3
'--------------------------------------------------------------------------
Function execute_cp_3(par1)
On Error Resume Next
Dim Check_Sticker : Check_Sticker = par1.print_st
Dim exit_cp3 	  : exit_cp3 = 0	
Dim Info_Msg : Info_Msg = " : Analysis CP3 : Checking Sticker (" & Check_Sticker & ")"
 
	Select Case CINT(Check_Sticker)
		   Case 0
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Analysis")
		   Case 1
				HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	par1.print_sticker
				par1.StoreToDB_Analysis
		   		exit_cp3 = 1 
		   		Call GF_LogToFile_("Execute", Info_Msg ,"Analysis")
		   Case Else
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Analysis")
	End Select
	
	If exit_cp3 = 1 Then
		par1.deactivate_analysis_3
	 End If
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function execute_cp_3 is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
End Function


' Description : Execute Checkpoint 4
'--------------------------------------------------------------------------
Function execute_cp_4(par1)
On Error Resume Next
Dim Check_Complete : Check_Complete = par1.complete
Dim exit_cp4 	  : exit_cp4 = 0	
	
	Select Case CINT(Check_Complete)
		   Case 0
			 	HMIRuntime.Trace(Now & " : Analysis CP4 : New Scan Start" & vbCrlf )
			 	Call GF_LogToFile_("Execute", " : Analysis CP4 : New Scan Start" ,"Analysis")
		   Case 1
				HMIRuntime.Trace(Now & " : Analysis CP4 : Checking Sticker (1)" & vbCrlf )
				Call GF_LogToFile_("Execute", " : Analysis CP4 : Checking Sticker (1)" ,"Analysis")
				par1.Update_end_time
				par1.generate_sticker_report
				par1.generate_reports
			 	par1.generate_guid
			 	par1.ClearAnalysisInternalTags
		   		exit_cp4 = 1 
	End Select
	
	If exit_cp4 = 1 Then
		par1.deactivate_analysis_4
	 End If
	
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function execute_cp_4 is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
End Function