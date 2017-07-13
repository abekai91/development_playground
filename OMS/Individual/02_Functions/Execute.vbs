'Checkpoint 1 : EnerTech ************************************************************
Function filling_individual_exec_enertech_1(par1)
On Error Resume Next
	Dim Info_Msg , Check_Status , Filling_Types
	
	Check_Status  = par1.work_status
	Filling_Types = "Filling Individual [EnerTech] CheckPoint(1)"
	Info_Msg      = Filling_Types & " ("& par1.get_rack_name & "|" & Check_Status &")" 
	
	par1.ClearIndividualInternalTag
	par1.SentValue_Medical_Plc
	 
	Select Case CLng(Check_Status)
		   Case 0 , 1
		   		Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
		   Case 2
			 	par1.Deactivate_CheckingPoint_1
			 	Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
		   Case 12
		   		par1.Reset_FillIndividual_EnerTech
		   		par1.ClearIndividualInternalTag
		   		par1.ClearPLC_Ind_EnerTech()
		   Case Else
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf)
			 	Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
	End Select
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function filling_individual_exec_enertech_1 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If	
End Function

'CheckPoint 2 : EnerTech ************************************************************
Function filling_individual_exec_enertech_2(par1)
On Error Resume Next
	Dim Check_Status , Info_Msg , Filling_Types
	
	Check_Status  = par1.work_status
	Filling_Types = "Filling Individual [EnerTech] CheckPoint(2)"  
	Info_Msg      = Filling_Types & " ("& par1.get_rack_name & "|" & Check_Status &")"
	
	Select Case CLng(Check_Status)
		   Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf)
			 	Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
		   Case 9
		   		If par1.capture_start_timestamp = "" Then
		   			par1.capture_start_timestamp = TimeNow()
		   			Call GF_LogToFile_("Exec", "Start Time : " & TimeNow(),"Individual") 
		   		End If
			 	Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
		   Case 11
				'remove checking internal tag time is empty
		   		'If par1.capture_end_timestamp = "" Then
		   		'	par1.capture_end_timestamp = TimeNow()
		   		'	par1.Deactivate_CheckingPoint_2
		   		'	Call GF_LogToFile_("Exec", "End Time : " & TimeNow(),"Individual") 
		   		'End If
				par1.capture_end_timestamp = TimeNow()
		   		par1.Deactivate_CheckingPoint_2
		   		Call GF_LogToFile_("Exec", "End Time : " & TimeNow(),"Individual") 
		   		Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
		   Case 12
		   		par1.Reset_FillIndividual_EnerTech
		   		par1.ClearIndividualInternalTag
		   		par1.ClearPLC_Ind_EnerTech()
		   		Call GF_LogToFile_("Exec", Info_Msg ,"Individual")			
		   Case Else
			    HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			    Call GF_LogToFile_("Exec", Info_Msg ,"Individual")	
	End Select
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function filling_individual_exec_enertech_2 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function

'CheckPoint 3 : EnerTech ************************************************************
Function filling_individual_exec_enertech_3(par1)
On Error Resume Next	
	Dim Check_Status , Info_Msg , Filling_Types
	
	Check_Status  = par1.fill_status 
	Filling_Types = "Filling Individual [EnerTech] CheckPoint(3)"
	Info_Msg      = Filling_Types & " ("& par1.get_rack_name & "|" & Check_Status &")"
	
	Select Case CLng(Check_Status)
		   Case 0, 2, 3, 4, 5, 6, 7, 8, 10
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
		   Case 1
			 	par1.analysis_required
				par1.Deactivate_CheckingPoint_3
				par1.ClearIndividualInternalTag
				par1.ClearPLC_Ind_EnerTech()
				Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
		   Case 12
		   		par1.Reset_FillIndividual_EnerTech
		   		par1.ClearIndividualInternalTag
		   		par1.ClearPLC_Ind_EnerTech()
		   Case Else
			    HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			    Call GF_LogToFile_("Exec", Info_Msg ,"Individual")
	End Select
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function filling_individual_exec_enertech_3 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function



'CheckPoint 1 : Cryostar ************************************************************
Function filling_individual_exec_cryostar_1(par1)
On Error Resume Next
	Dim Check_Status , Info_Msg , Filling_Types
	
	Check_Status  = par1.work_status
	Filling_Types = "Filling Individual [Cryostar] CheckPoint(1)" 
	Info_Msg      = Filling_Types & " ("& par1.get_rack_name & "|" & Check_Status &")"
	
	par1.ClearIndividualInternalTag
	par1.SentValue_Industry_Plc
	 
	Select Case CLng(Check_Status)
		   Case 0 , 1 
		   		Call GF_LogToFile_("Execute", Info_Msg ,"Individual")
		   Case 2
			 	par1.Deactivate_Cryostar_CheckingPoint_1
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Individual")
		   Case 12
		   		par1.Reset_FillIndividual_Cryostar
		   		par1.ClearIndividualInternalTag
		   		par1.ClearPLC_Ind_Cryostar()
		   Case Else
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf)
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Individual")
	End Select
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function filling_individual_exec_cryostar_1 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function

'CheckPoint 2 : Cryostar ************************************************************
Function filling_individual_exec_cryostar_2(par1)
On Error Resume Next
	Dim Check_Status , Info_Msg , Filling_Types
	
	Check_Status  = par1.work_status
	Filling_Types = "Filling Individual [Cryostar] CheckPoint(2)"
	Info_Msg      = Filling_Types & " ("& par1.get_rack_name & "|" & Check_Status &")" 
	
	Select Case CLng(Check_Status)
		   Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Individual")
		   Case 9
		   		If par1.capture_start_timestamp = "" Then
		   			par1.capture_start_timestamp = TimeNow() 
		   			Call GF_LogToFile_("Execute", Info_Msg, "Individual")
		   		End If
		   Case 11
				'remove checking internal tage is empty
		   		'If par1.capture_end_timestamp = "" Then
		   		'	par1.capture_end_timestamp = TimeNow()
		   		'	par1.Deactivate_Cryostar_CheckingPoint_2
		   		'	par1.ClearIndividualInternalTag
		   		'	par1.ClearPLC_Ind_Cryostar()
		   		'	Call GF_LogToFile_("RESET", Info_Msg, "Individual")
		   		'End If
				
					par1.capture_end_timestamp = TimeNow()
		   			par1.Deactivate_Cryostar_CheckingPoint_2
		   			par1.ClearIndividualInternalTag
		   			par1.ClearPLC_Ind_Cryostar()
		   			Call GF_LogToFile_("RESET", Info_Msg, "Individual")
				
		   Case 12
		   		par1.Reset_FillIndividual_Cryostar
		   		par1.ClearIndividualInternalTag
		   		par1.ClearPLC_Ind_Cryostar()
		   		Call GF_LogToFile_("RESET", Info_Msg ,"Individual")		
		   Case Else
			    HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			    Call GF_LogToFile_("Execute", Info_Msg ,"Individual")	
	End Select
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function filling_individual_exec_cryostar_2 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function

'Current Time : EnerTech/Cryostar ************************************************************
Function TimeNow()
On Error Resume Next
	TimeNow = YeAR(Date()) & "-" & _
			  TimeConvert(Month(Date()),2) & "-" & _
			  TimeConvert(DaY(Date()),2) & " " & _
			  Right("0" & Hour(Time),2) & ":" & _
			  Right("0" & Minute(Time),2) & ":" & _
			  Right("0" & Second(Time),2)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function TimeNow is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If		  
End Function

'Time Format : EnerTech/Cryostar ************************************************************
Function TimeConvert(n, totalDigits) 
On Error Resume Next
	If totalDigits > Len(n) Then 
		TimeConvert = String(totalDigits-Len(n),"0") & n 
	Else 
		TimeConvert = n 
	End If 
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Execute.bmo - Function TimeConvert is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If	
End Function