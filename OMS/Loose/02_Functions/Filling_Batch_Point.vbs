'Filling Loose CheckPoint 1 : EnerTech ************************************************************
Function Status_Batch_Point1(par1)
	
	 Dim Check_Status , Info_Msg , Filling_Types
	 
	 Check_Status  = par1.getStatus_DB302
	 Filling_Types = "Filling Loose CheckPoint(1)"
	 Info_Msg      = Filling_Types & " ("& par1.getRack_name & "|" & Check_Status &")"
	 
	 par1.ClearLooseInternalTag
	 par1.set_ActStatus_DB301
	 par1.set_db301_recipe
	 
	Select Case CINT(Check_Status)
		   Case 0 , 1
		   		Call GF_LogToFile_("Exec", Info_Msg ,"Loose")
		   Case 2
			 	par1.set_DeactStatus_DB301
			 	par1.set_DeactCheckPoint_1
			 	Call GF_LogToFile_("Exec", Info_Msg ,"Loose")
		   Case 12
		   		par1.Reset_FillLoose_EnerTech
		   		par1.ClearLooseInternalTag
		   		par1.ClearPLC_Loose_EnerTech()
		   Case Else
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf)
			 	Call GF_LogToFile_("Exec", Info_Msg ,"Loose")
	End Select
	
End Function

'Filling Loose CheckPoint 2 : EnerTech ************************************************************
Function Status_Batch_Point2(par1)
	
	Dim Check_Status, Final_Status , Info_Msg , Filling_Types
	
	Check_Status  = par1.getStatus_DB302
	Filling_Types = "Filling Loose CheckPoint(2)"
	Info_Msg      = Filling_Types & " ("& par1.getRack_name & "|" & Check_Status &")"
	 
	Select Case CINT(Check_Status)
		   Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10
			 	HMIRuntime.Trace(Now & Info_Msg & vbCrlf )
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Loose")
		   Case 9
		   		If par1.time_start = "" Then
		   			par1.setS_Time
			 		par1.setS_Date
		   		End If
		   		
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Loose")
		   Case 11
		   		If par1.time_end = "" Then
			   		par1.setE_Time
			   		par1.setE_Date
			   		par1.set_DeactFilling_Batch	
		   		End If
		   		
		   		Call GF_LogToFile_("Execute", Info_Msg ,"Loose")
		   		
		   Case 12
		   		par1.Reset_FillLoose_EnerTech
		   		par1.ClearLooseInternalTag
		   		par1.ClearPLC_Loose_EnerTech()
		   		Call GF_LogToFile_("Execute", Info_Msg ,"Loose")
		   Case Else
		   		HMIRuntime.Trace(Now & Info_Msg & vbCrlf)
			 	Call GF_LogToFile_("Execute", Info_Msg ,"Loose")
	End Select
	
End Function