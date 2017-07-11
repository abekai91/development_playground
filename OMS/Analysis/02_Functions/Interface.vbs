Function Interface_analysis(par1,par2)
	On Error Resume Next
	Select Case CINT(par2)
		   Case 1
			 	Call execute_cp_1(par1)
		   Case 2
			 	Call execute_cp_2(par1)
		   Case 3
			 	Call execute_cp_3(par1)
		   Case 4
			 	Call execute_cp_4(par1)
		   Case Else
			 	HMIRuntime.Trace(Now & " - Parameter Interface analysis Reach Limit ("& par2 & ")")
	End Select
	
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Interface.bmo - Function Interface_analysis is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
End Function