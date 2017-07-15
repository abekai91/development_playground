'Description : Interface [template for loose fillings] 
'----------------------------------------------------------------------------------------------------
Function Interface(par1,par2)
On Error Resume Next		
	Select Case CINT(par2)
		   Case 1
			 	Call Status_Batch_Point1(par1)
		   Case 2
			 	Call Status_Batch_Point2(par1)
		   Case Else
			 	HMIRuntime.Trace(Now & " - Parameter Interface Reach Limit ("& par2 & ")")
	End Select
If Err.Number <> 0 Then
	Call GF_LogError("Error", "Filling_Structure.bmo - Function Interface is not Workings [" & Err.Description & "]","Loose")
	Err.Clear
End If
End Function