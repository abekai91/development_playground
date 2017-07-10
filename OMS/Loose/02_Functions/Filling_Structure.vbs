Function Interface(par1,par2)
	
	Select Case CINT(par2)
		   Case 1
			 	Call Status_Batch_Point1(par1)
		   Case 2
			 	Call Status_Batch_Point2(par1)
		   Case Else
			 	HMIRuntime.Trace(Now & " - Parameter Interface Reach Limit ("& par2 & ")")
	End Select
End Function