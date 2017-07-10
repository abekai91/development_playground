'******************************************************************************************************
' Description  : Interface Individual (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: 27 April 2017
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function Interface_Individual(par1,par2)
	
	Select Case CINT(par2)
		   Case 1
			 	Call filling_individual_exec_enertech_1(par1)
		   Case 2
			 	Call filling_individual_exec_enertech_2(par1)
		   Case 3
			 	Call filling_individual_exec_enertech_3(par1)
		   Case 4
			 	Call filling_individual_exec_cryostar_1(par1)
		   Case 5
			 	Call filling_individual_exec_cryostar_2(par1)
		   Case Else
			 	HMIRuntime.Trace(Now & " - Parameter Interface  Individual Reach Limit ("& par2 & ")")
	End Select
End Function