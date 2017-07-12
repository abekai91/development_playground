'******************************************************************************************************
' Description  : param_individual_one (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_one(objConn,objRS,index)
On Error Resume Next
	Set objRS = objConn.execute("Select prod_detail from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_one =  objRS(0).value
If Err.Number <> 0 Then
		Call GF_LogError("Error", "Parameters.bmo - Function param_individual_one is not Workings [" & Err.Description & "]","Individual")
		Err.Clear
End If
End Function


'******************************************************************************************************
' Description  : param_individual_two (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_two(objConn,objRS,index)
On Error Resume Next
	Set objRS = objConn.execute("Select user_id from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_two =  objRS(0).value
If Err.Number <> 0 Then
		Call GF_LogError("Error", "Parameters.bmo - Function param_individual_two is not Workings [" & Err.Description & "]","Individual")
		Err.Clear
End If	
End Function

'******************************************************************************************************
' Description  : param_individual_three (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_three(objConn,objRS,index)
On Error Resume Next
	Set objRS = objConn.execute("Select IsPrefilled from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_three =  objRS(0).value
If Err.Number <> 0 Then
		Call GF_LogError("Error", "Parameters.bmo - Function param_individual_three is not Workings [" & Err.Description & "]","Individual")
		Err.Clear
End If		
End Function


'******************************************************************************************************
' Description  : param_individual_four (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_four(objConn,objRS,index)
On Error Resume Next
	Set objRS = objConn.execute("Select shift_batch from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_four =  objRS(0).value
If Err.Number <> 0 Then
		Call GF_LogError("Error", "Parameters.bmo - Function param_individual_four is not Workings [" & Err.Description & "]","Individual")
		Err.Clear
End If	
End Function