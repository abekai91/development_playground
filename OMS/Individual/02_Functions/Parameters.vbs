'******************************************************************************************************
' Description  : param_individual_one (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_one(objConn,objRS,index)

	Set objRS = objConn.execute("Select prod_detail from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_one =  objRS(0).value
	
End Function


'******************************************************************************************************
' Description  : param_individual_two (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_two(objConn,objRS,index)

	Set objRS = objConn.execute("Select user_id from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_two =  objRS(0).value
	
End Function

'******************************************************************************************************
' Description  : param_individual_three (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_three(objConn,objRS,index)

	Set objRS = objConn.execute("Select IsPrefilled from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_three =  objRS(0).value
	
End Function


'******************************************************************************************************
' Description  : param_individual_four (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function param_individual_four(objConn,objRS,index)

	Set objRS = objConn.execute("Select shift_batch from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_individual_three =  objRS(0).value
	
End Function