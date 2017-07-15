'Description : param_one [parameter one get the product details] 
'----------------------------------------------------------------------------------------------------
Function param_one(objConn,objRS,index)
On Error Resume Next	
	Set objRS = objConn.execute("Select prod_detail from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
	param_one =  objRS(0).value
If Err.Number <> 0 Then
	Call GF_LogError("Error", "Filling_Parameter.bmo - Function param_one is not Workings [" & Err.Description & "]","Loose")
	Err.Clear
End If		
End Function


'Description : param_two [parameter two get the user/filler ID] 
'----------------------------------------------------------------------------------------------------
Function param_two(objConn,objRS,index)
On Error Resume Next
	Set objRS = objConn.execute("Select user_id from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_two =  objRS(0).value
If Err.Number <> 0 Then
	Call GF_LogError("Error", "Filling_Parameter.bmo - Function param_two is not Workings [" & Err.Description & "]","Loose")
	Err.Clear
End If		
End Function