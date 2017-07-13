' Description : parameter 1
'--------------------------------------------------------------------------
Function param_ana_one(objConn,objRS,index) 'Information to Fill
	On Error Resume Next
	Set objRS = objConn.execute("Select cyl_id,batch_id,prod_id,cyl_to_analyse from analysis_rack where" & _
     				  		   " rack_id = "& index &"")
    
	param_ana_one =  objRS(0).value & "|" & objRS(1).value & "|" & objRS(2).value & "|" & objRS(3).value
	
	If Err.Number <> 0 Then
		   Call GF_LogError("Error", "Parameters.bmo - Function param_ana_one is not Workings [" & Err.Description & "]","Analysis")
		   Err.Clear
	End If
End Function

' Description : parameter 2
'--------------------------------------------------------------------------
Function param_ana_two(objConn,objRS,index) 'OMS Trigger Fill
	On Error Resume Next
	Set objRS = objConn.execute("Select oms_fill from analysis_rack where" & _
     				  		   " rack_id = "& index &"")
    
	param_ana_two =  objRS(0).value
	
	If Err.Number <> 0 Then
		   Call GF_LogError("Error", "Parameters.bmo - Function param_ana_two is not Workings [" & Err.Description & "]","Analysis")
		   Err.Clear
	End If
End Function

' Description : parameter 3
'--------------------------------------------------------------------------
Function param_ana_three(objConn,objRS,index) 'OMS Trigger Fill
	On Error Resume Next
	Set objRS = objConn.execute("Select GUID from analysis_rack where" & _
     				  		   " rack_id = "& index &"")
    
	param_ana_three =  objRS(0).value
	
	If Err.Number <> 0 Then
		   Call GF_LogError("Error", "Parameters.bmo - Function param_ana_three is not Workings [" & Err.Description & "]","Analysis")
		   Err.Clear
	End If
End Function

' Description : parameter 4
'--------------------------------------------------------------------------
Function param_ana_four(objConn,objRS,index) 'OMS Trigger Fill
	On Error Resume Next
	Set objRS = objConn.execute("Select user_id from analysis_rack where" & _
     				  		   " rack_id = "& index &"")
    
	param_ana_four =  objRS(0).value
	If Err.Number <> 0 Then
		   Call GF_LogError("Error", "Parameters.bmo - Function param_ana_four is not Workings [" & Err.Description & "]","Analysis")
		   Err.Clear
	End If
End Function