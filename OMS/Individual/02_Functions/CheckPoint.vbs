'******************************************************************************************************
' Description  : Check_Point_DB360 (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function Check_Point_DB360(objRS,objConn,totalRack,RackNo)
On Error Resume Next
	totalRack = 0
	RackNo = 0
	Set objRS = objConn.execute("Select count(distinct rack_id) from codabix_trigger where" & _
								" state = 1 and db_360 = 1 and db_361 = 0 and db_362 = 0" & _
								" and trigger_name = 'DB2001_Trigger'")

	totalRack = objRS(0).value

'Get Active Racks
'----------------------------------------------------------------------------------------------------

	RackNo = 0
	Set objRS = objConn.execute("Select distinct state , db_360 , rack_id from codabix_trigger where" & _
							" state = 1 and db_360 = 1 and db_361 = 0 and db_362 = 0" & _
							" and trigger_name = 'DB2001_Trigger'") 


	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(2).value	
		objRS.MoveNext
	Loop

	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_DB360 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function

'******************************************************************************************************
' Description  : Check_Point_DB361 (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Function Check_Point_DB361(objRS,objConn,totalRack,RackNo)
On Error Resume Next
	totalRack = 0
	RackNo = 0
	Set objRS = objConn.execute("Select count(distinct rack_id) from codabix_trigger where" & _
								" state = 1 and db_360 = 0 and db_361 = 1 and db_362 = 0" & _
								" and trigger_name = 'DB2001_Trigger'")

	totalRack = objRS(0).value

'Get Active Racks
'----------------------------------------------------------------------------------------------------

	RackNo = 0
	Set objRS = objConn.execute("Select distinct state , db_361 , rack_id from codabix_trigger where" & _
							" state = 1 and db_360 = 0 and db_361 = 1 and db_362 = 0" & _
							" and trigger_name = 'DB2001_Trigger'") 


	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(2).value	
		objRS.MoveNext
	Loop

	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_DB361 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function


'******************************************************************************************************
' Description  : Check_Point_DB362 (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************

Function Check_Point_DB362(objRS,objConn,totalRack,RackNo)
On Error Resume Next
	totalRack = 0
	RackNo = 0
	Set objRS = objConn.execute("Select count(distinct rack_id) from codabix_trigger where" & _
								" state = 1 and db_360 = 0 and db_361 = 0 and db_362 = 1" & _
								" and trigger_name = 'DB2001_Trigger'")

	totalRack = objRS(0).value

'Get Active Racks
'----------------------------------------------------------------------------------------------------

	RackNo = 0
	Set objRS = objConn.execute("Select distinct state , db_362 , rack_id from codabix_trigger where" & _
							" state = 1 and db_360 = 0 and db_361 = 0 and db_362 = 1" & _
							" and trigger_name = 'DB2001_Trigger'") 


	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(2).value	
		objRS.MoveNext
	Loop
	
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_DB362 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If

End Function


'******************************************************************************************************
' Description  : Check_Point_DB330 (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 27 April 2017
'******************************************************************************************************
Function Check_Point_DB330(objRS,objConn,totalRack,RackNo)
On Error Resume Next
	totalRack = 0
	RackNo = 0
	Set objRS = objConn.execute("Select count(distinct rack_id) from codabix_trigger where" & _
								" state = 1 and db_330 = 1 and db_331 = 0 " & _
								" and trigger_name = 'DB2001_Trigger'")

	totalRack = objRS(0).value

'Get Active Racks
'----------------------------------------------------------------------------------------------------

	RackNo = 0
	Set objRS = objConn.execute("Select distinct state , db_330 , rack_id from codabix_trigger where" & _
							" state = 1 and db_330 = 1 and db_331 = 0 " & _
							" and trigger_name = 'DB2001_Trigger'") 


	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(2).value	
		objRS.MoveNext
	Loop
	
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_DB330 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function

'******************************************************************************************************
' Description  : Check_Point_DB361 (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 27 April 2017
'******************************************************************************************************
Function Check_Point_DB331(objRS,objConn,totalRack,RackNo)
On Error Resume Next
	totalRack = 0
	RackNo = 0
	Set objRS = objConn.execute("Select count(distinct rack_id) from codabix_trigger where" & _
								" state = 1 and db_330 = 0 and db_331 = 1 " & _
								" and trigger_name = 'DB2001_Trigger'")

	totalRack = objRS(0).value

'Get Active Racks
'----------------------------------------------------------------------------------------------------

	RackNo = 0
	Set objRS = objConn.execute("Select distinct state , db_331 , rack_id from codabix_trigger where" & _
							" state = 1 and db_330 = 0 and db_331 = 1 " & _
							" and trigger_name = 'DB2001_Trigger'") 


	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(2).value	
		objRS.MoveNext
	Loop

	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_DB331 is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Function