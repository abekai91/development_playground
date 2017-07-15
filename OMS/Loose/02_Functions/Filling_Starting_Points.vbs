'******************************************************************************************************
' Description   : Checked point 1  
' Created By    : Ahmad
' Noted         : 
' Dated Created : 16 Febuary 2017
'******************************************************************************************************
Function Check_Point_1(objRS,objConn,totalRack,RackNo)
On Error Resume Next	
Dim t_point,t_count,Order_Status
	totalRack = 0
	RackNo 	  = 0
	t_count   = 0

'Count Active Rack	
'-----------------
	Set objRS = objConn.execute("Select count(rack_id) from codabix_trigger where" & _
     				  		   " state=1 and db_301=1" & _
     				  		   " and trigger_name='DB2001_Trigger'")
    
	t_count		=  objRS(0).value
    totalRack   =  t_count
    
'Get Active Rack  
'-----------------

	RackNo 	  = 0   				  		   
	Set objRS = objConn.execute("Select state , db_301 , rack_id  from codabix_trigger where" & _
     				  		   " state=1 and db_301=1" & _
     				  		   " and trigger_name='DB2001_Trigger'")
	
	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(2).value
		objRS.MoveNext	
	Loop	
If Err.Number <> 0 Then
	Call GF_LogError("Error", "Filling_Starting_Points.bmo - Function Check_Point_1 is not Workings [" & Err.Description & "]","Loose")
	Err.Clear
End If
End Function


'******************************************************************************************************
' Description   : Checked point 2
' Created By    : Ahmad
' Noted         : 
' Dated Created : 16 Febuary 2017
'******************************************************************************************************
Function Check_Point_2(objRS,objConn,totalRack,RackNo)
On Error Resume Next
Dim t_point,t_count,Order_Status
	totalRack = 0 
	RackNo = 0
	
'Count Active Rack	
'-----------------
	Set objRS = objConn.execute("Select count(rack_id) from codabix_trigger where" & _
     				  		   " state=1 and db_301=0" & _
     				  		   " and trigger_name='DB2001_Trigger'")
    t_count		=  objRS(0).value
    totalRack   =  t_count
    
    
'Get Active Rack  
'-----------------	  		   
	Set objRS = objConn.execute("Select state , db_301 , rack_id  from codabix_trigger where" & _
     				  		   " state=1 and db_301=0" & _
     				  		   " and trigger_name='DB2001_Trigger'")
	
	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(2).value
		objRS.MoveNext	
	Loop
If Err.Number <> 0 Then
	Call GF_LogError("Error", "Filling_Starting_Points.bmo - Function Check_Point_2 is not Workings [" & Err.Description & "]","Loose")
	Err.Clear
End If	
End Function