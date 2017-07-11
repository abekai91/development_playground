' Description : Checkpoint 1
'--------------------------------------------------------------------------
Function Check_Point_Ana_1(objRS,objConn,totalRack,RackNo)
On Error Resume Next
Dim t_point,t_count,Order_Status
	totalRack = 0
	RackNo 	  = 0
	t_count   = 0

'Count Active Rack	
	Set objRS = objConn.execute("Select count(distinct rack_id) from analysis_rack where" & _
     				  		   " occupied=1 and cp1=1 and cp2=0 and cp3=0 and cp4=0")
    
	t_count		=  objRS(0).value
    totalRack   =  t_count
    
'Get Active Rack  
	RackNo 	  = 0   				  		   
	Set objRS = objConn.execute("Select distinct rack_id from analysis_rack where" & _
     				  		   " occupied=1 and cp1=1 and cp2=0 and cp3=0 and cp4=0")
	
	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(0).value
		objRS.MoveNext	
	Loop	
	
	If Err.Number <> 0 Then
		  Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_Ana_1 is not Workings [" & Err.Description & "]","Analysis")
		  Err.Clear
	End If
End Function

' Description : Checkpoint 2
'--------------------------------------------------------------------------
Function Check_Point_Ana_2(objRS,objConn,totalRack,RackNo)
On Error Resume Next
Dim t_point,t_count,Order_Status
	totalRack = 0
	RackNo 	  = 0
	t_count   = 0

'Count Active Rack	
	Set objRS = objConn.execute("Select count(distinct rack_id) from analysis_rack where" & _
     				  		   " occupied=1 and cp1=0 and cp2=1 and cp3=0 and cp4=0")
    
	t_count		=  objRS(0).value
    totalRack   =  t_count
    
'Get Active Rack  
	RackNo 	  = 0   				  		   
	Set objRS = objConn.execute("Select distinct rack_id from analysis_rack where" & _
     				  		   " occupied=1 and cp1=0 and cp2=1 and cp3=0 and cp4=0")
	
	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(0).value
		objRS.MoveNext	
	Loop
	
	If Err.Number <> 0 Then
		  Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_Ana_2 is not Workings [" & Err.Description & "]","Analysis")
		  Err.Clear
	End If	
End Function

' Description : Checkpoint 3
'--------------------------------------------------------------------------
Function Check_Point_Ana_3(objRS,objConn,totalRack,RackNo)
On Error Resume Next
Dim t_point,t_count,Order_Status
	totalRack = 0
	RackNo 	  = 0
	t_count   = 0

'Count Active Rack	
	Set objRS = objConn.execute("Select count(distinct rack_id) from analysis_rack where" & _
     				  		   " occupied=1 and cp1=0 and cp2=0 and cp3=1 and cp4=0")
    
	t_count		=  objRS(0).value
    totalRack   =  t_count
    
'Get Active Rack  
	RackNo 	  = 0   				  		   
	Set objRS = objConn.execute("Select distinct rack_id from analysis_rack where" & _
     				  		   " occupied=1 and cp1=0 and cp2=0 and cp3=1 and cp4=0")
	
	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(0).value
		objRS.MoveNext	
	Loop	
	
	If Err.Number <> 0 Then
		  Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_Ana_3 is not Workings [" & Err.Description & "]","Analysis")
		  Err.Clear
	End If
	
End Function

' Description : Checkpoint 4
'--------------------------------------------------------------------------
Function Check_Point_Ana_4(objRS,objConn,totalRack,RackNo)
On Error Resume Next
Dim t_point,t_count,Order_Status
	totalRack = 0
	RackNo 	  = 0
	t_count   = 0

'Count Active Rack
	Set objRS = objConn.execute("Select count(distinct rack_id) from analysis_rack where" & _
     				  		   " occupied=1 and cp1=0 and cp2=0 and cp3=0 and cp4=1")
    
	t_count		=  objRS(0).value
    totalRack   =  t_count
    
'Get Active Rack  
	RackNo 	  = 0   				  		   
	Set objRS = objConn.execute("Select distinct rack_id from analysis_rack where" & _
     				  		   " occupied=1 and cp1=0 and cp2=0 and cp3=0 and cp4=1")
	
	Do While Not objRS.EOF
		RackNo = RackNo & "," & objRS(0).value
		objRS.MoveNext	
	Loop	
	
	If Err.Number <> 0 Then
		  Call GF_LogError("Error", "CheckPoint.bmo - Function Check_Point_Ana_4 is not Workings [" & Err.Description & "]","Analysis")
		  Err.Clear
	End If
End Function