'******************************************************************************************************
' Description  : Maintable Queue (Action Scripts : FillingLoose_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 15 Feburary 2017
'******************************************************************************************************

Sub FillingLoose_Q

'Description : Declared MySQL Variables and Objects
'-----------------------------------------------------------------------------------------

Dim objConn,objRS

	Set objConn = CreateObject("ADODB.Connection")
	Set objRS   = CreateObject("ADODB.Recordset")
	
'Description : Declared MySQL Configuration
'-----------------------------------------------------------------------------------------	
Dim DB, UID, PASS, SVR
	
	SVR  = HMIRuntime.Tags("Server").Read
	DB   = HMIRuntime.Tags("Database").Read
	UID  = HMIRuntime.Tags("UID").Read
	PASS = HMIRuntime.Tags("PASS").Read

'Description : Open MySQL Database connection
'-----------------------------------------------------------------------------------------
	Call Mysql_Open_Conn(objConn,objRS,SVR,DB,UID,PASS)
	'Call Mysql_Init_Check(objConn)	
	
'Description : Declared Total Rack
'-----------------------------------------------------------------------------------------	
Dim F1,F2,F3,F4,F5,F6,F7

	Set F1 = (New Filling)(Array("Rack_1",1 ,param_one(objConn,objRS,1),param_two(objConn,objRS,1)))
	Set F2 = (New Filling)(Array("Rack_2",2 ,param_one(objConn,objRS,2),param_two(objConn,objRS,2)))
	Set F3 = (New Filling)(Array("Rack_3",3 ,param_one(objConn,objRS,3),param_two(objConn,objRS,3)))
	Set F4 = (New Filling)(Array("Rack_4",4 ,param_one(objConn,objRS,4),param_two(objConn,objRS,4)))
	Set F5 = (New Filling)(Array("Rack_5",5 ,param_one(objConn,objRS,5),param_two(objConn,objRS,5)))
	Set F6 = (New Filling)(Array("Rack_6",6 ,param_one(objConn,objRS,6),param_two(objConn,objRS,6)))
	Set F7 = (New Filling)(Array("Rack_7",7 ,param_one(objConn,objRS,7),param_two(objConn,objRS,7)))

'Calculate Total Available Rack In Check Point-1
'-----------------------------------------------------------------------------------------
Dim totalRack,cp1
	cp1 = 0
	Call Check_Point_1(objRS,objConn,totalRack,RackNo)
	
'Checked the Active Rack In Check Point-1
'-----------------------------------------------------------------------------------------
Dim rackList_cp1 : rackList_cp1 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp1 = Split(RackNo,",")
		cp1 = 1
	End If
	
'Calculate Total Available Rack In Check Point-2
'-----------------------------------------------------------------------------------------	
Dim cp2 : cp2 = 0

	Call Check_Point_2(objRS,objConn,totalRack,RackNo)

'Checked the Active Rack In Check Point-2
'-----------------------------------------------------------------------------------------
Dim rackList_cp2 : rackList_cp2 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp2 = Split(RackNo,",")
		cp2 = 1
	End If
	
'Description : Run process for Check Point-1
'-----------------------------------------------------------------------------------------
Dim c1
Dim point_1 : point_1 = 1

	If cp1 > 0 Then
		For Each c1 In rackList_cp1
		
			Select Case CInt(c1)
				   Case 1
					 	Call Interface(F1,point_1)
				   Case 2
					 	Call Interface(F2,point_1)
				   Case 3
					 	Call Interface(F3,point_1)
				   Case 4
					 	Call Interface(F4,point_1)
				   Case 5
					 	Call Interface(F5,point_1)
				   Case 6
					    Call Interface(F6,point_1)
				   Case 7
					 	Call Interface(F7,point_1)
				   case else
					 	'HMIRuntime.Trace(Now & " CP 1 - Racks Not Available Inside Database ("& c1 & ")" & vbCrlf)
			End Select	
		Next
	End If	

'Description : Run process for Check Point-2
'-----------------------------------------------------------------------------------------
Dim c2
Dim point_2 : point_2 = 2
	If cp2 > 0 Then
		For Each c2 In rackList_cp2
		
			Select Case CInt(c2)
				   Case 1
					 	Call Interface(F1,point_2)
				   Case 2
					 	Call Interface(F2,point_2)
				   Case 3
					 	Call Interface(F3,point_2)
				   Case 4
					 	Call Interface(F4,point_2)
				   Case 5
					 	Call Interface(F5,point_2)
				   Case 6
					    Call Interface(F6,point_2)
				   Case 7
					 	Call Interface(F7,point_2)
				   case else
					 	'HMIRuntime.Trace( Now & " CP 2 - Racks Not Available Inside Database ("& c2 & ")" & vbCrlf)
			End Select	
		Next
	End If				
	
'Decription : Analysis 
'---------------------------------------------------------------------------
 	'Call codabix_tag_value()   
	
'Description : Closed MySQL Object Connection	
'---------------------------------------------------------------------------	
	Call Mysql_Close_Conn(objRS,objConn)

End Sub