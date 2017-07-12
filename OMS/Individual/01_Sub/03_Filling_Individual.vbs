'******************************************************************************************************
' Description  : Individual Queue (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 April 2017
'******************************************************************************************************

Sub FillingIndividual_Q
On Error Resume Next
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
	

'Description : Declared Total Rack For Medical EnerTech
'-----------------------------------------------------------------------------------------	
Dim F1,F2,F3,F4,F5,F6,F7,F8,F9

	Set F1 = (New Filling_Individual)(Array("CO2_R1",1 ,param_individual_one(objConn,objRS,30),param_individual_two(objConn,objRS,30),param_individual_three(objConn,objRS,30),30,param_individual_four(objConn,objRS,30)))
	Set F2 = (New Filling_Individual)(Array("CO2_R2",2 ,param_individual_one(objConn,objRS,31),param_individual_two(objConn,objRS,31),param_individual_three(objConn,objRS,31),31,param_individual_four(objConn,objRS,31)))
	Set F3 = (New Filling_Individual)(Array("CO2_R3",3 ,param_individual_one(objConn,objRS,32),param_individual_two(objConn,objRS,32),param_individual_three(objConn,objRS,32),32,param_individual_four(objConn,objRS,32)))
	Set F4 = (New Filling_Individual)(Array("N2O_R1",4 ,param_individual_one(objConn,objRS,33),param_individual_two(objConn,objRS,33),param_individual_three(objConn,objRS,33),33,param_individual_four(objConn,objRS,33)))
	Set F5 = (New Filling_Individual)(Array("N2O_R2",5 ,param_individual_one(objConn,objRS,34),param_individual_two(objConn,objRS,34),param_individual_three(objConn,objRS,34),34,param_individual_four(objConn,objRS,34)))
	Set F6 = (New Filling_Individual)(Array("N2O_R3",6 ,param_individual_one(objConn,objRS,35),param_individual_two(objConn,objRS,35),param_individual_three(objConn,objRS,35),35,param_individual_four(objConn,objRS,35)))
	Set F7 = (New Filling_Individual)(Array("EOX_1",7  ,param_individual_one(objConn,objRS,36),param_individual_two(objConn,objRS,36),param_individual_three(objConn,objRS,36),36,param_individual_four(objConn,objRS,36)))
	Set F8 = (New Filling_Individual)(Array("EOX_2",8  ,param_individual_one(objConn,objRS,37),param_individual_two(objConn,objRS,37),param_individual_three(objConn,objRS,37),37,param_individual_four(objConn,objRS,37)))
	Set F9 = (New Filling_Individual)(Array("EOX_3",9  ,param_individual_one(objConn,objRS,38),param_individual_two(objConn,objRS,38),param_individual_three(objConn,objRS,38),38,param_individual_four(objConn,objRS,38)))

'Description : Declared Total Rack For Industry Cryostar
'-----------------------------------------------------------------------------------------	
Dim F10,F11,F12

	Set F10 = (New Filling_Individual)(Array("CO2_R1",10 ,param_individual_one(objConn,objRS,30),param_individual_two(objConn,objRS,30),param_individual_three(objConn,objRS,30),30,param_individual_four(objConn,objRS,30)))
	Set F11 = (New Filling_Individual)(Array("CO2_R2",11 ,param_individual_one(objConn,objRS,31),param_individual_two(objConn,objRS,31),param_individual_three(objConn,objRS,31),31,param_individual_four(objConn,objRS,31)))
	Set F12 = (New Filling_Individual)(Array("CO2_R3",12 ,param_individual_one(objConn,objRS,32),param_individual_two(objConn,objRS,32),param_individual_three(objConn,objRS,32),32,param_individual_four(objConn,objRS,32)))
	
		
	
'Calculate Total Available Rack In Check Point-DB360 [cp1]
'-----------------------------------------------------------------------------------------
Dim totalRack,cp_db360
	cp_db360 = 0
	Call Check_Point_DB360(objRS,objConn,totalRack,RackNo)
	
'Checked the Active Rack In Check Point-db360 [cp1]
'-----------------------------------------------------------------------------------------
Dim rackList_cp_db360 : rackList_cp_db360 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp_db360 = Split(RackNo,",")
		cp_db360 = 1
	End If


	
'Calculate Total Available Rack In Check Point-DB361 [cp2]
'-----------------------------------------------------------------------------------------	
Dim cp_db361 : cp_db361 = 0

	Call Check_Point_DB361(objRS,objConn,totalRack,RackNo)

'Checked the Active Rack In Check Point-DB361 [cp2]
'-----------------------------------------------------------------------------------------
Dim rackList_cp_db361 : rackList_cp_db361 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp_db361 = Split(RackNo,",")
		cp_db361 = 1
	End If	

'Calculate Total Available Rack In Check Point-DB362 [cp3]
'-----------------------------------------------------------------------------------------	
Dim cp_db362 : cp_db362 = 0

	Call Check_Point_DB362(objRS,objConn,totalRack,RackNo)

'Checked the Active Rack In Check Point-DB362 [cp3]
'-----------------------------------------------------------------------------------------
Dim rackList_cp_db362 : rackList_cp_db362 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp_db362 = Split(RackNo,",")
		cp_db362 = 1
	End If	
	

'*****************************Cryostar Parts *******************************

'Calculate Total Available Rack In Check Point-DB330 [cp1]
'-----------------------------------------------------------------------------------------
Dim cp_db330 : cp_db330 = 0
	Call Check_Point_DB330(objRS,objConn,totalRack,RackNo)
	
'Checked the Active Rack In Check Point-db360 [cp1]
'-----------------------------------------------------------------------------------------
Dim rackList_cp_db330 : rackList_cp_db330 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp_db330 = Split(RackNo,",")
		cp_db330 = 1
	End If


	
'Calculate Total Available Rack In Check Point-DB331 [cp2]
'-----------------------------------------------------------------------------------------	
Dim cp_db331 : cp_db331 = 0

	Call Check_Point_DB331(objRS,objConn,totalRack,RackNo)

'Checked the Active Rack In Check Point-DB361 [cp2]
'-----------------------------------------------------------------------------------------
Dim rackList_cp_db331 : rackList_cp_db331 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp_db331 = Split(RackNo,",")
		cp_db331 = 1
	End If	
	
'************************ End Cryostar Parts *******************************	
	
	
	
'Description : Run process for Check Point-1
'-----------------------------------------------------------------------------------------
Dim c1, c330
Dim point_1 : point_1 = 1
Dim point_330 : point_330 = 4

	If cp_db360 > 0 Then
		For Each c1 In rackList_cp_db360
		
			Select Case CInt(c1)
				   Case 30
					 	Call Interface_Individual(F1,point_1)
				   Case 31
					 	Call Interface_Individual(F2,point_1)
				   Case 32
					 	Call Interface_Individual(F3,point_1)
				   Case 33
					 	Call Interface_Individual(F4,point_1)
				   Case 34
					 	Call Interface_Individual(F5,point_1)
				   Case 35
					    Call Interface_Individual(F6,point_1)
				   Case 36
					 	Call Interface_Individual(F7,point_1)
				   Case 37
					 	Call Interface_Individual(F8,point_1)
				   Case 38
					 	Call Interface_Individual(F9,point_1)
				   case else
					 	'HMIRuntime.Trace(Now & " CP 1 - Racks Not Available Inside Database ("& c1 & ")" & vbCrlf)
			End Select	
		Next
	End If
		
	If cp_db330 > 0 Then
		For Each c330 In rackList_cp_db330
		
			Select Case CInt(c330)
				   Case 30
					 	Call Interface_Individual(F10,point_330)
				   Case 31
					 	Call Interface_Individual(F11,point_330)
				   Case 32
					 	Call Interface_Individual(F12,point_330)
				   case else
					 	'HMIRuntime.Trace(Now & " CP 1 - Racks Not Available Inside Database ("& c330 & ")" & vbCrlf)
			End Select	
		Next
		
	End If	

'Description : Run process for Check Point-2
'-----------------------------------------------------------------------------------------
Dim c2, c331
Dim point_2 : point_2 = 2
Dim point_331 : point_331 = 5


	If cp_db361 > 0 Then
		For Each c2 In rackList_cp_db361
		
			Select Case CInt(c2)
				   Case 30
					 	Call Interface_Individual(F1,point_2)
				   Case 31
					 	Call Interface_Individual(F2,point_2)
				   Case 32
					 	Call Interface_Individual(F3,point_2)
				   Case 33
					 	Call Interface_Individual(F4,point_2)
				   Case 34
					 	Call Interface_Individual(F5,point_2)
				   Case 35
					    Call Interface_Individual(F6,point_2)
				   Case 36
					 	Call Interface_Individual(F7,point_2)
				   Case 37
					 	Call Interface_Individual(F8,point_2)
				   Case 38
					 	Call Interface_Individual(F9,point_2)
				   case else
					 	'HMIRuntime.Trace( Now & " CP 2 - Racks Not Available Inside Database ("& c2 & ")" & vbCrlf)
			End Select	
		Next
	End If	
	
	If cp_db331 > 0 Then
		For Each c331 In rackList_cp_db331
		
			Select Case CInt(c331)
				   Case 30
					 	Call Interface_Individual(F10,point_331)
				   Case 31
					 	Call Interface_Individual(F11,point_331)
				   Case 32
					 	Call Interface_Individual(F12,point_331)
				   case else
					 	'HMIRuntime.Trace( Now & " CP 2 - Racks Not Available Inside Database ("& c331 & ")" & vbCrlf)
			End Select	
		Next
	End If	
	
'Description : Run process for Check Point-3
'-----------------------------------------------------------------------------------------
Dim c3
Dim point_3 : point_3 = 3
	If cp_db362 > 0 Then
		For Each c3 In rackList_cp_db362
		
			Select Case CInt(c3)
				   Case 30
					 	Call Interface_Individual(F1,point_3)
				   Case 31
					 	Call Interface_Individual(F2,point_3)
				   Case 32
					 	Call Interface_Individual(F3,point_3)
				   Case 33
					 	Call Interface_Individual(F4,point_3)
				   Case 34
					 	Call Interface_Individual(F5,point_3)
				   Case 35
					    Call Interface_Individual(F6,point_3)
				   Case 36
					 	Call Interface_Individual(F7,point_3)
				   Case 37
					 	Call Interface_Individual(F8,point_3)
				   Case 38
					 	Call Interface_Individual(F9,point_3)
				   case else
					 	'HMIRuntime.Trace( Now & " CP 3 - Racks Not Available Inside Database ("& c3 & ")" & vbCrlf)
			End Select	
		Next
	End If	

'Description : Closed MySQL Object Connection	
'---------------------------------------------------------------------------	
	Call Mysql_Close_Conn(objRS,objConn)
	
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "03_Filling_Individual.bmo - Sub FillingIndividual_Q is not Workings [" & Err.Description & "]","Individual")
		    Err.Clear
	End If
End Sub