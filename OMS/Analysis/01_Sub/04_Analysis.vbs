'******************************************************************************************************
' Description  : Analysis Queue (Action Scripts : Analysis_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 02 March 2017
'******************************************************************************************************

Sub Analysis_Q
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
	
'Description : Cylinder Counter	
'-----------------------------------------------------------------------------------------
	Call cylinder_count()
	
'Description : Declared Total Rack
'-----------------------------------------------------------------------------------------	
Dim A1,A2,A3,A4,A5,A6

	Set A1 = (New Analysis)(Array("AYM_1"    ,1 ,param_ana_one(objConn,objRS,1),param_ana_two(objConn,objRS,1),param_ana_three(objConn,objRS,1),param_ana_four(objConn,objRS,1)))
	Set A2 = (New Analysis)(Array("AYM_2_3"  ,2 ,param_ana_one(objConn,objRS,2),param_ana_two(objConn,objRS,2),param_ana_three(objConn,objRS,2),param_ana_four(objConn,objRS,2)))
	Set A3 = (New Analysis)(Array("AYM_4_5"  ,3 ,param_ana_one(objConn,objRS,3),param_ana_two(objConn,objRS,3),param_ana_three(objConn,objRS,3),param_ana_four(objConn,objRS,3)))
	Set A4 = (New Analysis)(Array("AYI_1_2"  ,4 ,param_ana_one(objConn,objRS,4),param_ana_two(objConn,objRS,4),param_ana_three(objConn,objRS,4),param_ana_four(objConn,objRS,4)))
	Set A5 = (New Analysis)(Array("AYI_3_4_5",5 ,param_ana_one(objConn,objRS,5),param_ana_two(objConn,objRS,5),param_ana_three(objConn,objRS,5),param_ana_four(objConn,objRS,5)))
	Set A6 = (New Analysis)(Array("AYI_6"    ,6 ,param_ana_one(objConn,objRS,6),param_ana_two(objConn,objRS,6),param_ana_three(objConn,objRS,6),param_ana_four(objConn,objRS,6)))

'Calculate Total Available Rack In Check Point_1
'-----------------------------------------------------------------------------------------
Dim totalRack,cp1
	cp1 = 0
	Call Check_Point_Ana_1(objRS,objConn,totalRack,RackNo)
	
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

	Call Check_Point_Ana_2(objRS,objConn,totalRack,RackNo)

'Checked the Active Rack In Check Point-2
'-----------------------------------------------------------------------------------------
Dim rackList_cp2 : rackList_cp2 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp2 = Split(RackNo,",")
		cp2 = 1
	End If
	
'Calculate Total Available Rack In Check Point-3
'-----------------------------------------------------------------------------------------	
Dim cp3 : cp3 = 0

	Call Check_Point_Ana_3(objRS,objConn,totalRack,RackNo)

'Checked the Active Rack In Check Point-2
'-----------------------------------------------------------------------------------------
Dim rackList_cp3 : rackList_cp3 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp3 = Split(RackNo,",")
		cp3 = 1
	End If
	
'Calculate Total Available Rack In Check Point-4
'-----------------------------------------------------------------------------------------	
Dim cp4 : cp4 = 0

	Call Check_Point_Ana_4(objRS,objConn,totalRack,RackNo)

'Checked the Active Rack In Check Point-2
'-----------------------------------------------------------------------------------------
Dim rackList_cp4 : rackList_cp4 = ""
	
	If CInt(totalRack) > 0 Then
		rackList_cp4 = Split(RackNo,",")
		cp4 = 1
	End If
	
'Description : Run process for Check Point-1
'-----------------------------------------------------------------------------------------
Dim c1
Dim point_1 : point_1 = 1

	If cp1 > 0 Then
		For Each c1 In rackList_cp1
		
			Select Case CInt(c1)
				   Case 1
					 	Call Interface_analysis(A1,point_1)
				   Case 2
					 	Call Interface_analysis(A2,point_1)
				   Case 3
					 	Call Interface_analysis(A3,point_1)
				   Case 4
					 	Call Interface_analysis(A4,point_1)
				   Case 5
					 	Call Interface_analysis(A5,point_1)
				   Case 6
					    Call Interface_analysis(A6,point_1)
				   case else
					 	'HMIRuntime.Trace(Now & " CP 1 - Analysis Racks Not Available Inside Database ("& c1 & ")" & vbCrlf)
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
					 	Call Interface_analysis(A1,point_2)
				   Case 2
					 	Call Interface_analysis(A2,point_2)
				   Case 3
					 	Call Interface_analysis(A3,point_2)
				   Case 4
					 	Call Interface_analysis(A4,point_2)
				   Case 5
					 	Call Interface_analysis(A5,point_2)
				   Case 6
					    Call Interface_analysis(A6,point_2)
				   case else
					 	'HMIRuntime.Trace( Now & " CP 2 - Analysis Racks Not Available Inside Database ("& c2 & ")" & vbCrlf)
			End Select	
		Next
	End If	

'Description : Run process for Check Point-3
'-----------------------------------------------------------------------------------------
Dim c3
Dim point_3 : point_3 = 3
	If cp3 > 0 Then
		For Each c3 In rackList_cp3
		
			Select Case CInt(c3)
				   Case 1
					 	Call Interface_analysis(A1,point_3)
				   Case 2
					 	Call Interface_analysis(A2,point_3)
				   Case 3
					 	Call Interface_analysis(A3,point_3)
				   Case 4
					 	Call Interface_analysis(A4,point_3)
				   Case 5
					 	Call Interface_analysis(A5,point_3)
				   Case 6
					    Call Interface_analysis(A6,point_3)
				   case else
					 	'HMIRuntime.Trace( Now & " CP 3 - Analysis Racks Not Available Inside Database ("& c3 & ")" & vbCrlf)
			End Select	
		Next
	End If		
	
'Description : Run process for Check Point-4
'-----------------------------------------------------------------------------------------
Dim c4
Dim point_4 : point_4 = 4
	If cp4 > 0 Then
		For Each c4 In rackList_cp4
		
			Select Case CInt(c4)
				   Case 1
					 	Call Interface_analysis(A1,point_4)
				   Case 2
					 	Call Interface_analysis(A2,point_4)
				   Case 3
					 	Call Interface_analysis(A3,point_4)
				   Case 4
					 	Call Interface_analysis(A4,point_4)
				   Case 5
					 	Call Interface_analysis(A5,point_4)
				   Case 6
					    Call Interface_analysis(A6,point_4)
				   case else
					 	'HMIRuntime.Trace( Now & " CP 4 - Analysis Racks Not Available Inside Database ("& c4 & ")" & vbCrlf)
			End Select	
		Next
	End If	


'Description : Individual Analysis Cylic (By End Of Shift)	
'---------------------------------------------------------------------------
Dim masterlist,time_h
Dim morning_trigger,noon_trigger,evening_trigger
	
	time_h = Hour(Now())
	
	'Read shift trigger time
		morning_trigger = HMIRuntime.Tags("individual_report_morning").Read
		noon_trigger = HMIRuntime.Tags("individual_report_noon").Read
		evening_trigger = HMIRuntime.Tags("individual_report_evening").Read
		
		If morning_trigger = "" Then
			HMIRuntime.Tags("individual_report_morning").Write 0
			morning_trigger = 0
		Elseif noon_trigger = "" Then
			HMIRuntime.Tags("individual_report_noon").Write 0
			noon_trigger = 0
		Elseif evening_trigger = "" Then
			HMIRuntime.Tags("individual_report_evening").Write 0
			evening_trigger = 0
		End If
		
		'check end of shift start from midnight to morning
		If time_h = 7 Then
			HMIRuntime.Tags("individual_report_morning").Write 0
			HMIRuntime.Tags("individual_report_noon").Write 0
			If HMIRuntime.Tags("individual_report_evening").Read = 0 Then
				Set masterlist = New MasterList_Reports 
				Set masterlist.Individual  = New Individual
				masterlist.Individual.Generate_Report
				HMIRuntime.Tags("individual_report_evening").Write 1
			End If
			
		Elseif time_h = 15 Then
			HMIRuntime.Tags("individual_report_noon").Write 0
			HMIRuntime.Tags("individual_report_evening").Write 0
			If HMIRuntime.Tags("individual_report_morning").Read = 0 Then
				Set masterlist = New MasterList_Reports 
				Set masterlist.Individual  = New Individual
				masterlist.Individual.Generate_Report
				HMIRuntime.Tags("individual_report_morning").Write 1
			End If
			
			
		Elseif time_h = 22 Then
			HMIRuntime.Tags("individual_report_morning").Write 0
			HMIRuntime.Tags("individual_report_evening").Write 0
			If HMIRuntime.Tags("individual_report_noon").Read = 0 Then
				Set masterlist = New MasterList_Reports 
				Set masterlist.Individual  = New Individual
				masterlist.Individual.Generate_Report
				HMIRuntime.Tags("individual_report_noon").Write 1
			End If
			
		End If
	
	

	
'Description : Closed MySQL Object Connection	
'---------------------------------------------------------------------------	
	Call Mysql_Close_Conn(objRS,objConn)
	
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "04_Analysis.bmo - Sub Analysis_Q is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
	End If

End Sub