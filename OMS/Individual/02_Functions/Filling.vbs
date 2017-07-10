'******************************************************************************************************
' Description  : Class Filling Individual (Action Scripts : FillingIndividual_Queue.bac)
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************

Class Filling_Individual
    Private rack_name , rack_no, prod_detail, uid, isPrefilled, dbRack_No, arr_result(10)
	Private prod_qty(5), prod_recipe(5), prod_cyl(5), qty , shift_batch

'Description : Initial constructor
'----------------------------------------------------------------------------------------------------
	
    Public Default Function Init(parameters)
         Select Case UBound(parameters)
             Case 5
             	Set Init = InitFiveParam(parameters(0), parameters(1), parameters(2), parameters(3), parameters(4),parameters(5))
			 Case 6
             	Set Init = InitSixParam(parameters(0), parameters(1), parameters(2), parameters(3), parameters(4),parameters(5),parameters(6))
             Case Else
                Set Init = Me
         End Select
    End Function
 
'Description : Function InitFiveParam [to assign parameter value into variable]
'---------------------------------------------------------------------------------------------------- 
    Private Function InitFiveParam(parameter1, parameter2, parameter3, parameter4, parameter5,parameter6)
    	 
    	rack_name 		 = parameter1
        rack_no 		 = parameter2
        prod_detail		 = parameter3
        uid				 = parameter4
		isPrefilled      = parameter5
        dbRack_No        = parameter6
       
        Call product_filter
        
        Set InitFiveParam = Me
    End Function
	
	Private Function InitSixParam(parameter1, parameter2, parameter3, parameter4, parameter5,parameter6,parameter7)
    	 
    	rack_name 		 = parameter1
        rack_no 		 = parameter2
        prod_detail		 = parameter3
        uid				 = parameter4
		isPrefilled      = parameter5
        dbRack_No        = parameter6
        shift_batch		 = parameter7
		
        Call product_filter
        
        Set InitFiveParam = Me
    End Function

'Description : Function product_filter [to filter the product name before sent to enertech plc]
'---------------------------------------------------------------------------------------------------- 	
	Public Function product_filter()
		
		Dim re : Set re = New RegExp
		re.Global = True
		re.Pattern = "[a-zA-Z_]"

		Dim long_str , mid_str , low_str
		Dim del_1 , del_2 , del_3
	
		Dim filter_1(5)
		Dim filter_2(5)
		
		prod_detail = Replace(prod_detail,",NULL|0","")

		long_str=Split(prod_detail,",")
		del_1= ubound(long_str)
		qty = del_1
		
		For i=0 To qty   
			filter_1(i) = long_str(i)
		Next

		For i = 0 To qty
			mid_str = Split(filter_1(i),"|")
			del_2 = ubound(Mid_str)

			For j = 0 to del_2
				prod_qty(i) = mid_str(1)
				filter_2(i) = Mid_str(0)
			Next
		Next

		For i = 0 To qty
			low_str = Split(filter_2(i), "-")
			del_3 = ubound(low_str)

			For j = 0 To del_3
				If del_3 = 2 Then
					prod_recipe(i) = low_str(0) 
					prod_recipe(i) = re.Replace(prod_recipe(i),"")
					prod_cyl(i) = low_str(2) & "-" & low_str(1) 
				Else
					prod_recipe(i) = low_str(0)
					prod_recipe(i) = re.Replace(prod_recipe(i),"")
					prod_cyl(i) = low_str(1)	
				End If	
			Next
		Next

	End Function

'Description : Function Filling_Idividual_Tag [to assign internal/plc tag ]
'---------------------------------------------------------------------------------------------------- 	
	Private Function Filling_Idividual_Tag(Byval index, Byval DB_name) 

		If DB_name = "DB360_Order_Status" Then	'Datablock 360 / Datablock 330
				Dim DB360_Order_Status
					DB360_Order_Status 		= Array("RESERVED SPACE" , _
											   		"PLC_01/fill_ind_360.CO2_R1_Order_Status" , _
											   		"PLC_01/fill_ind_360.CO2_R2_Order_Status" , _
											   		"PLC_01/fill_ind_360.CO2_R3_Order_Status" , _
											   		"PLC_01/fill_ind_360.N2O_R1_Order_Status" , _
											   		"PLC_01/fill_ind_360.N2O_R2_Order_Status" , _
											   		"PLC_01/fill_ind_360.N2O_R3_Order_Status" , _
													"PLC_01/fill_ind_360.EOX_1_Order_Status" , _
													"PLC_01/fill_ind_360.EOX_2_Order_Status" , _
													"PLC_01/fill_ind_360.EOX_3_Order_Status" , _
													"PLC_01/CrySt_DB1.FS4_2_CO2_WO_Status" , _
													"PLC_01/CrySt_DB1.FS4_3_CO2_WO_Status" , _
											   		"PLC_01/CrySt_DB1.FS4_4_CO2_WO_Status") 
					Filling_Idividual_Tag = DB360_Order_Status(index)		
		
		Elseif DB_name = "DB360_Recipe_Code" Then	'Datablock 360 / Datablock 330
					Dim DB360_Recipe_Code
						DB360_Recipe_Code 	= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_360.CO2_R1_Recipe_Code" , _
													"PLC_01/fill_ind_360.CO2_R2_Recipe_Code" , _
													"PLC_01/fill_ind_360.CO2_R3_Recipe_Code" , _
													"PLC_01/fill_ind_360.N2O_R1_Recipe_Code" , _
													"PLC_01/fill_ind_360.N2O_R2_Recipe_Code" , _
													"PLC_01/fill_ind_360.N2O_R3_Recipe_Code" , _
													"PLC_01/fill_ind_360.EOX_1_Recipe_Code" , _
													"PLC_01/fill_ind_360.EOX_2_Recipe_Code" , _
													"PLC_01/fill_ind_360.EOX_3_Recipe_Code" , _
													"PLC_01/CrySt_DB1.FS4_2_CO2_Recipe" , _
													"PLC_01/CrySt_DB1.FS4_3_CO2_Recipe" , _
													"PLC_01/CrySt_DB1.FS4_4_CO2_Recipe") 
					Filling_Idividual_Tag = DB360_Recipe_Code(index)		

		Elseif DB_name = "DB360_Cyl_Code" Then	'Datablock 360 / Datablock 330
					Dim DB360_Cyl_Code
						DB360_Cyl_Code 	= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_360.CO2_R1_Cyl_Code" , _
													"PLC_01/fill_ind_360.CO2_R2_Cyl_Code" , _
													"PLC_01/fill_ind_360.CO2_R3_Cyl_Code" , _
													"PLC_01/fill_ind_360.N2O_R1_Cyl_Code" , _
													"PLC_01/fill_ind_360.N2O_R2_Cyl_Code" , _
													"PLC_01/fill_ind_360.N2O_R3_Cyl_Code" , _
													"PLC_01/fill_ind_360.EOX_1_Cyl_Code" , _
													"PLC_01/fill_ind_360.EOX_2_Cyl_Code" , _
													"PLC_01/fill_ind_360.EOX_3_Cyl_Code" , _
													"PLC_01/CrySt_DB1.FS4_2_CO2_Cyl_Code" , _
													"PLC_01/CrySt_DB1.FS4_3_CO2_Cyl_Code" , _
													"PLC_01/CrySt_DB1.FS4_4_CO2_Cyl_Code") 
					Filling_Idividual_Tag = DB360_Cyl_Code(index)

		Elseif DB_name = "DB360_Cyl_Qty" Then	'Datablock 360 / Datablock 330
					Dim DB360_Cyl_Qty
						DB360_Cyl_Qty 	= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_360.CO2_R1_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.CO2_R2_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.CO2_R3_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.N2O_R1_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.N2O_R2_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.N2O_R3_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.EOX_1_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.EOX_2_Cyl_Quantity" , _
													"PLC_01/fill_ind_360.EOX_3_Cyl_Quantity" , _
													"PLC_01/CrySt_DB1.FS4_2_CO2_Cyl_Quant" , _
													"PLC_01/CrySt_DB1.FS4_3_CO2_Cyl_Quant" , _
													"PLC_01/CrySt_DB1.FS4_4_CO2_Cyl_Quant") 
					Filling_Idividual_Tag = DB360_Cyl_Qty(index)
					
		Elseif DB_name = "DB360_Fill_Stat" Then	'Datablock 360
					Dim DB360_Fill_Stat
						DB360_Fill_Stat 	= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_360.CO2_R1_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.CO2_R2_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.CO2_R3_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.N2O_R1_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.N2O_R2_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.N2O_R3_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.EOX_1_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.EOX_2_Cyl_Fill_Stat" , _
													"PLC_01/fill_ind_360.EOX_3_Cyl_Fill_Stat") 
					Filling_Idividual_Tag = DB360_Fill_Stat(index)
		
							
		Elseif DB_name = "DB361_Status" Then	'Datablock 361 / Datablock 330	
				Dim DB361_Status
					DB361_Status 			= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_361.CO2_R1_Status" , _
													"PLC_01/fill_ind_361.CO2_R2_Status" , _
													"PLC_01/fill_ind_361.CO2_R3_Status" , _
													"PLC_01/fill_ind_361.N2O_R1_Status" , _
													"PLC_01/fill_ind_361.N2O_R2_Status" , _
													"PLC_01/fill_ind_361.N2O_R3_Status" , _
													"PLC_01/fill_ind_361.EOX_1_Status" , _
													"PLC_01/fill_ind_361.EOX_2_Status" , _
													"PLC_01/fill_ind_361.EOX_3_Status" , _
													"PLC_01/CrySt_DB2.FS4_2_CO2_Status" , _
													"PLC_01/CrySt_DB2.FS4_3_CO2_Status" , _
													"PLC_01/CrySt_DB2.FS4_4_CO2_Status")
					Filling_Idividual_Tag = DB361_Status(index)
		
		ElseIf DB_name = "DB362_Filling_Status" Then	'Datablock 362	
				Dim DB362_Filling_Status 
					DB362_Filling_Status 	= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_362.CO2_R1_Stat" , _
													"PLC_01/fill_ind_362.CO2_R2_Stat" , _
													"PLC_01/fill_ind_362.CO2_R3_Stat" , _
													"PLC_01/fill_ind_362.N2O_R1_Stat" , _
													"PLC_01/fill_ind_362.N2O_R2_Stat" , _
													"PLC_01/fill_ind_362.N2O_R3_Stat" , _
													"PLC_01/fill_ind_362.EOX_1_Stat" , _
													"PLC_01/fill_ind_362.EOX_2_Stat" , _
													"PLC_01/fill_ind_362.EOX_3_Stat")
					Filling_Idividual_Tag = DB362_Filling_Status(index)
		
		ElseIf DB_name = "DB362_Filling_Weight_M" Then	'Datablock 362	
				Dim DB362_Filling_Weight_M 
					DB362_Filling_Weight_M 	= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_362.CO2_R1_Weight_M" , _
													"PLC_01/fill_ind_362.CO2_R2_Weight_M" , _
													"PLC_01/fill_ind_362.CO2_R3_Weight_M" , _
													"PLC_01/fill_ind_362.N2O_R1_Weight_M" , _
													"PLC_01/fill_ind_362.N2O_R2_Weight_M" , _
													"PLC_01/fill_ind_362.N2O_R3_Weight_M" , _
													"PLC_01/fill_ind_362.EOX_1_Weight_M" , _
													"PLC_01/fill_ind_362.EOX_2_Weight_M" , _
													"PLC_01/fill_ind_362.EOX_3_Weight_M")
					Filling_Idividual_Tag = DB362_Filling_Weight_M(index)
		
		ElseIf DB_name = "DB362_Filling_Result" Then	'Datablock 362	
				Dim DB362_Filling_Result 
					DB362_Filling_Result 	= Array("RESERVED SPACE" , _
													"PLC_01/fill_ind_362.CO2_R1_Result" , _
													"PLC_01/fill_ind_362.CO2_R2_Result" , _
													"PLC_01/fill_ind_362.CO2_R3_Result" , _
													"PLC_01/fill_ind_362.N2O_R1_Result" , _
													"PLC_01/fill_ind_362.N2O_R2_Result" , _
													"PLC_01/fill_ind_362.N2O_R3_Result" , _
													"PLC_01/fill_ind_362.EOX_1_Result" , _
													"PLC_01/fill_ind_362.EOX_2_Result" , _
													"PLC_01/fill_ind_362.EOX_3_Result")
					Filling_Idividual_Tag = DB362_Filling_Result(index)		

		Elseif DB_name = "S_dateTime" Then
				Dim S_dateTime	
					S_dateTime 		= Array("RESERVED SPACE" , _
										"S_F_Individual_DateTime_1" ,  _
										"S_F_Individual_DateTime_2" ,  _
										"S_F_Individual_DateTime_3" ,  _
										"S_F_Individual_DateTime_4" ,  _
										"S_F_Individual_DateTime_5" ,  _
										"S_F_Individual_DateTime_6" ,  _
										"S_F_Individual_DateTime_7" ,  _
										"S_F_Individual_DateTime_8" ,  _
										"S_F_Individual_DateTime_9" ,  _
										"S_F_Individual_DateTime_10" ,  _
										"S_F_Individual_DateTime_11" ,  _
										"S_F_Individual_DateTime_12")
					Filling_Idividual_Tag = S_dateTime(index)
		
		Elseif DB_name = "E_dateTime" Then
				Dim E_dateTime	
					E_dateTime 		= Array("RESERVED SPACE" , _
										"E_F_Individual_DateTime_1" ,  _
										"E_F_Individual_DateTime_2" ,  _
										"E_F_Individual_DateTime_3" ,  _
										"E_F_Individual_DateTime_4" ,  _
										"E_F_Individual_DateTime_5" ,  _
										"E_F_Individual_DateTime_6" ,  _
										"E_F_Individual_DateTime_7" ,  _
										"E_F_Individual_DateTime_8" ,  _
										"E_F_Individual_DateTime_9" ,  _
										"E_F_Individual_DateTime_10" ,  _
										"E_F_Individual_DateTime_11" ,  _
										"E_F_Individual_DateTime_12")
					Filling_Idividual_Tag = E_dateTime(index)
					
		Elseif DB_name = "Cancel_Time" Then
				Dim Cancel_Time	
					Cancel_Time 		= Array("RESERVED SPACE" , _
										"Cancel_Trigger_Filling_Individual_Time_1" ,  _
										"Cancel_Trigger_Filling_Individual_Time_2" ,  _
										"Cancel_Trigger_Filling_Individual_Time_3" ,  _
										"Cancel_Trigger_Filling_Individual_Time_4" ,  _
										"Cancel_Trigger_Filling_Individual_Time_5" ,  _
										"Cancel_Trigger_Filling_Individual_Time_6" ,  _
										"Cancel_Trigger_Filling_Individual_Time_7" ,  _
										"Cancel_Trigger_Filling_Individual_Time_8" ,  _
										"Cancel_Trigger_Filling_Individual_Time_9" ,  _
										"Cancel_Trigger_Filling_Individual_Time_10" ,  _
										"Cancel_Trigger_Filling_Individual_Time_11" ,  _
										"Cancel_Trigger_Filling_Individual_Time_12")
					Filling_Idividual_Tag = Cancel_Time(index)
					
		Elseif DB_name = "Cancel_Trigger" Then
				Dim Cancel_Trigger	
					Cancel_Trigger 		= Array("RESERVED SPACE" , _
										"Cancel_Trigger_Filling_Individual_Bool_1" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_2" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_3" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_4" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_5" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_6" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_7" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_8" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_9" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_10" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_11" ,  _
										"Cancel_Trigger_Filling_Individual_Bool_12")
					Filling_Idividual_Tag = Cancel_Trigger(index)
					
		Elseif DB_name = "Individual_QI_Full" Then
				Dim Individual_QI_Full	
					Individual_QI_Full 	= Array("RESERVED SPACE" , _
										"Individual_QI_Full_1" ,  _
										"Individual_QI_Full_2" ,  _
										"Individual_QI_Full_3" ,  _
										"Individual_QI_Full_4" ,  _
										"Individual_QI_Full_5" ,  _
										"Individual_QI_Full_6" ,  _
										"Individual_QI_Full_7" ,  _
										"Individual_QI_Full_8" ,  _
										"Individual_QI_Full_9" ,  _
										"Individual_QI_Full_10" ,  _
										"Individual_QI_Full_11" ,  _
										"Individual_QI_Full_12")
					Filling_Idividual_Tag = Individual_QI_Full(index)
					
					
	'developement in progress	
		End If
	End Function
	
'Description : Let Property for Analysis [to set the value property of objects]
'---------------------------------------------------------------------------------------------------- 
Dim Tags 
	Public Property Let order_status(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"DB360_Order_Status") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let recipe_code(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"DB360_Recipe_Code") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let cyl_code(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"DB360_Cyl_Code") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let quantity(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"DB360_Cyl_Qty") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let cyl_fill_state(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"DB360_Fill_Stat") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let capture_start_timestamp(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"S_dateTime") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let capture_end_timestamp(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"E_dateTime") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	Public Property Let Individual_QI_State(ByVal value)
		Tags = Filling_Idividual_Tag(rack_no,"Individual_QI_Full") 
		HMIRuntime.Tags(Tags).Write value
	End Property
	
	
	
'Description : Get Property for Analysis [to get the value property of objects]
'---------------------------------------------------------------------------------------------------- 
	Public Property Get work_status
		
		Tags = Filling_Idividual_Tag(rack_no,"DB361_Status") 
		work_status= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get fill_status
		Tags = Filling_Idividual_Tag(rack_no,"DB362_Filling_Status") 
		fill_status= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get fill_weight
		Tags = Filling_Idividual_Tag(rack_no,"DB362_Filling_Weight_M") 
		fill_weight= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get fill_result
		Tags = Filling_Idividual_Tag(rack_no,"DB362_Filling_Result") 
		fill_result= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get fill_start_time
		Tags = Filling_Idividual_Tag(rack_no,"S_dateTime") 
		fill_start_time= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get fill_end_time
		Tags = Filling_Idividual_Tag(rack_no,"E_dateTime") 
		fill_end_time= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get Individual_QI_State
		Tags = Filling_Idividual_Tag(rack_no,"Individual_QI_Full") 
		Individual_QI_State= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get capture_start_timestamp
		Tags = Filling_Idividual_Tag(rack_no,"S_dateTime") 
		capture_start_timestamp = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get capture_end_timestamp
		Tags = Filling_Idividual_Tag(rack_no,"E_dateTime") 
		capture_end_timestamp = HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Property Get get_rack_no
		get_rack_no = rack_no
	End Property
	
	Public Property Get get_rack_name
		get_rack_name = rack_name
	End Property
	
'Description : Function Sent the oms value into Enertech PLC [Action Script Phase 1] 
'---------------------------------------------------------------------------------------------------- 
	Public Function SentValue_Medical_Plc()
			order_status = 1
			recipe_code = prod_recipe(0)
			cyl_code = prod_cyl(0)
			quantity = prod_qty(0)
			cyl_fill_state = isPrefilled
			
			Call GF_LogToFile_("Exec", "Sent Into EnerTech PLC ","Individual")
			Call GF_LogToFile_("Exec", "Set Status = 1 ","Individual")
			Call GF_LogToFile_("Exec", "[ Recipe = " & prod_recipe(0) & ", Cylinder = " & prod_cyl(0) & ", Product = " & prod_qty(0) & ", isPrefilled = " & isPrefilled & "]" ,"Individual")

	End Function
	
'Description : Function Sent the oms value into Cryostar PLC [Action Script Phase 1] 
'---------------------------------------------------------------------------------------------------- 
	Public Function SentValue_Industry_Plc()
			order_status = 1
			recipe_code = prod_recipe(0)
			cyl_code = prod_cyl(0)
			quantity = prod_qty(0)
			
			Call GF_LogToFile_("Execute", "Sent Into Cryostar PLC ","Individual")
			Call GF_LogToFile_("Execute", "Set Status = 1 ","Individual")
			Call GF_LogToFile_("Execute", "[ Recipe = " & prod_recipe(0) & ", Cylinder = " & prod_cyl(0) & " , Quantity = " & prod_qty(0) & "]" ,"Individual")

	End Function

'Description : Function to check the QI/Full Status [Action Script Phase 3] 
'---------------------------------------------------------------------------------------------------- 	
	Public Function analysis_required()
		If prod_recipe(0) = 211 Then
			If isPrefilled = 0 Then
				Individual_QI_State = 1
			Elseif isPrefilled = 1 Then
				Individual_QI_State = 2
			End If
		Else
			Individual_QI_State = 2
		End If
		
	End Function
	
'Description : Function Deactivate Check Point 1 [Action Script Phase 1]
'---------------------------------------------------------------------------------------------------- 
	Public Function Deactivate_CheckingPoint_1()
			order_status = 0
			Call Mysql_Non_Query("Update codabix_trigger Set db_360 = 0 , db_361 =1 , db_362 = 0 Where state = 1 and rack_id = "& dbRack_No & "")
			
			Call GF_LogToFile_("Execute", "Set Status = 0 ", "Individual")
			Call GF_LogToFile_("Execute", "EnerTech Deactivate CP 1 : " & dbRack_No , "Individual")
	End Function
	
'Description : Function Deactivate Check Point 1 [Action Script Phase 1]
'---------------------------------------------------------------------------------------------------- 
	Public Function Deactivate_Cryostar_CheckingPoint_1()
			order_status = 0
			Call Mysql_Non_Query("Update codabix_trigger Set db_330 = 0 , db_331 =1  Where state = 1 and rack_id = "& dbRack_No & "")
	
			Call GF_LogToFile_("Execute", "Set Status = 0 ", "Individual")
			Call GF_LogToFile_("Execute", "Cryostar Deactivate CP 1 : " & dbRack_No , "Individual")
	End Function

'Description : Function Deactivate Check Point 2 [Action Script Phase 2]
'---------------------------------------------------------------------------------------------------- 
	Public Function Deactivate_CheckingPoint_2()
			order_status = 0
			Call Mysql_Non_Query("Update codabix_trigger Set db_360 = 0 , db_361 = 0 , db_362 = 1 Where state = 1 and rack_id = "& dbRack_No & "")
			Call GF_LogToFile_("Execute", "EnerTech Deactivate CP 2 : " & dbRack_No , "Individual")
	End Function
	
'Description : Function Deactivate Check Point 2 [Action Script Phase 2]
'---------------------------------------------------------------------------------------------------- 
	Public Function Deactivate_Cryostar_CheckingPoint_2()
			order_status = 0
			Call Mysql_Non_Query("Update codabix_trigger Set db_330 = 0 , db_331 = 0  Where state = 1 and rack_id = "& dbRack_No & "")
			Call Deactivate_Cryostar_All()
	End Function
	
'Description : Function Reset Filling(12) EnerTech
'----------------------------------------------------------------------------------------------------  
	Public Function Reset_FillIndividual_EnerTech
		Call Mysql_Non_Query("Update codabix_trigger Set db_360 = 0 , db_361 = 0 , db_362 = 0  , state = 0 , cylinder_id = 0 , oms_batch = '' , shift_batch = '' , user_id = 0 , prod_detail = '' , IsPrefilled = 0 Where state = 1 and rack_id = "& dbRack_No & "")
		Call Mysql_Non_Query("Update filling_rack Set occupied = 0 , cylinder_type = '' , user_id = ''  Where rack_id="& dbRack_No & "")
		Call GF_LogToFile_("RESET", " Filling EnerTech : " & dbRack_No , "Individual")
	End Function
	
'Description : Function Reset Filling(12) EnerTech
'----------------------------------------------------------------------------------------------------  
	Public Function Reset_FillIndividual_Cryostar
		Call Mysql_Non_Query("Update codabix_trigger Set db_330 = 0 , db_331 = 0  , state = 0 , cylinder_id = 0 , oms_batch = '' , shift_batch = '' ,user_id = 0 , prod_detail = '' , IsPrefilled = 0 Where state = 1 and rack_id = "& dbRack_No & "")
		Call Mysql_Non_Query("Update filling_rack Set occupied = 0 , cylinder_type = '' , user_id = ''  Where rack_id="& dbRack_No & "")
		
		Call GF_LogToFile_("RESET", " Filling Cryostar : " & dbRack_No , "Individual")
	End Function
	
'Description : Function Deactivate Check Point 3 [Action Script Phase 3]
'---------------------------------------------------------------------------------------------------- 	
	Public Function Deactivate_CheckingPoint_3()
	
			Call Mysql_Non_Query("Update codabix_trigger Set db_360 = 0 , db_361 = 0 , db_362 = 0 Where state = 1 and rack_id = "& dbRack_No & "")
			Call StoreToDB()
			Call Deactivate_All()
			Call GF_LogToFile_("Execute", " Deactivate CP 3 : " & dbRack_No , "Individual")
	End Function
	
'Description : Function Deactivate Check Point 3 [Action Script Phase 3]
'---------------------------------------------------------------------------------------------------- 	
	Public Function Deactivate_All()
			Call Mysql_Non_Query("Update codabix_trigger Set db_360 = 0 , db_361 = 0 , db_362 = 0  , state = 0 , cylinder_id = 0 , oms_batch = '' , shift_batch = '' ,user_id = 0 , prod_detail = '' , IsPrefilled = 0 Where state = 1 and rack_id = "& dbRack_No & "")
			Call Mysql_Non_Query("Update filling_rack Set occupied = 0 , cylinder_type = '' , user_id = ''  Where rack_id="& dbRack_No & "")
			Call GF_LogToFile_("Execute", " Deactivate All CheckPoint : " & dbRack_No , "Individual")
			Call ClearPLC_Ind_EnerTech()
	End Function
	
'Description : Clear all Internal Tag Values
'---------------------------------------------------------------------------------------------------
	Public Function ClearIndividualInternalTag()
		capture_start_timestamp = ""
		capture_end_timestamp = ""
		Individual_QI_State = ""
		
	End Function
	
'Description : Function Deactivate Check Point 3 [Action Script Phase 3]
'---------------------------------------------------------------------------------------------------- 	
	Public Function Deactivate_Cryostar_All()
			Dim filling_end_time_fromOS : filling_end_time_fromOS = DisplayDate(Now)
			
			Call Mysql_Non_Query("Insert INTO sub_fill_individual ( user_id , user_entry_start_date , user_entry_end_date , oms_batch ,shift_batch, cylinder_id , pallet_table_id ," & _
				" weight1 ,weight2 ,weight3, weight1_starttime , weight1_endtime ,weight2_starttime, weight2_endtime, weight3_starttime , weight3_endtime , result1 , result2 , result3 )"  & _
				" SELECT '"& uid &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , c.oms_batch ,c.shift_batch, c.cylinder_id , p.id ," & _ 
				" '0' , '0','0', '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , '-' , '-' , '-' , '-' , "& _
				" 0 , 0 , 0   FROM codabix_trigger c  JOIN pallet_table p on (p.cyl_id = c.cylinder_id) and p.cyl_state = 4  and p.oms_batch = c.oms_batch WHERE c.rack_id="& dbRack_No &"")
				
			Call Mysql_Non_Query("Update pallet_table Set filling_finish = 1 , analysis_mode = 1 ,shift_batch = '" & shift_batch & "', analysis_required = 1 , fill_mode = 1 Where cyl_id =(SELECT distinct cylinder_id from codabix_trigger WHERE rack_id = "& dbRack_No & ") and oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& dbRack_No & ") ")
					
				
			Call Mysql_Non_Query("Update codabix_trigger Set db_330 = 0 , db_331 = 0  , state = 0 , cylinder_id = 0 , oms_batch = '' ,shift_batch = '', user_id = 0 , prod_detail = '' , IsPrefilled = 0 Where state = 1 and rack_id = "& dbRack_No & "")
			Call Mysql_Non_Query("Update filling_rack Set occupied = 0 , cylinder_type = '' , user_id = ''  Where rack_id="& dbRack_No & "")
			Call ClearPLC_Ind_Cryostar()
			
	End Function

'Description : Function StoretoDB [Action Script Phase 3]
'---------------------------------------------------------------------------------------------------- 	
	Public Function StoreToDB()
		'1 = i , 2 = p , 3 = l [analysis_mode]
		'0 = No/prefilled , 1 = QI/FULL   [analysis_required] cylinder QI/FULL
		'1 = individual_sub , 2= fill_loos_sub , 3 = fill_mcp_sub , 4 = fill_cdp_sub  [fill_mode] cylinder QI/FULL
		Dim filling_end_time_fromOS : filling_end_time_fromOS = DisplayDate(Now)
		
		
		
		If prod_recipe(0) = 211 Then
			
			If Individual_QI_State = 1 Then
			
				If Left(get_rack_name,3) = "N2O" Then
				
					Call Mysql_Non_Query("Insert INTO sub_fill_individual ( user_id , user_entry_start_date , user_entry_end_date , oms_batch ,shift_batch, cylinder_id , pallet_table_id ," & _
					" weight2 , weight2_starttime , weight2_endtime , result2 )"  & _
					" SELECT '"& uid &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , c.oms_batch ,c.shift_batch, c.cylinder_id , p.id ," & _ 
					"'" & fill_weight &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' ,"& _
					"" & fill_result &"   FROM codabix_trigger c  JOIN pallet_table p on (p.cyl_id = c.cylinder_id) and p.cyl_state = 3  and p.oms_batch = c.oms_batch WHERE c.rack_id="& dbRack_No &"")
					
					
				Elseif Left(get_rack_name,3) = "EOX" Then
					
					Call Mysql_Non_Query("Insert INTO sub_fill_individual ( user_id , user_entry_start_date , user_entry_end_date , oms_batch ,shift_batch, cylinder_id , pallet_table_id ," & _
					" weight3 , weight3_starttime , weight3_endtime , result3 )"  & _
					" SELECT '"& uid &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , c.oms_batch ,c.shift_batch, c.cylinder_id , p.id ," & _ 
					"'" & fill_weight &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' ,"& _
					"" & fill_result &"   FROM codabix_trigger c  JOIN pallet_table p on (p.cyl_id = c.cylinder_id) and p.cyl_state = 3  and p.oms_batch = c.oms_batch WHERE c.rack_id="& dbRack_No &"")
				Else
					
					Call Mysql_Non_Query("Insert INTO sub_fill_individual ( user_id , user_entry_start_date , user_entry_end_date , oms_batch ,shift_batch, cylinder_id , pallet_table_id ," & _
					" weight1 , weight1_starttime , weight1_endtime , result1 )"  & _
					" SELECT '"& uid &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , c.oms_batch ,c.shift_batch, c.cylinder_id , p.id ," & _ 
					"'" & fill_weight &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' ,"& _
					"" & fill_result &"   FROM codabix_trigger c  JOIN pallet_table p on (p.cyl_id = c.cylinder_id) and p.cyl_state = 3  and p.oms_batch = c.oms_batch WHERE c.rack_id="& dbRack_No &"")
					
				End If
				
				Call Mysql_Non_Query("Update pallet_table Set filling_finish = 0 , analysis_mode = 1 ,analysis_required = 0 ,shift_batch = '" & shift_batch & "', fill_mode = 1 Where cyl_id =(SELECT distinct cylinder_id from codabix_trigger WHERE rack_id = "& dbRack_No & ") and oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& dbRack_No & ") ")
				
				
			Elseif Individual_QI_State = 2 Then
				If Left(get_rack_name,3) = "N2O" Then
					Call Mysql_Non_Query("Update sub_fill_individual Set weight2 = '"& fill_weight &"' , weight2_starttime = '"& fill_start_time &"' , weight2_endtime = '" & filling_end_time_fromOS & "' , result2 = "& fill_result &"  Where cylinder_id=(SELECT distinct cylinder_id  from codabix_trigger WHERE rack_id = "& dbRack_No & ") and oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& dbRack_No & ") ")
				
				Elseif Left(get_rack_name,3) = "EOX" Then
					Call Mysql_Non_Query("Update sub_fill_individual Set weight3 = '"& fill_weight &"' , weight3_starttime = '"& fill_start_time &"' , weight3_endtime = '" & filling_end_time_fromOS & "' , result3 = "& fill_result &"  Where cylinder_id=(SELECT distinct cylinder_id  from codabix_trigger WHERE rack_id = "& dbRack_No & ") and oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& dbRack_No & ") ")
				Else
					Call Mysql_Non_Query("Update sub_fill_individual Set weight1 = '"& fill_weight &"' , weight1_starttime = '"& fill_start_time &"' , weight1_endtime = '" & filling_end_time_fromOS & "' , result1 = "& fill_result &"  Where cylinder_id=(SELECT distinct cylinder_id  from codabix_trigger WHERE rack_id = "& dbRack_No & ") and oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& dbRack_No & ") ")
					
				End If
				Call Mysql_Non_Query("Update pallet_table Set filling_finish = 1 , analysis_mode = 1 ,analysis_required = 1 ,shift_batch = '" & shift_batch & "', fill_mode = 1 Where cyl_id =(SELECT distinct cylinder_id from codabix_trigger WHERE rack_id = "& dbRack_No & ") and oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& dbRack_No & ") ")
				
			End If
		
		Else
			If Left(get_rack_name,3) = "N2O" Then
				Call Mysql_Non_Query("Insert INTO sub_fill_individual ( user_id , user_entry_start_date , user_entry_end_date , oms_batch , shift_batch , cylinder_id , pallet_table_id ," & _
				" weight1 ,weight2 ,weight3, weight1_starttime , weight1_endtime ,weight2_starttime, weight2_endtime, weight3_starttime , weight3_endtime , result1 , result2 , result3 )"  & _
				" SELECT '"& uid &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , c.oms_batch ,c.shift_batch, c.cylinder_id , p.id ," & _ 
				" '0' , '" & fill_weight &"' , '0' , '-','-', '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , '-' , '-' ,"& _
				" 0 , " & fill_result &" , 0   FROM codabix_trigger c  JOIN pallet_table p on (p.cyl_id = c.cylinder_id) and p.cyl_state = 4  and p.oms_batch = c.oms_batch WHERE c.rack_id="& dbRack_No &"")
				
				
			Elseif Left(get_rack_name,3) = "EOX" Then
				Call Mysql_Non_Query("Insert INTO sub_fill_individual ( user_id , user_entry_start_date , user_entry_end_date , oms_batch ,shift_batch, cylinder_id , pallet_table_id ," & _
				" weight1 ,weight2 ,weight3, weight1_starttime , weight1_endtime ,weight2_starttime, weight2_endtime, weight3_starttime , weight3_endtime , result1 , result2 , result3 )"  & _
				" SELECT '"& uid &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , c.oms_batch ,c.shift_batch, c.cylinder_id , p.id ," & _ 
				" 0 ,0 ,'" & fill_weight &"' , '-' , '-' , '-' , '-' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' ,"& _
				" 0 , 0 , " & fill_result &"    FROM codabix_trigger c  JOIN pallet_table p on (p.cyl_id = c.cylinder_id) and p.cyl_state = 4  and p.oms_batch = c.oms_batch WHERE c.rack_id="& dbRack_No &"")
				
				
			Else
				Call Mysql_Non_Query("Insert INTO sub_fill_individual ( user_id , user_entry_start_date , user_entry_end_date , oms_batch ,shift_batch, cylinder_id , pallet_table_id ," & _
				" weight1 ,weight2 ,weight3, weight1_starttime , weight1_endtime ,weight2_starttime, weight2_endtime, weight3_starttime , weight3_endtime , result1 , result2 , result3 )"  & _
				" SELECT '"& uid &"' , '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , c.oms_batch ,c.shift_batch, c.cylinder_id , p.id ," & _ 
				"'" & fill_weight &"' , '0','0', '"& fill_start_time &"' , '" & filling_end_time_fromOS & "' , '-' , '-' , '-' , '-' , "& _
				"" & fill_result &" , 0 , 0   FROM codabix_trigger c  JOIN pallet_table p on (p.cyl_id = c.cylinder_id) and p.cyl_state = 4  and p.oms_batch = c.oms_batch WHERE c.rack_id="& dbRack_No &"")
				
				
			End If
				
				Call Mysql_Non_Query("Update pallet_table Set filling_finish = 1 , analysis_mode = 1 ,analysis_required = 1 ,shift_batch = '" & shift_batch & "', fill_mode = 1 Where cyl_id =(SELECT distinct cylinder_id from codabix_trigger WHERE rack_id = "& dbRack_No & ") and oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& dbRack_No & ") ")
				
		End If

	End Function

'Description : Clear PLC Tags
'----------------------------------------------------------------------------------------------------	
	Public Function ClearPLC_Ind_EnerTech()
		order_status   = 0
		recipe_code    = ""
		cyl_code       = ""
		quantity       = 0
		cyl_fill_state = 0
	End Function
	
	Public Function ClearPLC_Ind_Cryostar()
		order_status  = 0
		recipe_code   = ""
		cyl_code      = ""
		quantity      = 0
	End Function
	
'Description : Others Method
'---------------------------------------------------------------------------------------------------- 
	Public Function Test()
		Test = rack_no
	End Function

	
End Class