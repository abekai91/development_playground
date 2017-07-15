'******************************************************************************************************
' Description  : Class Filling Loose (Action Scripts : )
' Author by    : Ahmad Syazwan
' Modified Date: -
' Created Date : 18 Feburary 2017
'******************************************************************************************************
Class Filling
    Private rack_name , rack_no , prod_detail , uid , qty
	Private arr_result(10) , prod_qty(5) , prod_recipe(5) , prod_cyl(5) 

'Description : Initial constructor
'----------------------------------------------------------------------------------------------------
    Public Default Function Init(parameters)
	On Error Resume Next
         Select Case UBound(parameters)
             Case 0
                Set Init = InitOneParam(parameters(0))
             Case 1
                Set Init = InitTwoParam(parameters(0), parameters(1))
             Case 2 
             	Set Init = InitThreeParam(parameters(0), parameters(1), parameters(2))
             Case 3 
             	Set Init = InitFourParam(parameters(0), parameters(1), parameters(2), parameters(3))
             Case Else
                Set Init = Me
         End Select
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function Init is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
    End Function

'Description : InitOneParam [assign one parameters]
'----------------------------------------------------------------------------------------------------
    Private Function InitOneParam(parameter1)
	On Error Resume Next
        If TypeName(parameter1) = "String" Then
            rack_name = parameter1
        Else
            rack_no = parameter1
        End If
        Set InitOneParam = Me
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function InitOneParam is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
    End Function

'Description : InitTwoParam [assign two parameters]
'----------------------------------------------------------------------------------------------------
    Private Function InitTwoParam(parameter1, parameter2)
	On Error Resume Next
        rack_name 		 = parameter1
        rack_no 		 = parameter2
        Set InitTwoParam = Me
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function InitTwoParam is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
    End Function

'Description : InitThreeParam [assign three parameters]
'----------------------------------------------------------------------------------------------------	
	Private Function InitThreeParam(parameter1, parameter2, parameter3)
	On Error Resume Next
    	rack_name 		 = parameter1
        rack_no 		 = parameter2
        prod_detail		 = parameter3
        Call product_filter
        Set InitThreeParam = Me
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function InitThreeParam is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
    End Function

'Description : InitFourParam [assign four parameters]
'----------------------------------------------------------------------------------------------------    
    Private Function InitFourParam(parameter1, parameter2, parameter3, parameter4)
	On Error Resume Next
    	rack_name 		 = parameter1
        rack_no 		 = parameter2
        prod_detail		 = parameter3
        uid				 = parameter4
        Call product_filter
        Set InitFourParam = Me
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function InitFourParam is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
    End Function

'Description : PLC_Ext_Tag [assign internal/plc tags]
'----------------------------------------------------------------------------------------------------     
	Private Function PLC_Ext_Tag(Byval index, Byval DB_name) 
	On Error Resume Next
		If DB_name = "DB301_Order_Status" Then	'Datablock 301
				Dim DB301_Order_Status
					DB301_Order_Status 		= Array("RESERVED SPACE" , _
											   		"PLC_01/DB_Recipe_FL.MO2_R1_Order_Status" , _
											   		"PLC_01/DB_Recipe_FL.MO2_R2_Order_Status" , _
											   		"PLC_01/DB_Recipe_FL.MO2_R3_Order_Status" , _
											   		"PLC_01/DB_Recipe_FL.MO2_R4_Order_Status" , _
											   		"PLC_01/DB_Recipe_FL.AVA_R1_Order_Status" , _
											   		"PLC_01/DB_Recipe_FL.AVA_R2_Order_Status" , _
											   		"PLC_01/DB_Recipe_FL.AVA_R3_Order_Status") 
					PLC_Ext_Tag = DB301_Order_Status(index)	
					
		ElseIf DB_name = "DB302_Status" Then	'Datablock 302	
				Dim DB302_Status
					DB302_Status 			= Array("RESERVED SPACE" , _
													"PLC_01/Status_Batch_FL.MO2_R1_Status" , _
													"PLC_01/Status_Batch_FL.MO2_R2_Status" , _
													"PLC_01/Status_Batch_FL.MO2_R3_Status" , _
													"PLC_01/Status_Batch_FL.MO2_R4_Status" , _
													"PLC_01/Status_Batch_FL.AVA_R1_Status" , _
													"PLC_01/Status_Batch_FL.AVA_R2_Status" , _
													"PLC_01/Status_Batch_FL.AVA_R3_Status")
					PLC_Ext_Tag = DB302_Status(index)	
					
		ElseIf DB_name = "DB302_Batch_No" Then 'Datablock 302
				Dim DB302_Batch_No							
					DB302_Batch_No 			= Array("RESERVED SPACE" , _
													"PLC_01/Status_Batch_FL.MO2_R1_Batch_No" ,  _
													"PLC_01/Status_Batch_FL.MO2_R2_Batch_No" ,  _
													"PLC_01/Status_Batch_FL.MO2_R3_Batch_No" ,  _
													"PLC_01/Status_Batch_FL.MO2_R4_Batch_No" ,  _
													"PLC_01/Status_Batch_FL.AVA_R1_Batch_No" ,  _
													"PLC_01/Status_Batch_FL.AVA_R2_Batch_No" ,  _
													"PLC_01/Status_Batch_FL.AVA_R3_Batch_No")
					PLC_Ext_Tag = DB302_Batch_No(index)
					
		ElseIf DB_name = "DB303_Filling_Status" Then	'Datablock 303	
				Dim DB303_Filling_Status 
					DB303_Filling_Status 	= Array("RESERVED SPACE" , _
													"PLC_01/Process_Data_FL.MO2_R1_Filling_Status" , _
													"PLC_01/Process_Data_FL.MO2_R2_Filling_Status" , _
													"PLC_01/Process_Data_FL.MO2_R3_Filling_Status" , _
													"PLC_01/Process_Data_FL.MO2_R4_Filling_Status" , _
													"PLC_01/Process_Data_FL.AVA_R1_Filling_Status" , _
													"PLC_01/Process_Data_FL.AVA_R2_Filling_Status" , _
													"PLC_01/Process_Data_FL.AVA_R3_Filling_Status")
					PLC_Ext_Tag = DB303_Filling_Status(index)
					
		ElseIf DB_name = "DB303_Filling_Pressure" Then	'Datablock 303	
				Dim DB303_Filling_Pressure
					DB303_Filling_Pressure 	= Array("RESERVED SPACE" , _
													"PLC_01/Process_Data_FL.MO2_R1_Fill_Pressure" ,  _
													"PLC_01/Process_Data_FL.MO2_R2_Fill_Pressure" ,  _
													"PLC_01/Process_Data_FL.MO2_R3_Fill_Pressure" ,  _
													"PLC_01/Process_Data_FL.MO2_R4_Fill_Pressure" ,  _
													"PLC_01/Process_Data_FL.AVA_R1_Fill_Pressure" ,  _
													"PLC_01/Process_Data_FL.AVA_R2_Fill_Pressure" ,  _
													"PLC_01/Process_Data_FL.AVA_R3_Fill_Pressure")
					PLC_Ext_Tag = DB303_Filling_Pressure(index)								

		ElseIf DB_name = "DB303_Vacc_Pressure" Then	'Datablock 303											
				Dim DB303_Vacc_Pressure	
					DB303_Vacc_Pressure 	= Array("RESERVED SPACE" , _
													"PLC_01/Process_Data_FL.MO2_R1_Vacc_Pressure" , _
													"PLC_01/Process_Data_FL.MO2_R2_Vacc_Pressure" , _
													"PLC_01/Process_Data_FL.MO2_R3_Vacc_Pressure" , _
													"PLC_01/Process_Data_FL.MO2_R4_Vacc_Pressure" , _
													"PLC_01/Process_Data_FL.AVA_R1_Vacc_Pressure" , _
													"PLC_01/Process_Data_FL.AVA_R2_Vacc_Pressure" , _
													"PLC_01/Process_Data_FL.AVA_R3_Vacc_Pressure")
					PLC_Ext_Tag = DB303_Vacc_Pressure(index)	
					
		ElseIf DB_name = "DB303_Temp_1" Then	'Datablock 303											
				Dim DB303_Temp_1	
					DB303_Temp_1 			= Array("RESERVED SPACE" , _
													"PLC_01/Process_Data_FL.MO2_R1_Temperature_1" , _
													"PLC_01/Process_Data_FL.MO2_R2_Temperature_1" , _
													"PLC_01/Process_Data_FL.MO2_R3_Temperature_1" , _
													"PLC_01/Process_Data_FL.MO2_R4_Temperature_1" , _
													"PLC_01/Process_Data_FL.AVA_R1_Temperature_1" , _
													"PLC_01/Process_Data_FL.AVA_R2_Temperature_1" , _
													"PLC_01/Process_Data_FL.AVA_R3_Temperature_1")
					PLC_Ext_Tag = DB303_Temp_1(index)
					
		ElseIf DB_name = "DB303_Temp_2" Then	'Datablock 303											
				Dim DB303_Temp_2	
					DB303_Temp_2 			= Array("RESERVED SPACE" , _
													"PLC_01/Process_Data_FL.MO2_R1_Temperature_2" ,  _
													"PLC_01/Process_Data_FL.MO2_R2_Temperature_2" ,  _
													"PLC_01/Process_Data_FL.MO2_R3_Temperature_2" ,  _
													"PLC_01/Process_Data_FL.MO2_R4_Temperature_2" ,  _
													"PLC_01/Process_Data_FL.AVA_R1_Temperature_2" ,  _
													"PLC_01/Process_Data_FL.AVA_R2_Temperature_2" ,  _
													"PLC_01/Process_Data_FL.AVA_R3_Temperature_2")
					PLC_Ext_Tag = DB303_Temp_2(index)
					
		ElseIf DB_name = "DB303_Filling_Results" Then	'Datablock 303											
				Dim DB303_Filling_Results	
					DB303_Filling_Results 	= Array("RESERVED SPACE" , _
													"PLC_01/Process_Data_FL.MO2_R1_Filling_Result" ,  _
													"PLC_01/Process_Data_FL.MO2_R2_Filling_Result" ,  _
													"PLC_01/Process_Data_FL.MO2_R3_Filling_Result" ,  _
													"PLC_01/Process_Data_FL.MO2_R4_Filling_Result" ,  _
													"PLC_01/Process_Data_FL.AVA_R1_Filling_Result" ,  _
													"PLC_01/Process_Data_FL.AVA_R2_Filling_Result" ,  _
													"PLC_01/Process_Data_FL.AVA_R3_Filling_Result")
					PLC_Ext_Tag = DB303_Filling_Results(index)
		
		Elseif DB_name = "S_Time" Then
				Dim S_Time	
					S_Time 		= Array("RESERVED SPACE" , _
										"S_Time_1" ,  _
										"S_Time_2" ,  _
										"S_Time_3" ,  _
										"S_Time_4" ,  _
										"S_Time_5" ,  _
										"S_Time_6" ,  _
										"S_Time_7")
					PLC_Ext_Tag = S_Time(index)
										
		Elseif DB_name = "S_Date" Then
				Dim S_Date 	
					S_Date      = Array("RESERVED SPACE" , _
										"S_Date_1" ,  _
										"S_Date_2" ,  _
										"S_Date_3" ,  _
										"S_Date_4" ,  _
										"S_Date_5" ,  _
										"S_Date_6" ,  _
										"S_Date_7")
					PLC_Ext_Tag = S_Date(index)
					
		Elseif DB_name = "E_Time" Then
				Dim E_Time	
					E_Time      = Array("RESERVED SPACE" , _
										"E_Time_1" ,  _
										"E_Time_2" ,  _
										"E_Time_3" ,  _
										"E_Time_4" ,  _
										"E_Time_5" ,  _
										"E_Time_6" ,  _
										"E_Time_7")
					PLC_Ext_Tag = E_Time(index)
				
		Elseif DB_name = "E_Date" Then	
				Dim E_Date
					E_Date      = Array("RESERVED SPACE" , _
										"E_Date_1" ,  _
										"E_Date_2" ,  _
										"E_Date_3" ,  _
										"E_Date_4" ,  _
										"E_Date_5" ,  _
										"E_Date_6" ,  _
										"E_Date_7")
					PLC_Ext_Tag = E_Date(index)
					
		Elseif DB_name = "DB301_Recipe_Code_P1" Then	'Datablock 301
						Dim DB301_Recipe_Code_P1
							DB301_Recipe_Code_P1 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_11_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_11_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_11_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_11_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_11_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_11_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_11_Recipe_Code") 
							PLC_Ext_Tag = DB301_Recipe_Code_P1(index)	
							
							
		ElseIf DB_name = "DB301_Recipe_Code_P2" Then	'Datablock 301
						Dim DB301_Recipe_Code_P2
							DB301_Recipe_Code_P2 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_12_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_12_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_12_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_12_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_12_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_12_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_12_Recipe_Code") 
							PLC_Ext_Tag = DB301_Recipe_Code_P2(index)
							
		ElseIf DB_name = "DB301_Recipe_Code_P3" Then	'Datablock 301
						Dim DB301_Recipe_Code_P3
							DB301_Recipe_Code_P3 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_13_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_13_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_13_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_13_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_13_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_13_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_13_Recipe_Code") 
							PLC_Ext_Tag = DB301_Recipe_Code_P3(index)
							
		ElseIf DB_name = "DB301_Recipe_Code_P4" Then	'Datablock 301
						Dim DB301_Recipe_Code_P4
							DB301_Recipe_Code_P4 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_14_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_14_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_14_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_14_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_14_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_14_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_14_Recipe_Code") 
							PLC_Ext_Tag = DB301_Recipe_Code_P4(index)
							
		ElseIf DB_name = "DB301_Recipe_Code_P5" Then	'Datablock 301
						Dim DB301_Recipe_Code_P5
							DB301_Recipe_Code_P5 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_15_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_15_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_15_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_15_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_15_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_15_Recipe_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_15_Recipe_Code") 
							PLC_Ext_Tag = DB301_Recipe_Code_P5(index)

		'Datablock for Cylinder Code
			
		ElseIf DB_name = "DB301_Cylinder_Code_P1" Then	'Datablock 301
						Dim DB301_Cylinder_Code_P1
							DB301_Cylinder_Code_P1 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_11_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_11_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_11_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_11_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_11_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_11_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_11_Cylinder_Code") 
							PLC_Ext_Tag = DB301_Cylinder_Code_P1(index)	
							
		ElseIf DB_name = "DB301_Cylinder_Code_P2" Then	'Datablock 301
						Dim DB301_Cylinder_Code_P2
							DB301_Cylinder_Code_P2 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_12_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_12_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_12_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_12_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_12_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_12_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_12_Cylinder_Code") 
							PLC_Ext_Tag = DB301_Cylinder_Code_P2(index)	
							
		ElseIf DB_name = "DB301_Cylinder_Code_P3" Then	'Datablock 301
						Dim DB301_Cylinder_Code_P3
							DB301_Cylinder_Code_P3 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_13_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_13_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_13_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_13_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_13_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_13_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_13_Cylinder_Code") 
							PLC_Ext_Tag = DB301_Cylinder_Code_P3(index)	

		ElseIf DB_name = "DB301_Cylinder_Code_P4" Then	'Datablock 301
						Dim DB301_Cylinder_Code_P4
							DB301_Cylinder_Code_P4 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_14_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_14_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_14_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_14_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_14_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_14_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_14_Cylinder_Code") 
							PLC_Ext_Tag = DB301_Cylinder_Code_P4(index)	
							
		ElseIf DB_name = "DB301_Cylinder_Code_P5" Then	'Datablock 301
						Dim DB301_Cylinder_Code_P5
							DB301_Cylinder_Code_P5 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_15_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_15_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_15_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_15_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_15_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_15_Cylinder_Code" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_15_Cylinder_Code") 
							PLC_Ext_Tag = DB301_Cylinder_Code_P5(index)	
							
		'Datablock for Quantity
		ElseIf DB_name = "DB301_Qty_P1" Then	'Datablock 301
						Dim DB301_Qty_P1
							DB301_Qty_P1 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_11_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_11_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_11_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_11_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_11_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_11_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_11_Quantity") 
							PLC_Ext_Tag = DB301_Qty_P1(index)
							
		ElseIf DB_name = "DB301_Qty_P2" Then	'Datablock 301
						Dim DB301_Qty_P2
							DB301_Qty_P2 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_12_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_12_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_12_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_12_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_12_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_12_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_12_Quantity") 
							PLC_Ext_Tag = DB301_Qty_P2(index)
							
		ElseIf DB_name = "DB301_Qty_P3" Then	'Datablock 301
						Dim DB301_Qty_P3
							DB301_Qty_P3 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_13_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_13_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_13_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_13_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_13_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_13_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_13_Quantity") 
							PLC_Ext_Tag = DB301_Qty_P3(index)
							
		ElseIf DB_name = "DB301_Qty_P4" Then	'Datablock 301
						Dim DB301_Qty_P4
							DB301_Qty_P4 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_14_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_14_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_14_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_14_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_14_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_14_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_14_Quantity") 
							PLC_Ext_Tag = DB301_Qty_P4(index)
							
		ElseIf DB_name = "DB301_Qty_P5" Then	'Datablock 301
						Dim DB301_Qty_P5
							DB301_Qty_P5 		= Array("RESERVED SPACE" , _
															"PLC_01/DB_Recipe_FL.MO2_R1_15_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R2_15_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R3_15_Quantity" , _
															"PLC_01/DB_Recipe_FL.MO2_R4_15_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R1_15_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R2_15_Quantity" , _
															"PLC_01/DB_Recipe_FL.AVA_R3_15_Quantity") 
							PLC_Ext_Tag = DB301_Qty_P5(index)

		End If
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function PLC_Ext_Tag is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : Test [developer test functions]
'---------------------------------------------------------------------------------------------------- 	
		Public Function Test()
		On Error Resume Next
			Test = prod_qty(0)
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function Test is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
		End If
		End Function

'Description : Test [developer test functions]
'---------------------------------------------------------------------------------------------------- 		
		Public Function set_db301_recipe()
		On Error Resume Next
		  Select Case CInt(qty)
			Case 0
				Call product_one
			Case 1
				Call product_two
			Case 2
				Call product_three
			Case 3
				Call product_four
			Case 4
				Call Product_five
			case else
				HMIRuntime.Trace(Now & " Filling Product have reach the limits ("& qty & ")" & vbCrlf)
				Call GF_LogToFile_("Execute", "Products Max 5 [Limits]" ,"Loose")
		  End Select
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function set_db301_recipe is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
		End If
		End Function

'Description : product_one [assign 1 recipe/prod into Loose filling racks]
'---------------------------------------------------------------------------------------------------- 
		Public Function product_one
		On Error Resume Next
		Dim Tags
			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_recipe(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_cyl(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P1") 
			HMIRuntime.Tags(Tags).Write prod_qty(0)
			
			Call GF_LogToFile_("Execute", "One Products" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(0) & "|" & prod_cyl(0) & "|" & prod_qty(0) & "]" ,"Loose")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function product_one is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
		End If	
		End Function

'Description : product_two [assign 2 recipe/prod into Loose filling racks]
'---------------------------------------------------------------------------------------------------- 
		Public Function product_two
		On Error Resume Next
		Dim Tags
			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_recipe(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_cyl(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P1") 
			HMIRuntime.Tags(Tags).Write prod_qty(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_recipe(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_cyl(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P2") 
			HMIRuntime.Tags(Tags).Write prod_qty(1)
			
			Call GF_LogToFile_("Execute", "Two Products" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(0) & "|" & prod_cyl(0) & "|" & prod_qty(0) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(1) & "|" & prod_cyl(1) & "|" & prod_qty(1) & "]" ,"Loose")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function product_two is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
		End If
		End Function

'Description : product_three [assign 3 recipe/prod into Loose filling racks]
'---------------------------------------------------------------------------------------------------- 
		Public Function product_three
		On Error Resume Next
		Dim Tags
			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_recipe(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_cyl(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P1") 
			HMIRuntime.Tags(Tags).Write prod_qty(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_recipe(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_cyl(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P2") 
			HMIRuntime.Tags(Tags).Write prod_qty(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P3") 
			HMIRuntime.Tags(Tags).Write prod_recipe(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P3") 
			HMIRuntime.Tags(Tags).Write prod_cyl(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P3") 
			HMIRuntime.Tags(Tags).Write prod_qty(2)
			
			Call GF_LogToFile_("Execute", "Three Products" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(0) & "|" & prod_cyl(0) & "|" & prod_qty(0) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(1) & "|" & prod_cyl(1) & "|" & prod_qty(1) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(2) & "|" & prod_cyl(2) & "|" & prod_qty(2) & "]" ,"Loose")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function product_three is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
		End If	
		End Function

'Description : product_four [assign 4 recipe/prod into Loose filling racks]
'---------------------------------------------------------------------------------------------------- 
		Public Function product_four
		On Error Resume Next
		Dim Tags
			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_recipe(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_cyl(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P1") 
			HMIRuntime.Tags(Tags).Write prod_qty(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_recipe(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_cyl(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P2") 
			HMIRuntime.Tags(Tags).Write prod_qty(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P3") 
			HMIRuntime.Tags(Tags).Write prod_recipe(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P3") 
			HMIRuntime.Tags(Tags).Write prod_cyl(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P3") 
			HMIRuntime.Tags(Tags).Write prod_qty(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P4") 
			HMIRuntime.Tags(Tags).Write prod_recipe(3)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P4") 
			HMIRuntime.Tags(Tags).Write prod_cyl(3)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P4") 
			HMIRuntime.Tags(Tags).Write prod_qty(3)
			
			Call GF_LogToFile_("Execute", "Four Products" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(0) & "|" & prod_cyl(0) & "|" & prod_qty(0) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(1) & "|" & prod_cyl(1) & "|" & prod_qty(1) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(2) & "|" & prod_cyl(2) & "|" & prod_qty(2) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(3) & "|" & prod_cyl(3) & "|" & prod_qty(3) & "]" ,"Loose")
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function product_four is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
		End If	
		End Function

'Description : product_five [assign 5 recipe/prod into Loose filling racks]
'---------------------------------------------------------------------------------------------------- 
		Public Function product_five
		On Error Resume Next
		Dim Tags
			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_recipe(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_cyl(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P1") 
			HMIRuntime.Tags(Tags).Write prod_qty(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_recipe(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P2") 
			HMIRuntime.Tags(Tags).Write prod_cyl(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P2") 
			HMIRuntime.Tags(Tags).Write prod_qty(1)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P3") 
			HMIRuntime.Tags(Tags).Write prod_recipe(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P3") 
			HMIRuntime.Tags(Tags).Write prod_cyl(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P3") 
			HMIRuntime.Tags(Tags).Write prod_qty(2)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P4") 
			HMIRuntime.Tags(Tags).Write prod_recipe(3)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P4") 
			HMIRuntime.Tags(Tags).Write prod_cyl(3)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P4") 
			HMIRuntime.Tags(Tags).Write prod_qty(3)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P5") 
			HMIRuntime.Tags(Tags).Write prod_recipe(4)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P5") 
			HMIRuntime.Tags(Tags).Write prod_cyl(4)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P5") 
			HMIRuntime.Tags(Tags).Write prod_qty(4)
			
			Call GF_LogToFile_("Execute", "Five Products" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(0) & "|" & prod_cyl(0) & "|" & prod_qty(0) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(1) & "|" & prod_cyl(1) & "|" & prod_qty(1) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(2) & "|" & prod_cyl(2) & "|" & prod_qty(2) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(3) & "|" & prod_cyl(3) & "|" & prod_qty(3) & "]" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(4) & "|" & prod_cyl(4) & "|" & prod_qty(4) & "]" ,"Loose")
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function product_five is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
		End If	
		End Function

	
'Description : product_filter [filter product name and quantity from arrays]
'---------------------------------------------------------------------------------------------------- 	
	Public Function product_filter()
	On Error Resume Next	
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
    If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function product_filter is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If	
	End Function

'Description : pd [convert timestamp]
'----------------------------------------------------------------------------------------------------	
	Function pd(n, totalDigits) 
	On Error Resume Next
		if totalDigits > len(n) then 
			pd = String(totalDigits-len(n),"0") & n 
		else 
			pd = n 
		end if 
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function pd is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If	
	End Function 

'Description : pd [convert timestamp]
'----------------------------------------------------------------------------------------------------	
	Public Function getRack_name()
	On Error Resume Next
		getRack_name = rack_name 
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getRack_name is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If	
	End Function

'Description : set_ActStatus_DB301 [sent work order status = 1 into PLC]
'----------------------------------------------------------------------------------------------------	
	Public Function set_ActStatus_DB301()
	On Error Resume Next
	Dim set_status : set_status = 1
	Dim Tags 
		Tags = PLC_Ext_Tag(rack_no,"DB301_Order_Status") 
		HMIRuntime.Tags(Tags).Write set_status
		
		Call GF_LogToFile_("Execute", "Sent Into EnerTech PLC ","Loose")
		Call GF_LogToFile_("Execute", "Set Status = 1 ","Loose")
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function set_ActStatus_DB301 is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If	
	End Function

'Description : set_DeactStatus_DB301 [sent work order status = 0 into PLC]
'----------------------------------------------------------------------------------------------------	
	Public Function set_DeactStatus_DB301()
	On Error Resume Next
	Dim set_status : set_status = 0
	Dim Tags 
	
		Tags = PLC_Ext_Tag(rack_no,"DB301_Order_Status") 
		HMIRuntime.Tags(Tags).Write set_status
		
		Call GF_LogToFile_("Execute", "Set Status = 0 ", "Loose")
		Call GF_LogToFile_("Execute", "EnerTech Deactivate CP 1 : " & rack_no , "Loose")
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function set_DeactStatus_DB301 is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If	
	End Function

'Description : set_DeactStatus_DB303 [sent plc reading result DB303 = 0 into PLC]
'----------------------------------------------------------------------------------------------------		
	Public Function set_DeactStatus_DB303()
	On Error Resume Next
	Dim set_status : set_status = 0
	Dim Tags 
	
		Tags = PLC_Ext_Tag(rack_no,"DB303_Filling_Status") 
		HMIRuntime.Tags(Tags).Write set_status
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function set_DeactStatus_DB303 is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : getStatus_DB302 [Read the order status from PLC]
'----------------------------------------------------------------------------------------------------	
	Public Function getStatus_DB302()
	On Error Resume Next
	Dim get_status 
	Dim Tags 
		
		Tags = PLC_Ext_Tag(rack_no,"DB302_Status") 
		get_status = HMIRuntime.Tags(Tags).Read 
		getStatus_DB302 = get_status
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getStatus_DB302 is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : setS_Time [Set filling start time into Internal Tags]
'----------------------------------------------------------------------------------------------------	
	Public Function setS_Time()
	On Error Resume Next
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"S_Time") 
		HMIRuntime.Tags(Tags).Write Right("0" & Hour(Time),2) & ":" & Right("0" & Minute(Time),2)  & ":" & Right("0" & Second(Time),2) 'Time
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function setS_Time is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : setS_Date [Set filling start date into Internal Tags]
'----------------------------------------------------------------------------------------------------	
	Public Function setS_Date()
	On Error Resume Next
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"S_Date") 
		HMIRuntime.Tags(Tags).Write YEAR(Date()) & "-" & Pd(Month(date()),2) & "-" & Pd(DAY(date()),2) 
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function setS_Date is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : setE_Time [Set filling end time into Internal Tags]
'----------------------------------------------------------------------------------------------------	
	Public Function setE_Time()
	On Error Resume Next
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"E_Time") 
		HMIRuntime.Tags(Tags).Write Right("0" & Hour(Time),2) & ":" & Right("0" & Minute(Time),2)  & ":" & Right("0" & Second(Time),2) 'Time
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function setE_Time is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : setE_Date [Set filling end date into Internal Tags]
'----------------------------------------------------------------------------------------------------	
	Public Function setE_Date()
	On Error Resume Next
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"E_Date") 
		HMIRuntime.Tags(Tags).Write YEAR(Date()) & "-" & Pd(Month(date()),2) & "-" & Pd(DAY(date()),2) 
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function setE_Date is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : getStatus_DB303 [Read filling result/last status]
'----------------------------------------------------------------------------------------------------	
	Public Function getStatus_DB303()
	On Error Resume Next
	Dim get_status 
	Dim Tags 
		
		Tags = PLC_Ext_Tag(rack_no,"DB303_Filling_Status") 
		get_status = HMIRuntime.Tags(Tags).Read 
		getStatus_DB303 = get_status
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getStatus_DB303 is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : ClearLooseInternalTag [Clear Internal Tags inside Wincc]
'----------------------------------------------------------------------------------------------------	
	Public Function ClearLooseInternalTag()
	On Error Resume Next
	Dim Tags 
		Tags = PLC_Ext_Tag(rack_no,"S_Time") 
		HMIRuntime.Tags(Tags).Write ""
		
		Tags = PLC_Ext_Tag(rack_no,"S_Date") 
		HMIRuntime.Tags(Tags).Write ""
		
		Tags = PLC_Ext_Tag(rack_no,"E_Time") 
		HMIRuntime.Tags(Tags).Write ""
		
		Tags = PLC_Ext_Tag(rack_no,"E_Date") 
		HMIRuntime.Tags(Tags).Write ""
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function ClearLooseInternalTag is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : setFilling_Results [keep filling result inside local arrays] * no longer working accordingly
'----------------------------------------------------------------------------------------------------	
	Public Function setFilling_Results()
	On Error Resume Next
	Dim Tags 
		
		Tags = PLC_Ext_Tag(rack_no,"DB303_Filling_Pressure") 
		arr_result(0) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"DB303_Vacc_Pressure") 
		arr_result(1) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"DB303_Temp_1") 
		arr_result(2) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"DB303_Temp_2") 
		arr_result(3) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"DB303_Filling_Results") 
		arr_result(4) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"DB302_Batch_No") 
		arr_result(5) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"S_Time")
		arr_result(6) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"E_Time")
		arr_result(7) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"S_Date")
		arr_result(8) = HMIRuntime.Tags(Tags).Read 
		
		Tags = PLC_Ext_Tag(rack_no,"E_Date")
		arr_result(9) = HMIRuntime.Tags(Tags).Read
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function setFilling_Results is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If	
	End Function

'Description : object property 
'----------------------------------------------------------------------------------------------------	
	Public Function getFilling_Pressure()
	On Error Resume Next
		getFilling_Pressure = arr_result(0)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getFilling_Pressure is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getVac_Pressure()
	On Error Resume Next
		getVac_Pressure = arr_result(1)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getVac_Pressure is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getTemp1_Pressure()
	On Error Resume Next
		getTemp1_Pressure = arr_result(2)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getTemp1_Pressure is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getTemp2_Pressure()
	On Error Resume Next
		getTemp2_Pressure = arr_result(3)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getTemp2_Pressure is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getFilling_Result()
	On Error Resume Next
		getFilling_Result = arr_result(4)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getFilling_Result is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getBatch_Result()
	On Error Resume Next
		getBatch_Result = arr_result(5)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getBatch_Result is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getS_Time()
	On Error Resume Next
		getS_Time = arr_result(6)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getS_Time is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getE_Time()
	On Error Resume Next
		getE_Time = arr_result(7)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getE_Time is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getS_Date()
	On Error Resume Next
		getS_Date = arr_result(8)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getS_Date is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
	
	Public Function getE_Date()
	On Error Resume Next
		getE_Date = arr_result(9)
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function getE_Date is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : set_DeactCheckPoint_1 [Deactivate check point 1] 
'----------------------------------------------------------------------------------------------------	
	Public Function set_DeactCheckPoint_1()
	On Error Resume Next
		Call Mysql_Non_Query("Update codabix_trigger Set db_301 = 0 Where rack_id = "& rack_no & "")
		Call GF_LogToFile_("Execute", "EnerTech Deactivate CP 1 : " & rack_no , "Loose")
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function set_DeactCheckPoint_1 is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : set_DeactCheckPoint_2 [Deactivate check point 2] 
'----------------------------------------------------------------------------------------------------	
	Public Function set_DeactCheckPoint_2()
	On Error Resume Next
		Call Mysql_Non_Query("Update codabix_trigger Set state = 0 , prod_detail = '' , user_id = 0 Where rack_id = "& rack_no &"")
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function set_DeactCheckPoint_2 is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : StoreToDB [save into mysql database] 
'----------------------------------------------------------------------------------------------------		
	Public Function StoreToDB()
	On Error Resume Next
	Dim user_entry_start_date : user_entry_start_date = getS_Date & " " & HMIRuntime.Tags(PLC_Ext_Tag(rack_no,"S_Time")).Read  'getS_Time
	Dim user_entry_end_date : user_entry_end_date =  getE_Date & " " & HMIRuntime.Tags(PLC_Ext_Tag(rack_no,"E_Time")).Read  'getE_Time
	Dim filling_end_time_fromOS : filling_end_time_fromOS = DisplayDate(Now)
	
		Call Mysql_Non_Query("Insert INTO sub_fill_loose (user_entry_start_date , user_entry_end_date , oms_batch ," & _
		" vacuum_pressure , filling_pressure1 , filling_temperature1 ,filling_temperature2 , result , user_id )" & _
		" SELECT '" & user_entry_start_date & "', '" & filling_end_time_fromOS & "' , oms_batch ," & _
		" " & getVac_Pressure &" , "& getFilling_Pressure &" , "& getTemp1_Pressure &" , "& getTemp2_Pressure &" , "& getFilling_Result &" , " & uid & "  FROM codabix_trigger WHERE rack_id="& rack_no &"" )
		
		Call Mysql_Non_Query("Update pallet_table Set analysis_required = 1, filling_finish = 1 ,analysis_mode = 3,	fill_mode = 2 where oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id="& rack_no &")")
		
		Call Mysql_Non_Query("Update codabix_trigger Set oms_batch = '' Where rack_id="& rack_no & "")
		
		Call Mysql_Non_Query("Update filling_rack Set occupied = 0 Where rack_id="& rack_no & "")
		
		Call Mysql_Non_Query("UPDATE analysis_filling_history SET status = 2 WHERE oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& rack_no & ") and " & _ 
							" rack_name = '"& rack_name &"' ")
							
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function StoreToDB is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : set_DeactFilling_Batch [Deactivate filling loose, after receive filling result from plc] 
'----------------------------------------------------------------------------------------------------	
	Public Function set_DeactFilling_Batch()
	On Error Resume Next
	Dim Final_Status
		Final_Status = getStatus_DB303()
		If Final_Status = 1 Then
			 	Call setFilling_Results()
			 	Call set_DeactCheckPoint_2()
			 	Call StoreToDB()
			 	Call set_DeactStatus_DB303()
			 	Call ClearLooseInternalTag
			 	Call ClearPLC_Loose_EnerTech()
		End If
		Call GF_LogToFile_("Execute", "Deactivate All Filling : " & rack_no ,"Loose")
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function set_DeactFilling_Batch is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : time_start [Read time start from internal tags] 
'----------------------------------------------------------------------------------------------------	
	Public Property Get time_start
	On Error Resume Next
		Tags = PLC_Ext_Tag(rack_no,"S_Time") 
		work_status= HMIRuntime.Tags(Tags).Read
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Property time_start is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Property
	
'Description : time_end [Read time end from internal tags] 
'----------------------------------------------------------------------------------------------------	
	Public Property Get time_end
	On Error Resume Next
		Tags = PLC_Ext_Tag(rack_no,"E_Time") 
		work_status= HMIRuntime.Tags(Tags).Read
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Property time_end is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Property

'Description : Reset_FillLoose_EnerTech [Reset filling Loose] 
'----------------------------------------------------------------------------------------------------		
	Public Function Reset_FillLoose_EnerTech
	On Error Resume Next
		Call Mysql_Non_Query("Update codabix_trigger Set db_301 = 0 , state = 0 , prod_detail = '' , user_id = 0 , oms_batch = '' , cylinder_id = 0 , IsPrefilled = 0  Where rack_id = "& rack_no & "")
		Call Mysql_Non_Query("Update filling_rack Set occupied = 0 , cylinder_type = '' , user_id = ''  Where rack_id="& rack_no & "")
		Call Mysql_Non_Query("UPDATE analysis_filling_history SET status = 3 WHERE oms_batch = (SELECT distinct oms_batch from codabix_trigger WHERE rack_id = "& rack_no & ") and " & _ 
							" rack_name = '"& rack_name &"' ")
		
		Call GF_LogToFile_("RESET", " Filling Cryostar : " & rack_no , "Loose")
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function Reset_FillLoose_EnerTech is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function

'Description : ClearPLC_Loose_EnerTech [Reset filling Loose Internal Tags] 
'----------------------------------------------------------------------------------------------------	
	Public Function ClearPLC_Loose_EnerTech()
	On Error Resume Next
		Dim Tags
			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P1") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P1") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P1") 
			HMIRuntime.Tags(Tags).Write 0

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P2") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P2") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P2") 
			HMIRuntime.Tags(Tags).Write 0

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P3") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P3") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P3") 
			HMIRuntime.Tags(Tags).Write 0

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P4") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P4") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P4") 
			HMIRuntime.Tags(Tags).Write 0

			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P5") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P5") 
			HMIRuntime.Tags(Tags).Write ""

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P5") 
			HMIRuntime.Tags(Tags).Write 0
	If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Filling.bmo - Function ClearPLC_Loose_EnerTech is not Workings [" & Err.Description & "]","Loose")
		    Err.Clear
	End If
	End Function
End Class