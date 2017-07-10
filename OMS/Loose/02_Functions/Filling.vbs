Class Filling
    Private rack_name
    Private rack_no
    Private prod_detail
    Private uid
	Private arr_result(10)
	
	Private prod_qty(5)
	Private prod_recipe(5)
	Private prod_cyl(5)
	
	Private qty

    Public Default Function Init(parameters)
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
    End Function

    Private Function InitOneParam(parameter1)
        If TypeName(parameter1) = "String" Then
            rack_name = parameter1
        Else
            rack_no = parameter1
        End If
        Set InitOneParam = Me
    End Function

    Private Function InitTwoParam(parameter1, parameter2)
        rack_name 		 = parameter1
        rack_no 		 = parameter2
        Set InitTwoParam = Me
    End Function
	
	Private Function InitThreeParam(parameter1, parameter2, parameter3)
    	rack_name 		 = parameter1
        rack_no 		 = parameter2
        prod_detail		 = parameter3
        
        Call product_filter
        
        Set InitThreeParam = Me
    End Function
    
    Private Function InitFourParam(parameter1, parameter2, parameter3, parameter4)
    	rack_name 		 = parameter1
        rack_no 		 = parameter2
        prod_detail		 = parameter3
        uid				 = parameter4
        
        Call product_filter
        
        Set InitFourParam = Me
    End Function
    
	Private Function PLC_Ext_Tag(Byval index, Byval DB_name) 

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
	End Function
	
		Public Function Test()
			Test = prod_qty(0)
		End Function
	
		Public Function set_db301_recipe()
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
		End Function

		Public Function product_one
		Dim Tags
			Tags = PLC_Ext_Tag(rack_no,"DB301_Recipe_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_recipe(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Cylinder_Code_P1") 
			HMIRuntime.Tags(Tags).Write prod_cyl(0)

			Tags = PLC_Ext_Tag(rack_no,"DB301_Qty_P1") 
			HMIRuntime.Tags(Tags).Write prod_qty(0)
			
			Call GF_LogToFile_("Execute", "One Products" ,"Loose")
			Call GF_LogToFile_("Execute", "[" & prod_recipe(0) & "|" & prod_cyl(0) & "|" & prod_qty(0) & "]" ,"Loose")
			
		End Function

		Public Function product_two
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
		End Function

		Public Function product_three
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
			
		End Function

		Public Function product_four
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
		End Function

		Public Function product_five
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
		End Function

	
	
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
	
	Function pd(n, totalDigits) 
		if totalDigits > len(n) then 
			pd = String(totalDigits-len(n),"0") & n 
		else 
			pd = n 
		end if 
	End Function 
			
	Public Function getRack_name()
		getRack_name = rack_name 
	End Function
		
	Public Function set_ActStatus_DB301()
	Dim set_status : set_status = 1
	Dim Tags 
		Tags = PLC_Ext_Tag(rack_no,"DB301_Order_Status") 
		HMIRuntime.Tags(Tags).Write set_status
		
		Call GF_LogToFile_("Execute", "Sent Into EnerTech PLC ","Loose")
		Call GF_LogToFile_("Execute", "Set Status = 1 ","Loose")
		
	End Function
	
	Public Function set_DeactStatus_DB301()
	Dim set_status : set_status = 0
	Dim Tags 
	
		Tags = PLC_Ext_Tag(rack_no,"DB301_Order_Status") 
		
		HMIRuntime.Tags(Tags).Write set_status
		
		Call GF_LogToFile_("Execute", "Set Status = 0 ", "Loose")
		Call GF_LogToFile_("Execute", "EnerTech Deactivate CP 1 : " & rack_no , "Loose")
	End Function
	
	Public Function set_DeactStatus_DB303()
	Dim set_status : set_status = 0
	Dim Tags 
	
		Tags = PLC_Ext_Tag(rack_no,"DB303_Filling_Status") 
		
		HMIRuntime.Tags(Tags).Write set_status
	End Function
	
	Public Function getStatus_DB302()
	Dim get_status 
	Dim Tags 
		
		Tags = PLC_Ext_Tag(rack_no,"DB302_Status") 
		get_status = HMIRuntime.Tags(Tags).Read 
		getStatus_DB302 = get_status
	End Function
	
	Public Function setS_Time()
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"S_Time") 
		HMIRuntime.Tags(Tags).Write Right("0" & Hour(Time),2) & ":" & Right("0" & Minute(Time),2)  & ":" & Right("0" & Second(Time),2) 'Time
	End Function
	
	Public Function setS_Date()
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"S_Date") 
		HMIRuntime.Tags(Tags).Write YEAR(Date()) & "-" & Pd(Month(date()),2) & "-" & Pd(DAY(date()),2) 
	End Function
	
	Public Function setE_Time()
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"E_Time") 
		HMIRuntime.Tags(Tags).Write Right("0" & Hour(Time),2) & ":" & Right("0" & Minute(Time),2)  & ":" & Right("0" & Second(Time),2) 'Time
	End Function
	
	Public Function setE_Date()
	Dim Tags
		Tags = PLC_Ext_Tag(rack_no,"E_Date") 
		HMIRuntime.Tags(Tags).Write YEAR(Date()) & "-" & Pd(Month(date()),2) & "-" & Pd(DAY(date()),2) 
	End Function
	
	Public Function getStatus_DB303()
	Dim get_status 
	Dim Tags 
		
		Tags = PLC_Ext_Tag(rack_no,"DB303_Filling_Status") 
		get_status = HMIRuntime.Tags(Tags).Read 
		getStatus_DB303 = get_status
	End Function
	
	Public Function ClearLooseInternalTag()
	Dim Tags 
		Tags = PLC_Ext_Tag(rack_no,"S_Time") 
		HMIRuntime.Tags(Tags).Write ""
		
		Tags = PLC_Ext_Tag(rack_no,"S_Date") 
		HMIRuntime.Tags(Tags).Write ""
		
		Tags = PLC_Ext_Tag(rack_no,"E_Time") 
		HMIRuntime.Tags(Tags).Write ""
		
		Tags = PLC_Ext_Tag(rack_no,"E_Date") 
		HMIRuntime.Tags(Tags).Write ""
	End Function
	
	Public Function setFilling_Results()
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
		
	End Function
	
	Public Function getFilling_Pressure()
		getFilling_Pressure = arr_result(0)
	End Function
	
	Public Function getVac_Pressure()
		getVac_Pressure = arr_result(1)
	End Function
	
	Public Function getTemp1_Pressure()
		getTemp1_Pressure = arr_result(2)
	End Function
	
	Public Function getTemp2_Pressure()
		getTemp2_Pressure = arr_result(3)
	End Function
	
	Public Function getFilling_Result()
		getFilling_Result = arr_result(4)
	End Function
	
	Public Function getBatch_Result()
		getBatch_Result = arr_result(5)
	End Function
	
	Public Function getS_Time()
		getS_Time = arr_result(6)
	End Function
	
	Public Function getE_Time()
		getE_Time = arr_result(7)
	End Function
	
	Public Function getS_Date()
		getS_Date = arr_result(8)
	End Function
	
	Public Function getE_Date()
		getE_Date = arr_result(9)
	End Function
	
	Public Function set_DeactCheckPoint_1()
		Call Mysql_Non_Query("Update codabix_trigger Set db_301 = 0 Where rack_id = "& rack_no & "")
		Call GF_LogToFile_("Execute", "EnerTech Deactivate CP 1 : " & rack_no , "Loose")
	End Function
	
	Public Function set_DeactCheckPoint_2()
		Call Mysql_Non_Query("Update codabix_trigger Set state = 0 , prod_detail = '' , user_id = 0 Where rack_id = "& rack_no &"")
	End Function
	
	Public Function StoreToDB()
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
	End Function
	
	Public Function set_DeactFilling_Batch()
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
	End Function

	'New Development
	
	Public Property Get time_start
		Tags = PLC_Ext_Tag(rack_no,"S_Time") 
		work_status= HMIRuntime.Tags(Tags).Read
	End Property
	
	
	Public Property Get time_end
		Tags = PLC_Ext_Tag(rack_no,"E_Time") 
		work_status= HMIRuntime.Tags(Tags).Read
	End Property
	
	Public Function Reset_FillLoose_EnerTech
		Call Mysql_Non_Query("Update codabix_trigger Set db_301 = 0 , state = 0 , prod_detail = '' , user_id = 0 , oms_batch = '' , cylinder_id = 0 , IsPrefilled = 0  Where rack_id = "& rack_no & "")
		Call Mysql_Non_Query("Update filling_rack Set occupied = 0 , cylinder_type = '' , user_id = ''  Where rack_id="& rack_no & "")
		
		Call GF_LogToFile_("RESET", " Filling Cryostar : " & rack_no , "Loose")
	End Function
	
	Public Function ClearPLC_Loose_EnerTech()
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
	End Function
End Class