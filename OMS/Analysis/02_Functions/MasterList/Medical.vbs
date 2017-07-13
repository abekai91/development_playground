Class Medical
	Private m_product, m_temperature , m_pressure
	Private objConn,objRS
	Private DB, UID, PASS, SVR , Pressure_Res , Pressure_cal
	
	Private Sub Class_Initialize()
		On Error Resume Next
		Set objConn = CreateObject("ADODB.Connection")
		Set objRS   = CreateObject("ADODB.Recordset")
		
		SVR  = HMIRuntime.Tags("Server").Read
		DB   = HMIRuntime.Tags("Database").Read
		UID  = HMIRuntime.Tags("UID").Read
		PASS = HMIRuntime.Tags("PASS").Read
	
		Call Mysql_Open_Conn(objConn,objRS,SVR,DB,UID,PASS)
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Medical.bmo - Function Class_Initialize is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Sub
	
	Public Property Get product
		product = m_product
	End Property

	Public Property Let product(ByVal value)
		m_product = value
	End Property
	
	Public Property Get temperature
		temperature = m_temperature
	End Property

	Public Property Let temperature(ByVal value)
		m_temperature = value
	End Property
	
	Public Property Get pressure
		pressure = m_pressure
	End Property

	Public Property Let pressure(ByVal value)
		m_pressure = value
	End Property
	
	Public Property Get pressure_calculation
		pressure_calculation = Pressure_cal
	End Property

	Public Property Let pressure_calculation(ByVal value)
		pressure_calculation = value
	End Property
	
	Public Property Get pressure_result
		On Error Resume Next
		Set objRS = objConn.execute("Select prod_name,val_1,val_2 " & _
									" from analysis_formula where prod_name = '" & m_product & "' and pressure_type = 'M'")
		
		If objRS.EOF  Then
			Pressure_Res = "PASS"
			Pressure_cal = 0
		Else
			Pressure_cal = (CDbl(objRS(1).value) * CDbl(m_temperature)) + CDbl(objRS(2).value)
			HMIRuntime.Trace("pressure calculation : (" & objRS(1).value & "x" & m_temperature & ") + " & objRS(2).value & " = " & Pressure_cal & vbCrlf  )
			
			If CDbl(m_pressure) >= Pressure_cal Then
				HMIRuntime.Trace("pressure calculation [PASS]: (pressure_db)" & m_pressure & ">" & "(pressure_cal)"& Pressure_cal & vbCrlf)
				Pressure_Res = "PASS"
			Else
				HMIRuntime.Trace("pressure calculation [FAIL]: (pressure_db)" & m_pressure & "<" & "(pressure_cal)"& Pressure_cal & vbCrlf)
				Pressure_Res = "FAIL"
			End If
		End If
		
		pressure_result = Pressure_Res
		
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Medical.bmo - Property pressure_result is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Property
	
	Public Property Get filling_clasification
		On Error Resume Next
		Set objRS = objConn.execute("Select filling_clasification " & _
									" from masterlist_medical where product_name = '" & m_product & "'")
		
		filling_clasification = objRS(0).value
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Medical.bmo - Property filling_clasification is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Property
	
	Public Property Get analyzer_name
		On Error Resume Next
		Set objRS = objConn.execute("Select analyzer_1,analyzer_2, " & _
									" analyzer_3,analyzer_4, " & _
									" analyzer_5,analyzer_6 " & _
									" from masterlist_medical where product_name = '" & m_product & "'")
		
		analyzer_name = objRS(0).value & "," & objRS(1).value & "," & objRS(2).value & "," & objRS(3).value & "," & objRS(4).value & "," & objRS(5).value
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Medical.bmo - Property analyzer_name is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Property
	
	Public Property Get analyzer_number
		analyzer_number = "1,2,3,4,5,6"
	End Property
	
	Public Property Get analyzer_val
		On Error Resume Next
		Set objRS = objConn.execute("Select analyzer_1_min,analyzer_1_max,analyzer_2_min,analyzer_2_max, " & _
									" analyzer_3_min,analyzer_3_max,analyzer_4_min,analyzer_4_max, " & _
									" analyzer_5_min,analyzer_5_max,analyzer_6_min,analyzer_6_max " & _
									" from masterlist_medical where product_name = '" & m_product & "'")
		
		analyzer_val = 	objRS(0).value & "|" & objRS(1).value & "," & objRS(2).value & "|" & objRS(3).value & "," &	objRS(4).value & "|" & objRS(5).value & "," & objRS(6).value & "|" & objRS(7).value & "," &	objRS(8).value & "|" & objRS(9).value & "," & objRS(10).value & "|" & objRS(11).value					
		If Err.Number <> 0 Then
		    Call GF_LogError("Error", "Medical.bmo - Property analyzer_val is not Workings [" & Err.Description & "]","Analysis")
		    Err.Clear
		End If
	End Property
	
	
End Class