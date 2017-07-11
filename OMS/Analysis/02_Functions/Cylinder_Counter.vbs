Function cylinder_count()
	On Error Resume Next
	
	Dim count_1 : count_1 = HMIRuntime.Tags("PLC_01/DB_PLC_AYM_1.Cyl_Counter").Read
	Dim count_2 : count_2 = HMIRuntime.Tags("PLC_01/DB_PLC_AYM_2-3.Cyl_Counter").Read
	Dim count_3 : count_3 = HMIRuntime.Tags("PLC_01/DB_PLC_AYM_4-5.Cyl_Counter").Read
	Dim count_4 : count_4 = HMIRuntime.Tags("PLC_01/DB_PLC_AYI_1-2.Cyl_Counter").Read
	Dim count_5 : count_5 = HMIRuntime.Tags("PLC_01/DB_PLC_AYI_3-4-5.Cyl_Counter").Read
	Dim count_6 : count_6 = HMIRuntime.Tags("PLC_01/DB_PLC_AYI_6.Cyl_Counter").Read
	

	Call Mysql_Non_Query("Update analysis_rack Set cyl_counter="& count_1 &"  Where loc_id = 'AYM_1'")
	Call Mysql_Non_Query("Update analysis_rack Set cyl_counter="& count_2 &"  Where loc_id in ('AYM_2','AYM_3')")
	Call Mysql_Non_Query("Update analysis_rack Set cyl_counter="& count_3 &"  Where loc_id in ('AYM_4','AYM_5')")
	
	Call Mysql_Non_Query("Update analysis_rack Set cyl_counter="& count_4 &"  Where loc_id in ('AYI_1','AYI_2')")
	Call Mysql_Non_Query("Update analysis_rack Set cyl_counter="& count_5 &"  Where loc_id in ('AYI_3','AYI_4','AYI_5')")
	Call Mysql_Non_Query("Update analysis_rack Set cyl_counter="& count_6 &"  Where loc_id = 'AYI_6'")
	
	If Err.Number <> 0 Then
		  Call GF_LogError("Error", "Cylinder_Counter.bmo - Function cylinder_count is not Workings [" & Err.Description & "]","Analysis")
		  Err.Clear
	End If
End Function