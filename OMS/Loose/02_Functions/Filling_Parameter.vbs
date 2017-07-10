Function param_one(objConn,objRS,index)

	Set objRS = objConn.execute("Select prod_detail from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_one =  objRS(0).value
	
End Function

Function param_two(objConn,objRS,index)

	Set objRS = objConn.execute("Select user_id from codabix_trigger where" & _
     				  		   " trigger_name='DB2001_Trigger' and rack_id = "& index &"")
    
	param_two =  objRS(0).value
	
End Function