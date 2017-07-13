Class MasterList_Sticker
	Private m_Report
	Private m_Cylinder
	
	Public Property Get Cyl_Sticker
		Set Cyl_Sticker = m_Cylinder
	End Property
	
	Public Property Set Cyl_Sticker(ByVal value)
		Set m_Cylinder = value
	End Property
	
	Public Property Get Report_Sticker
		Set Report_Sticker =  m_Report
	End Property
	
	Public Property Set Report_Sticker(ByVal value)
		Set m_Report = value
	End Property
	
End Class