Class MasterList_Analysis
	Private m_Medical
	Private m_Industry
	
	Public Property Get medical
		Set medical = m_Medical
	End Property
	
	Public Property Set medical(ByVal value)
		Set m_Medical = value
	End Property
	
	Public Property Get industry
		Set industry = m_Industry
	End Property
	
	Public Property Set industry(ByVal value)
		Set m_Industry = value
	End Property
	
End Class