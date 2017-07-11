Class MasterList_Reports
	Private m_Individual
	Private m_Loose_Palletize
	
	Public Property Get Individual
		Set Individual = m_Individual
	End Property
	
	Public Property Set Individual(ByVal value)
		Set m_Individual = value
	End Property
	
	Public Property Get Loose_Palletize
		Set Loose_Palletize =  m_Loose_Palletize
	End Property
	
	Public Property Set Loose_Palletize(ByVal value)
		Set m_Loose_Palletize = value
	End Property
	
End Class