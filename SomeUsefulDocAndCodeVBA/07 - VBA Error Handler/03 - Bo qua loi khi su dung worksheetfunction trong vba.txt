Private Sub tbName_Change()
	On error resume next
	me.lblSesc = application.WorksheetFuction.Vlookup(me.tbName, ThisWorkbook.sheets(.....)

	If Err.Number <> 0 then
		me.lblDesc = ""
	End if
End sub