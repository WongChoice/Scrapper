1::
2::
3::
	Column := A_ThisHotkey + 0	
	SendInput ^c
	ClipWait,1	
	if ErrorLevel	
		return
	xlApp := ComObjActive("Excel.Application")	; 
	xlCell_EmptyBelowData := xlApp.Columns(Column).Find("*",,,,,2).Offset(1,0)	
	if  !xlCell_EmptyBelowData.Address	
		xlCell_EmptyBelowData := xlApp.Cells(1,Column)	
	xlCell_EmptyBelowData.Value := Clipboard
return
