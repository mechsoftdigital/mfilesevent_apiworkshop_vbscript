
'Özet
'Çalışma Süresi nesnesi için geriye dönük tarih setlemesini engeller.

'Constants
Const timeSheetObjectType = 126

Dim selectedDate : selectedDate = PropertyValue.TypedValue.DisplayValue
Dim currentDate : currentDate = date

If selectedDate <> "" Then
	If ObjVer.Type = timeSheetObjectType Then
		If ObjVer.Version = 1 Then
		    IF DateValue(selectedDate) < DateValue(currentDate) Then
				err.raise mfscriptcancel ,"Geçmiş tarih için süre girilemez."
		    END IF
		End IF
	End If
End If
