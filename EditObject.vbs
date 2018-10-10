'Özet
'Sisteme girilen çalışma sürelerini proje bazlı toplayıp projenin üzerine toplam saat özelliğini doldurur / günceller.

'Gereksinimler
'Proje Nesnesi, Çalışma Süresi nesnesi

'Constants
Const ProjectObjectTypeId = 125
Const TimeSheetObjectTypeId = 126

Const WorkTimeClassId = 26

Const ProjectPropertyId = 1121
Const WorkTimePropId = 1125
Const WorkTimeTotalPropId = 1126

If ObjVer.Type = TimeSheetObjectTypeId Then
    Set Properties = Vault.ObjectPropertyOperations.GetProperties(ObjVer)

    'Check If Project & Worktime Properties are on Card
    If Properties.IndexOf(ProjectPropertyId) <> -1 and Properties.IndexOf(WorkTimePropId) Then

        Dim oProjectValue : oProjectValue = Properties.SearchForProperty(ProjectPropertyId).TypedValue.DisplayValue
        Dim oWorkTimeValue : oWorkTimeValue = Properties.SearchForProperty(WorkTimePropId).TypedValue.DisplayValue

        If oProjectValue <> "" and oWorkTimeValue <> "" Then

            Dim oProjectLkp : Set oProjectLkp = Properties.SearchForProperty(ProjectPropertyId).TypedValue.GetValueAsLookup()
            Dim oWorkTime : oWorkTime = Properties.SearchForProperty(WorkTimePropId).TypedValue.DisplayValue


            If Not oProjectLkp.Deleted Then
                'Define ObjID for Project
                Dim oObjID : Set oObjID  = CreateObject("MFilesAPI.ObjID")
                oObjID.ID = oProjectLkp.Item
                oObjID.Type = oProjectLkp.ObjectType

                'Get Project From Server
                Set oProjectVersionAndProperties = Vault.ObjectOperations.GetLatestObjectVersionAndProperties(oObjID, true)

                'Set project Properties
                Set projectProperties = oProjectVersionAndProperties.Properties

                'Get Current Total Time
				Dim currentTotalTime : currentTotalTime = 0
                Dim totalTime : totalTime = 0

                If projectProperties.IndexOf(WorkTimeTotalPropId) Then
                    totalTimeValue = projectProperties.SearchForProperty(WorkTimeTotalPropId).TypedValue.DisplayValue
                    If totalTimeValue <> "" Then
                        currentTotalTime = CInt(projectProperties.SearchForProperty(WorkTimeTotalPropId).TypedValue.DisplayValue)
                    End If
                End If

                'Build Conditions for finding timesheets related to this project

				Set oScP = CreateObject("MFilesAPI.SearchCondition")
				Set oScsP = CreateObject("MFilesAPI.SearchConditions")

                oScP.Expression.DataPropertyValuePropertyDef = 100
                oScP.ConditionType = MFConditionTypeEqual
                oScP.TypedValue.SetValue MFDatatypeLookup, WorkTimeClassId
                oScsP.Add -1, oScP

                oScP.Expression.DataPropertyValuePropertyDef = ProjectPropertyId
                oScP.ConditionType = MFConditionTypeEqual
                oScP.TypedValue.SetValue MFDataTypeLookup, oProjectLkp.Item
                oScsP.Add -1, oScP

                oScP.Expression.DataStatusValueType = MFStatusTypeDeleted
                oScP.ConditionType = MFConditionTypeEqual
                oScP.TypedValue.SetValue MFDatatypeBoolean, False
                oScsP.Add -1, oScP

                'Perform Search
                Set oSearchResultsP = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oScsP,MFSearchFlagNone,False)

                For Each xObject In oSearchResultsP
                    Set timeSheetProps  = Vault.ObjectPropertyOperations.GetProperties(xObject.ObjVer)
					If timeSheetProps.IndexOf(WorkTimePropId) <> -1 Then
						timeSheetValue = timeSheetProps.SearchForProperty(WorkTimePropId).TypedValue.DisplayValue
						If timeSheetValue <> "" Then
							timeSheetValue = CInt(timeSheetProps.SearchForProperty(WorkTimePropId).TypedValue.DisplayValue)

							totalTime = totalTime + timeSheetValue
						End If
					End If
                Next

				If currentTotalTime <> totalTime Then
					'Check-out project set new time
					If Vault.ObjectOperations.IsObjectCheckedOut(oObjID)Then
						Err.raise mfscriptcancel , "Süre girilen proje check-out edilmiş durumda. Lütfen projeyi check-in ediniz."
					End If

					Set oCheckedOut = Vault.ObjectOperations.CheckOut(oObjID)

					'Define totalTimeProperty
					Dim oTotalTimePropertyValue : Set oTotalTimePropertyValue = CreateObject("MFilesAPI.PropertyValue")
					oTotalTimePropertyValue.PropertyDef = WorkTimeTotalPropId
					oTotalTimePropertyValue.TypedValue.SetValue MFDataTypeInteger, totalTime

					'Set property on Project Object
					Set oCheckedOut = Vault.ObjectPropertyOperations.SetProperty(oCheckedOut.ObjVer, oTotalTimePropertyValue)

					'Check In the Object
					Vault.ObjectOperations.CheckIn oCheckedOut.ObjVer

				End If
            End If
        End If
    End If
End If



