'Özet
'Sistemde oluşturulan her firma için otomatik olarak projesi oluşturulur.

'Gereksinimler
'Proje Nesnesi, Firma Nesnesi

'Constants
Const CompanyObjectTypeId = 107
Const ProjectObjectTypeId = 125
Const RelatedCompanyPropId = 1122
Const ProjectClassId = 25

'Check the created object type
If ObjVer.Type = CompanyObjectTypeId Then

    'Get Current Object's Title
    Dim objInfo : Set objInfo = Vault.ObjectOperations.GetObjectInfo(ObjVer, true)

    Call CreateProjectForCompany(ObjVer.ID, objInfo.Title)
End If


Function CreateProjectForCompany(companyId, title)

    'Set PropertyValues
    Dim oPropertyValues : Set oPropertyValues = CreateObject("MFilesAPI.PropertyValues")
    Dim oPropertyValue : Set oPropertyValue = CreateObject("MFilesAPI.PropertyValue")
	Dim oAcl : Set oAcl = CreateObject("MFilesAPI.AccessControlList")
	Dim sFiles : Set sFiles = CreateObject("MFilesAPI.SourceObjectFiles")
    'Class
    oPropertyValue.PropertyDef = 100
    oPropertyValue.TypedValue.SetValue MFDataTypeLookup, ProjectClassId
    oPropertyValues.Add -1, oPropertyValue

    'Name
    oPropertyValue.PropertyDef = 0
    oPropertyValue.TypedValue.SetValue MFDataTypeText, title + " Projesi"
    oPropertyValues.Add -1, oPropertyValue

    'Related Company Field
    oPropertyValue.PropertyDef = RelatedCompanyPropId
    oPropertyValue.TypedValue.SetValue MFDataTypeLookup, companyId
    oPropertyValues.Add -1, oPropertyValue

    Set createdObject = Vault.ObjectOperations.CreateNewObject(ProjectObjectTypeId, oPropertyValues, sFiles, oAcl)

    Vault.ObjectOperations.CheckIn(createdObject.ObjVer)


End Function
