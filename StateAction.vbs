' This script is used for parallel approval
' This approval is done by transitioning the object
' back to the previous state if there are any items
' in the Approvers property. After an approver moves the
' object into the approved state, they are removed
' from the Approvers and added to the Approved by property

' This script should be placed in the Action Script of the Approved state

' Note: there is a companion script in the Automatic Transition of this workflow state
'  the companion script will transition the object to the Approved State or the Awaiting Approval State

' We do not want to use IDs in case this content is replicated
' Therefore, we will resolve aliases to IDs

' Resolve Property definition IDs
Dim approversPropertyID
approversPropertyID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias( "M-Files.QMS.SSDC.Approvers" )
Dim approvedByPropertyID
approvedByPropertyID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias( "M-Files.QMS.SSDC.ApprovedBy" )

Call RemoveLookup( approversPropertyID, CurrentUserID )
Call AddLookup( approvedByPropertyID, CurrentUserID )

Call SetLastModifiedBy(ObjVer, CurrentUserID.Value)

' Helper function for determining if the current object has a value in the given property
Function HasValue( id )
	HasValue = False
	If PropertyValues.IndexOf( id ) <> -1 Then
		If Not PropertyValues.SearchForProperty( id ).Value.IsNULL() Then
			HasValue = True
		End If
	End If
End Function

Function AddLookup( propId, lookupId )
	AddLookup = False

	Dim pv

	' Try to find existing property value.
	If PropertyValues.IndexOf( propId ) = -1 Then
		' Create the new object.
		Set pv = CreateObject("MFilesAPI.PropertyValue")
		pv.PropertyDef = propId
		Call pv.Value.SetValueToNULL( MFDataType.MFDatatypeMultiSelectLookup )
	Else
		Set pv = PropertyValues.SearchForProperty( propId )
	End If

	' Get the lookups.
	Dim lookups
	Set lookups = pv.Value.GetValueAsLookups()
	Dim match
	match = lookups.GetLookupIndexByItem( lookupId )
	If match = -1 Then
		' The item isn't in the list already. Add the item to the list and save.
		Dim lookup
		Set lookup = CreateObject( "MFilesAPI.Lookup" )
		lookup.Item = lookupId
		lookup.Version = -1 ' latest
		Call lookups.Add( -1, lookup )
		Call pv.Value.SetValueToMultiSelectLookup( lookups )
		AddLookup = True

		Call Vault.ObjectPropertyOperations.SetProperty( ObjVer, pv )
	End If
End Function

Function RemoveLookup( propId, lookupId )
	' Return null or unmodified property value.
	RemoveLookup = False
	
	' Only proceed if we found an existing property value.
	If PropertyValues.IndexOf( propId ) <> -1 Then
		' Property value exists.
		' Try to find passed item in lookups.
		Dim pv
		Set pv = PropertyValues.SearchForProperty( propId )
		Dim lookups
		Set lookups = pv.Value.GetValueAsLookups()
		Dim index
		index = lookups.GetLookupIndexByItem( lookupId )
		If index <> -1 Then
			' Lookup item exists. Remove it, and save the results.
			Call lookups.Remove( index )

			' TODO: this needs testing, do we need to set this back on PropertyValues?
			Call pv.Value.SetValueToMultiSelectLookup( lookups )
			RemoveLookup = True

			Call Vault.ObjectPropertyOperations.SetProperty( ObjVer, pv )
		End If
	End If
End Function

' Example usage: Dim updated : Set updated = SetLastModifiedBy(ObjVer, CurrentUserID.Value)
' Example usage: Call SetLastModifiedBy(ObjVer, CurrentUserID.Value)
Function SetLastModifiedBy(obj, lastModifiedByUser)
	' Set who modified object last
	' obj is an ObjVer
	' lastModifiedByUser is userId
	' returns ObjectVersionAndProperties
	Dim tvLastUser : Set tvLastUser = CreateObject("MFilesAPI.TypedValue")
	tvLastUser.SetValue 9, lastModifiedByUser

	Dim tvTime : Set tvTime = CreateObject("MFilesAPI.TypedValue")

	Set SetLastModifiedBy = Vault.ObjectPropertyOperations.SetLastModificationInfoAdmin(obj, True, tvLastUser, False, tvTime)
End Function