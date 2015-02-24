' This script is used for parallel approval
' This approval is done by transitioning the object
' back to the previous state if there are any items
' in the Approvers property. After an approver moves the
' object into the approved state, they are removed
' from the Approvers and added to the Approved by property

' This script should be placed in the Automatic State Transition of the Approved state

' Note: there is a companion script in the Action of this workflow state
'  the companion script will transfer the Approver to the ApprovedBy property

' We do not want to use IDs in case this content is replicated
' Therefore, we will resolve aliases to IDs

' Resolve Workflow state IDs
Dim approvedStateID
approvedStateID = VaultWorkflowOperations.GetWorkflowStateIDByAlias( "M-Files.QMS.SSDC.Workflow.NewMajorVer.State.ContentApproved" )
Dim awaitingStateID
awaitingStateID = VaultWorkflowOperations.GetWorkflowStateIDByAlias( "M-Files.QMS.SSDC.Workflow.NewMajorVer.State.PendApproval" )

' Resolve Property definition IDs
Dim approversPropertyID
approversPropertyID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias( "M-Files.QMS.SSDC.Approvers" )

' If Approvers is empty then we go to the approved state
If Not HasValue( approversPropertyID ) Then
	NextStateID = approvedStateID
Else
	' Otherwise we will go back to the awaiting state
	NextStateID = awaitingStateID
End If

AllowStateTransition = True

' Helper function for determining if the current object has a value in the given property
Function HasValue( id )
	HasValue = False
	If PropertyValues.IndexOf( id ) <> -1 Then
		If Not PropertyValues.SearchForProperty( id ).Value.IsNULL() Then
			HasValue = True
		End If
	End If
End Function