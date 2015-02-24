# ParallelApproval
M-Files parallel approval scripts

This repository contains two scripts that are used for parallel approval.

This approval is done by transitioning the object back to the previous 
state if there are any items in the Approvers property.
After an approver moves the object into the approved state, 
they are removed from the Approvers and added to the Approved by property.
