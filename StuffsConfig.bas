Option Explicit
Public Function GetConfig() As config
Dim ConfigReturn As config
ReDim ConfigReturn.BranchNumberSearchHeaders(0)
ConfigReturn.BranchNumberSearchHeaders(0) = "Branch Number:*"
ReDim ConfigReturn.CommentSearchHeaders(0)
ConfigReturn.CommentSearchHeaders(0) = "Comments"
ReDim ConfigReturn.CustomerNameSearchHeaders(0)
ConfigReturn.CustomerNameSearchHeaders(0) = "Customer Name:*"
ReDim ConfigReturn.DateSearchHeaders(0)
ConfigReturn.DateSearchHeaders(0) = "DATE:*"
ReDim ConfigReturn.DCASearchHeaders(0)
ConfigReturn.DCASearchHeaders(0) = "Data Compliance Analyst Who Completed Audit Verification:*"
ReDim ConfigReturn.FRNumberSearchHeaders(0)
ConfigReturn.FRNumberSearchHeaders(0) = "WAVE FR No:*"
ReDim ConfigReturn.RequestedBySearchHeaders(0)
ConfigReturn.RequestedBySearchHeaders(0) = "Requested By:*"
ReDim ConfigReturn.ResultsSearchHeaders(0)
ConfigReturn.ResultsSearchHeaders(0) = "Pass or Fail"
GetConfig = ConfigReturn
End Function
