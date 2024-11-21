
' Import the Specification namespace so we have access to Specification flow
Imports DriveWorks
Imports DriveWorks.Specification

<Condition("Current Form Has Errors", "embedded://CheckTaskListForFormErrors.Puzzle-16x16.png")>
Public Class CurrentFormHasWarning
    Inherits Condition

    Private Const PROCESSCLASS = "Specification Macro Condition"
    Private Const PROCESSTARGET = "Current Form Has Errors"
    Private Const PROCESSDESCRIPTION = "Specification condition evaluating the task list of the current active form"

    Protected Overrides Function Evaluate(ctx As SpecificationContext) As Boolean
        ctx.Report.BeginProcess(PROCESSCLASS, PROCESSTARGET, PROCESSDESCRIPTION)

        If ctx.TaskList.Count > 0 Then
            ctx.Report.WriteEntry(Reporting.ReportingLevel.Normal,
                                  Reporting.ReportEntryType.Information,
                                  PROCESSCLASS,
                                  PROCESSTARGET,
                                  "Form task list evaluated",
                                  String.Format("<{0}> form errors exist on the current form: <{1}>", ctx.TaskList.Count, ctx.ActiveDialogOrForm.Form.Name), "")
            Return True
        Else
            ctx.Report.WriteEntry(Reporting.ReportingLevel.Normal,
                                  Reporting.ReportEntryType.Information,
                                  PROCESSCLASS,
                                  PROCESSTARGET,
                                  String.Format("No form errors exist on the current form: <{0}>", ctx.ActiveDialogOrForm.Form.Name), "")
            Return False
        End If
        ctx.Report.EndProcess()
    End Function
End Class
