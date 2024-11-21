Imports DriveWorks
Imports DriveWorks.EventFlow
Imports DriveWorks.Specification
Imports Titan.Rules.Execution

<Task("Check Task List for Errors on Active Form", "embedded://CheckTaskListForFormErrors.Puzzle-16x16.png")>
Public Class CurrentFormErrors
    Inherits Task

    Private Const PROCESSCLASS = "Specification Macro Task"
    Private Const PROCESSTARGET = "Current Form Errors"
    Private Const PROCESSDESCRIPTION = "Specification task returning the task list of the current active form"

    Private ReadOnly mReturnValueOutput As NodeOutput = Me.Outputs.Register("Return Value",
                                                        "Table containing information from the form task list",
                                                        GetType(StandardArrayValue),
                                                        "Return Value")

    Protected Overrides Sub Execute(ByVal ctx As SpecificationContext)
        ctx.Report.BeginProcess(PROCESSCLASS, PROCESSTARGET, PROCESSDESCRIPTION)

        ' In case we want to add any warnings for the Success With Warnings status output
        Dim Warnings As String = String.Empty

        Dim taskList As New List(Of Object())
        Dim taskListHeaders As New List(Of Object)
        taskListHeaders.Add("Message")
        taskListHeaders.Add("ControlName")
        taskListHeaders.Add("FormName")
        taskListHeaders.Add("Target")
        taskListHeaders.Add("Type")
        taskList.Add(taskListHeaders.ToArray)
        For i As Integer = 0 To ctx.TaskList.Count - 1
            'If useForm And Not (ctx.TaskList.ElementAt(i).TargetDisplayName.ToString.Split(",").ElementAt(1).ToUpper = formName.ToUpper) Then
            '    rpt.WriteDebug("Error found on another form skipped", ctx.TaskList.ElementAt(i).TargetDisplayName.ToString)
            'Else
            Dim taskEntry As New List(Of Object)
            taskEntry.Add(ctx.TaskList.ElementAt(i).Message.ToString)
            taskEntry.Add(ctx.TaskList.ElementAt(i).TargetDisplayName.ToString.Split(",").ElementAt(0))
            taskEntry.Add(ctx.TaskList.ElementAt(i).TargetDisplayName.ToString.Split(",").ElementAt(1))
            taskEntry.Add(ctx.TaskList.ElementAt(i).Target.ToString)
            taskEntry.Add(ctx.TaskList.ElementAt(i).Type.ToString)
            taskList.Add(taskEntry.ToArray)
            'End If
        Next
        ' Convert the List of Lists to a 2D Array
        Dim rowCount As Integer = taskList.Count
        Dim taskListArray(5, rowCount)
        For i As Integer = 0 To rowCount - 1
            For col As Integer = 0 To 4
                taskListArray(i, col) = taskList(i).ElementAt(col)
            Next col
        Next i

        Dim dataArray As Object()() = taskList.Select(Function(taskEntries) taskEntries.ToArray()).ToArray()
        Dim stdArray As JaggedArrayValue
        stdArray = New JaggedArrayValue(dataArray, 2)
        'Output the IArray
        mReturnValueOutput.Fulfill(stdArray)

        ' If there were warnings, exit from the yellow dot, otherwise, exit through the green dot
        If String.IsNullOrEmpty(Warnings) Then
            Me.SetState(EventFlow.NodeExecutionState.Successful)
        Else
            Me.SetState(EventFlow.NodeExecutionState.SuccessfulWithWarnings, "Completed with Warnings", Warnings, Reporting.ReportingLevel.Verbose)
        End If

        ctx.Report.EndProcess()
    End Sub
End Class