Imports System
Imports CommandLine

Module Program
    Sub Main()
        Dim clArgs() As String = Environment.GetCommandLineArgs()
        Dim sla As Integer = 0
        Dim location As String = String.Empty
        Dim datetime As String = String.Empty

        If clArgs.Count = 7 Then
            For i As Integer = 1 To 3 Step 2
                If clArgs(i) = "-s" Then
                    sla = clArgs(i + 1)
                ElseIf clArgs(i) = "-l" Then
                    location = clArgs(i + 1)
                ElseIf clArgs(i) = "-dt" Then
                    datetime = clArgs(i + 1)
                End If
            Next
        End If

        Console.WriteLine(sla)
        Console.WriteLine(location)
        Console.WriteLine(datetime)
        Console.ReadLine()

    End Sub
End Module
