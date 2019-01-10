'20091119 - added try finally loops
'20150714 - modified code to write errors to a text file in preparation for server move to ITWListen.
'			Code mod required because of problems using event log on newer operating systems.

Public Class Form1
    Dim processit As Boolean = False
    Dim strLogString As String = ""



    'Public objIniFile As New INIFile("c:\newfeeds\HL7Mapper.ini")
    Public objIniFile As New INIFile("C:\FeedTester\Configs\ULH\HL7Mapper.ini")
    Public file1 As String = objIniFile.GetString("Scheduler", "file1", "(none)")
    Public file2 As String = objIniFile.GetString("Scheduler", "file2", "(none)")
    Public file3 As String = objIniFile.GetString("Scheduler", "file3", "(none)")
    Public file4 As String = objIniFile.GetString("Scheduler", "file4", "(none)")
    Public file5 As String = objIniFile.GetString("Scheduler", "file5", "(none)")
    Public file6 As String = objIniFile.GetString("Scheduler", "file6", "(none)")
    Public file7 As String = objIniFile.GetString("Scheduler", "file7", "(none)")
    Public file8 As String = objIniFile.GetString("Scheduler", "file8", "(none)")
    Public file9 As String = objIniFile.GetString("Scheduler", "file9", "(none)")
    Public file10 As String = objIniFile.GetString("Scheduler", "file10", "(none)")
    Public file11 As String = objIniFile.GetString("Scheduler", "file11", "(none)")
    Public file12 As String = objIniFile.GetString("Scheduler", "file12", "(none)")
    Public file13 As String = objIniFile.GetString("Scheduler", "file13", "(none)")
    Public file14 As String = objIniFile.GetString("Scheduler", "file14", "(none)")
    Public file15 As String = objIniFile.GetString("Scheduler", "file15", "(none)")
    Public file16 As String = objIniFile.GetString("Scheduler", "file16", "(none)")
    Public file17 As String = objIniFile.GetString("Scheduler", "file17", "(none)")
    Public file18 As String = objIniFile.GetString("Scheduler", "file18", "(none)")
    Public file19 As String = objIniFile.GetString("Scheduler", "file19", "(none)")
    Public file20 As String = objIniFile.GetString("Scheduler", "file20", "(none)")

    '20110905 increased size to 25 files
    Public file21 As String = objIniFile.GetString("Scheduler", "file21", "(none)")
    Public file22 As String = objIniFile.GetString("Scheduler", "file22", "(none)")
    Public file23 As String = objIniFile.GetString("Scheduler", "file23", "(none)")
    Public file24 As String = objIniFile.GetString("Scheduler", "file24", "(none)")
    Public file25 As String = objIniFile.GetString("Scheduler", "file25", "(none)")

    Public strLogDirectory As String = objIniFile.GetString("Settings", "logs", "(none)")



    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        processit = True
        runProcesses()

    End Sub

    Public Sub writeTolog(ByVal strMsg As String, ByVal eventType As Integer)
        '20150714 - use a text file to log errors instead of the event log
        Dim file As System.IO.StreamWriter
        Dim tempLogFileName As String = strLogDirectory & "Scheduler.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(tempLogFileName, True)
        file.WriteLine(DateTime.Now & " : " & strMsg & vbCrLf)
        file.Close()
    End Sub

    Public Sub writeToLog2(ByVal logText As String, ByVal eventType As Integer)
        Dim myLog As New EventLog()
        Try
            ' check for the existence of the log that the user wants to create.
            ' Create the source, if it does not already exist.
            If Not EventLog.SourceExists("Scheduler") Then
                EventLog.CreateEventSource("Scheduler", "SchedulerEvents")
            End If

            ' Create an EventLog instance and assign its source.

            myLog.Source = "Scheduler"

            ' Write an informational entry to the event log.
            If eventType = 1 Then
                myLog.WriteEntry(logText, EventLogEntryType.Error, 1)
            ElseIf eventType = 2 Then
                myLog.WriteEntry(logText, EventLogEntryType.Warning, 2)
            ElseIf eventType = 3 Then
                myLog.WriteEntry(logText, EventLogEntryType.Information, 3)
            End If


        Finally
            myLog.Close()
        End Try
    End Sub
    Public Sub runTheProcess(ByVal thefile As String)
        If thefile <> "" Then
            Dim splitterProcess As Process = Process.Start(thefile)
            Try

                System.Windows.Forms.Application.DoEvents()
                'Dim splitterProcess As Process = Process.Start("c:\newfeeds\programs\splitter.exe")

                If Len(txtStatus.Text) > 64000 Then txtStatus.Text = ""
                txtStatus.AppendText("Starting " & thefile & vbCrLf)
                System.Windows.Forms.Application.DoEvents()
                splitterProcess.WaitForExit()
                System.Windows.Forms.Application.DoEvents()
                txtStatus.AppendText("Stopping " & thefile & vbCrLf)

            Catch ex As Exception
                strLogString = strLogString & "Scheduler Error: " & thefile & vbCrLf
                strLogString = strLogString & ex.Message & vbCrLf
                writeToLog(strLogString, 1)
                strLogString = ""
                
            Finally
                If Not splitterProcess.HasExited Then
                    Me.Finalize()
                    splitterProcess.Kill()
                End If

                '20100128 make permanent 1 sec delay
                'If onedelay.Checked Then
                System.Windows.Forms.Application.DoEvents()
                txtStatus.AppendText("sleep for 1 sec.")
                System.Threading.Thread.Sleep(1000)
                System.Windows.Forms.Application.DoEvents()
                'txtStatus.AppendText("Awake.")
                'End If

                System.GC.Collect()
                'System.GC.WaitForPendingFinalizers()
                'System.GC.SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1)
            End Try


        End If
    End Sub

    Public Sub runProcesses()
        Try
            btnStart.Enabled = False
            Do While processit
                txtStatus.Text = ""
                runTheProcess(file1)
                runTheProcess(file2)
                runTheProcess(file3)
                runTheProcess(file4)
                runTheProcess(file5)
                runTheProcess(file6)
                runTheProcess(file7)
                runTheProcess(file8)
                runTheProcess(file9)
                runTheProcess(file10)
                runTheProcess(file11)
                runTheProcess(file12)
                runTheProcess(file13)
                runTheProcess(file14)
                runTheProcess(file15)
                runTheProcess(file16)
                runTheProcess(file17)
                runTheProcess(file18)
                runTheProcess(file19)
                runTheProcess(file20)
                '20110905 increased to 25 files
                runTheProcess(file21)
                runTheProcess(file22)
                runTheProcess(file23)
                runTheProcess(file24)
                runTheProcess(file25)
            Loop
        Finally
        End Try
    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        btnStart.Enabled = True

        processit = False
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '20100127 - start the scheduler when the program starts
        
    End Sub

    Private Sub btnReadINIFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReadINIFile.Click
        file1 = objIniFile.GetString("Scheduler", "file1", "(none)")
        file2 = objIniFile.GetString("Scheduler", "file2", "(none)")
        file3 = objIniFile.GetString("Scheduler", "file3", "(none)")
        file4 = objIniFile.GetString("Scheduler", "file4", "(none)")
        file5 = objIniFile.GetString("Scheduler", "file5", "(none)")
        file6 = objIniFile.GetString("Scheduler", "file6", "(none)")
        file7 = objIniFile.GetString("Scheduler", "file7", "(none)")
        file8 = objIniFile.GetString("Scheduler", "file8", "(none)")
        file9 = objIniFile.GetString("Scheduler", "file9", "(none)")
        file10 = objIniFile.GetString("Scheduler", "file10", "(none)")
        file11 = objIniFile.GetString("Scheduler", "file11", "(none)")
        file12 = objIniFile.GetString("Scheduler", "file12", "(none)")
        file13 = objIniFile.GetString("Scheduler", "file13", "(none)")
        file14 = objIniFile.GetString("Scheduler", "file14", "(none)")
        file15 = objIniFile.GetString("Scheduler", "file15", "(none)")
        file16 = objIniFile.GetString("Scheduler", "file16", "(none)")
        file17 = objIniFile.GetString("Scheduler", "file17", "(none)")
        file18 = objIniFile.GetString("Scheduler", "file18", "(none)")
        file19 = objIniFile.GetString("Scheduler", "file19", "(none)")
        file20 = objIniFile.GetString("Scheduler", "file20", "(none)")
        '20110905 increased to 25 files
        file21 = objIniFile.GetString("Scheduler", "file21", "(none)")
        file22 = objIniFile.GetString("Scheduler", "file22", "(none)")
        file23 = objIniFile.GetString("Scheduler", "file23", "(none)")
        file24 = objIniFile.GetString("Scheduler", "file24", "(none)")
        file25 = objIniFile.GetString("Scheduler", "file25", "(none)")
        txtStatus.Text = ""
    End Sub

    Private Sub Form1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel

    End Sub

    Private Sub Form1_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move

    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        '20100128 - added startup code here.
        btnStart.Enabled = False

        processit = True
        runProcesses()
    End Sub
End Class
