Imports System.IO

'20091119 - added try finally loops
'20150714 - modified code to write errors to a text file in preparation for server move to ITWListen.
'			Code mod required because of problems using event log on newer operating systems.

Public Class Form1
    Dim processit As Boolean = False
    Dim strLogString As String = ""

    'Private fullinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\Configurations\HL7Mapper.ini")) ' New test
    Private fullinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\..\Configs\ULH\HL7Mapper.ini")) ' New test
    'Private fullinipath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "...\..\..\..\..\..\..\..\..\Configs\ULH\HL7Mapper.ini")) ' New test
    'Private fullinipath As String = "E:\Feed Tester\Feeds\ULH\Configurations\HL7Mapper.ini" ' New test
    'Private fullinipath As String = "C:\FeedTester\Configs\ULH\HL7Mapper.ini"

    Public objIniFile As New INIFile(fullinipath) '20140817 - New Test

    Dim direct As String = objIniFile.GetString("Settings", "directory", "(none)") & ":\"
    Dim par As String = objIniFile.GetString("Settings", "parentDir", "(none)") & "\"

    Public file1 As String = direct & par & objIniFile.GetString("Scheduler", "file1", "(none)")
    Public file2 As String = direct & par & objIniFile.GetString("Scheduler", "file2", "(none)")
    Public file3 As String = direct & par & objIniFile.GetString("Scheduler", "file3", "(none)")
    Public file4 As String = direct & par & objIniFile.GetString("Scheduler", "file4", "(none)")
    Public file5 As String = direct & par & objIniFile.GetString("Scheduler", "file5", "(none)")
    Public file6 As String = direct & par & objIniFile.GetString("Scheduler", "file6", "(none)")
    Public file7 As String = direct & par & objIniFile.GetString("Scheduler", "file7", "(none)")
    Public file8 As String = direct & par & objIniFile.GetString("Scheduler", "file8", "(none)")
    Public file9 As String = direct & par & objIniFile.GetString("Scheduler", "file9", "(none)")
    Public file10 As String = direct & par & objIniFile.GetString("Scheduler", "file10", "(none)")
    Public file11 As String = direct & par & objIniFile.GetString("Scheduler", "file11", "(none)")
    Public file12 As String = direct & par & objIniFile.GetString("Scheduler", "file12", "(none)")
    Public file13 As String = direct & par & objIniFile.GetString("Scheduler", "file13", "(none)")
    Public file14 As String = direct & par & objIniFile.GetString("Scheduler", "file14", "(none)")
    Public file15 As String = direct & par & objIniFile.GetString("Scheduler", "file15", "(none)")
    Public file16 As String = direct & par & objIniFile.GetString("Scheduler", "file16", "(none)")
    Public file17 As String = direct & par & objIniFile.GetString("Scheduler", "file17", "(none)")
    Public file18 As String = direct & par & objIniFile.GetString("Scheduler", "file18", "(none)")
    Public file19 As String = direct & par & objIniFile.GetString("Scheduler", "file19", "(none)")
    Public file20 As String = direct & par & objIniFile.GetString("Scheduler", "file20", "(none)")

    '20110905 increased size to 25 files
    Public file21 As String = direct & par & objIniFile.GetString("Scheduler", "file21", "(none)")
    Public file22 As String = direct & par & objIniFile.GetString("Scheduler", "file22", "(none)")
    Public file23 As String = direct & par & objIniFile.GetString("Scheduler", "file23", "(none)")
    Public file24 As String = direct & par & objIniFile.GetString("Scheduler", "file24", "(none)")
    Public file25 As String = direct & par & objIniFile.GetString("Scheduler", "file25", "(none)")

    Public strLogDirectory As String = direct & par & objIniFile.GetString("Settings", "logs", "(none)")



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
        If thefile <> "" And thefile <> direct & par Then
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
                writeTolog(strLogString, 1)
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
            Dim lastrun As DateTime
            btnStart.Enabled = False
            Do While processit
                txtStatus.Text = ""
                runTheProcess(file1)
                If DateTime.Now.Minute Mod 5 = 0 Then 'Run every five minutes
                    runTheProcess(file2) 'Process Timeouts
                End If
                runTheProcess(file3)
                runTheProcess(file4)
                runTheProcess(file5)
                runTheProcess(file6)
                runTheProcess(file7)
                runTheProcess(file8)
                runTheProcess(file9)
                runTheProcess(file10)
                'If DateTime.Now.Hour = 20 And DateTime.Now.Minute = 55 And lastrun.Day() <> Now.Day() Then 'at 9:01 PM
                runTheProcess(file11) 'DFT
                '    lastrun = Now
                'End If
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
        file1 = direct & par & objIniFile.GetString("Scheduler", "file1", "(none)")
        file2 = direct & par & objIniFile.GetString("Scheduler", "file2", "(none)")
        file3 = direct & par & objIniFile.GetString("Scheduler", "file3", "(none)")
        file4 = direct & par & objIniFile.GetString("Scheduler", "file4", "(none)")
        file5 = direct & par & objIniFile.GetString("Scheduler", "file5", "(none)")
        file6 = direct & par & objIniFile.GetString("Scheduler", "file6", "(none)")
        file7 = direct & par & objIniFile.GetString("Scheduler", "file7", "(none)")
        file8 = direct & par & objIniFile.GetString("Scheduler", "file8", "(none)")
        file9 = direct & par & objIniFile.GetString("Scheduler", "file9", "(none)")
        file10 = direct & par & objIniFile.GetString("Scheduler", "file10", "(none)")
        file11 = direct & par & objIniFile.GetString("Scheduler", "file11", "(none)")
        file12 = direct & par & objIniFile.GetString("Scheduler", "file12", "(none)")
        file13 = direct & par & objIniFile.GetString("Scheduler", "file13", "(none)")
        file14 = direct & par & objIniFile.GetString("Scheduler", "file14", "(none)")
        file15 = direct & par & objIniFile.GetString("Scheduler", "file15", "(none)")
        file16 = direct & par & objIniFile.GetString("Scheduler", "file16", "(none)")
        file17 = direct & par & objIniFile.GetString("Scheduler", "file17", "(none)")
        file18 = direct & par & objIniFile.GetString("Scheduler", "file18", "(none)")
        file19 = direct & par & objIniFile.GetString("Scheduler", "file19", "(none)")
        file20 = direct & par & objIniFile.GetString("Scheduler", "file20", "(none)")
        '20110905 increased to 25 files
        file21 = direct & par & objIniFile.GetString("Scheduler", "file21", "(none)")
        file22 = direct & par & objIniFile.GetString("Scheduler", "file22", "(none)")
        file23 = direct & par & objIniFile.GetString("Scheduler", "file23", "(none)")
        file24 = direct & par & objIniFile.GetString("Scheduler", "file24", "(none)")
        file25 = direct & par & objIniFile.GetString("Scheduler", "file25", "(none)")
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
