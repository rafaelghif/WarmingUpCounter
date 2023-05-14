Imports System.ComponentModel
Imports System.IO

Public Class WarmingUpCounterForm
    Private ReadOnly databasePath As String = GetIniValue("CONFIG", "databasePath", $"{CurrentDirectory}/INIFILES/config.ini")
    Private ReadOnly logPath As String = GetIniValue("CONFIG", "logPath", $"{CurrentDirectory}/INIFILES/config.ini")
    Private ReadOnly warmingUpMinimum As Integer = GetIniValue("CONFIG", "warmingUpMinimum", $"{CurrentDirectory}/INIFILES/config.ini", 3600)

    Private ReadOnly runningStatus As String() = {"Idle", "Idle", "Idle", "Idle", "Idle", "Idle"}
    Private ReadOnly countGroup As String() = {0, 0, 0, 0, 0, 0}

    Private Sub WarmingUpCounterForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Initialization()
    End Sub
    Private Sub WarmingUpCounterForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If MsgBox("Apakah kamu yakin ingin menutup aplikasi?", vbYesNo) = MsgBoxResult.Yes Then
            HandleClose()
            e.Cancel = False
        Else
            e.Cancel = True
        End If
    End Sub
    Private Sub Initialization()
        For Each control As Control In Controls
            control.BackColor = SystemColors.Control
            For Each comp As Control In control.Controls
                If comp.Name.Contains("Label") Then
                    comp.Text = 0
                ElseIf comp.Name.Contains("Linkage") Then
                    comp.Text = ""
                ElseIf comp.Name.Contains("Button") Then
                    comp.Text = "Start"
                    comp.Enabled = False
                ElseIf comp.Name.Contains("Reset") Then
                    comp.Enabled = False
                    comp.Visible = False
                ElseIf comp.Name.Contains("TextBox") Then
                    comp.Enabled = True
                End If
            Next
        Next
    End Sub
    Private Sub HandleClose()
        For Each control As Control In Controls
            Dim groupNumber As Integer = control.Tag
            Dim groupNumberSelector As Integer = groupNumber - 1

            If runningStatus(groupNumberSelector) <> "Idle" Then
                Dim serialNumber As String = ""
                Dim linkage As String = ""
                Dim warmingUpTime As Integer = countGroup(groupNumberSelector)
                Dim warmingStatus As String = "FAIL"
                For Each comp As Control In control.Controls
                    If comp.Name.Contains("TextBox") Then
                        serialNumber = comp.Text
                    ElseIf comp.Name.Contains("Linkage") Then
                        linkage = comp.Text
                    End If
                Next

                If warmingUpTime >= warmingUpMinimum Then
                    warmingStatus = "PASS"
                End If

                WriteLogFile(groupNumber, serialNumber, linkage, warmingUpTime, warmingStatus)
            End If
        Next
    End Sub
    Private Sub WriteLogFile(groupNumber As Integer, serialNumber As String, linkage As String, warmingUpTime As Integer, warmingUpStatus As String)
        Dim currentDatetime As String = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim filePath As String = $"{logPath}/{serialNumber}.txt"
        Using Writer As New StreamWriter(filePath, True)
            Writer.WriteLine($"{groupNumber},{serialNumber},{linkage},{warmingUpTime},{currentDatetime},{warmingUpStatus}")
        End Using
    End Sub
    Private Sub SearchDatabase(serialNumber As String, groupNumber As Integer)
        Dim datas As String() = File.ReadAllLines(databasePath)
        Dim isFound As Boolean = False
        For Each data As String In datas
            Dim dataArr As String() = data.Split(",")

            If dataArr(1) = serialNumber.ToUpper() Then
                isFound = True
                For Each control As Control In Controls
                    If Not control.Name = $"GroupBox{groupNumber}" Then
                        Continue For
                    End If

                    For Each comp As Control In control.Controls
                        If comp.Name.Contains("Linkage") Then
                            comp.Text = dataArr(30)
                        ElseIf comp.Name.Contains("Button") Then
                            comp.Enabled = True
                        End If
                    Next

                    Exit For
                Next
                Exit For
            End If
        Next
        If isFound = False Then
            MsgBox("Serial Number Not Found!", MsgBoxStyle.Critical)
        End If
    End Sub
    Private Sub HandleClickButton(groupNumber As Integer)
        Dim groupNumberSelector As Integer = groupNumber - 1
        Dim timer As Timer = GetTimerByGroupNumber(groupNumber)
        Dim control As Control = GetGroupBoxByGroupNumber(groupNumber)

        Select Case runningStatus(groupNumberSelector)
            Case "Idle"
                StartTimerAndChangeControlState(control, timer)
            Case "Running"
                StopTimerAndChangeControlState(control, timer)
            Case "Stop"
                ResetOrContinueTimerAndChangeControlState(control, timer)
        End Select
    End Sub
    Private Sub HandleReset(groupNumber As Integer)
        Dim control As Control = GetGroupBoxByGroupNumber(groupNumber)
        runningStatus(CInt(control.Tag) - 1) = "Idle"
        countGroup(CInt(control.Tag) - 1) = 0
        control.BackColor = SystemColors.Control

        For Each comp As Control In control.Controls
            If comp.Name.Contains("TextBox") Then
                comp.Text = ""
                comp.Enabled = True
            ElseIf comp.Name.Contains("Linkage") Then
                comp.Text = ""
            ElseIf comp.Name.Contains("Label") Then
                comp.Text = countGroup(CInt(control.Tag) - 1)
            ElseIf comp.Name.Contains("Button") Then
                comp.Text = "Start"
            End If
        Next
    End Sub
    Private Function GetGroupBoxByGroupNumber(groupNumber As Integer) As Control
        Return Controls.OfType(Of Control).FirstOrDefault(Function(c) c.Name = $"GroupBox{groupNumber}")
    End Function
    Private Function GetTimerByGroupNumber(groupNumber As Integer) As Timer
        Select Case groupNumber
            Case 1
                Return Timer1
            Case 2
                Return Timer2
            Case 3
                Return Timer3
            Case 4
                Return Timer4
            Case 5
                Return Timer5
            Case 6
                Return Timer6
            Case Else
                Return Nothing
        End Select
    End Function
    Private Sub StartTimerAndChangeControlState(control As Control, timer As Timer)
        control.BackColor = Color.DarkOrange
        runningStatus(CInt(control.Tag) - 1) = "Running"

        For Each comp As Control In control.Controls
            If comp.Name.Contains("Button") Then
                comp.Text = "Stop"
            ElseIf comp.Name.Contains("TextBox") Then
                comp.Enabled = False
            End If
        Next

        timer.Enabled = True
        timer.Interval = 1000
        timer.Start()
    End Sub
    Private Sub StopTimerAndChangeControlState(control As Control, timer As Timer)
        For Each comp As Control In control.Controls
            If countGroup(CInt(control.Tag) - 1) < warmingUpMinimum Then
                control.BackColor = Color.Red
                If comp.Name.Contains("Button") Then
                    comp.Text = "Continue"
                ElseIf comp.Name.Contains("Reset") Then
                    comp.Enabled = True
                    comp.Visible = True
                End If
            Else
                control.BackColor = Color.LightGreen
                If comp.Name.Contains("Button") Then
                    comp.Text = "Reset"
                ElseIf comp.Name.Contains("Reset") Then
                    comp.Enabled = False
                    comp.Visible = False
                End If
            End If
        Next

        runningStatus(CInt(control.Tag) - 1) = "Stop"
        timer.Enabled = False
    End Sub
    Private Sub ResetOrContinueTimerAndChangeControlState(control As Control, timer As Timer)
        Dim buttonText As String = "Reset"
        For Each comp As Control In control.Controls
            If comp.Name.Contains("Button") Then
                buttonText = comp.Text
                Exit For
            End If
        Next

        If buttonText = "Continue" Then

            control.BackColor = Color.DarkOrange

            For Each comp As Control In control.Controls
                If comp.Name.Contains("Button") Then
                    comp.Text = "Stop"
                ElseIf comp.Name.Contains("Reset") Then
                    comp.Enabled = False
                    comp.Visible = False
                End If
            Next

            runningStatus(CInt(control.Tag) - 1) = "Running"
            timer.Start()
        Else
            control.BackColor = SystemColors.Control

            Dim groupNumber As Integer = control.Tag
            Dim groupNumberSelector As Integer = groupNumber - 1

            Dim serialNumber As String = ""
            Dim linkage As String = ""
            Dim warmingUpTime As Integer = countGroup(groupNumberSelector)
            Dim warmingStatus As String = "FAIL"

            If warmingUpTime >= warmingUpMinimum Then
                warmingStatus = "PASS"
            End If

            runningStatus(CInt(control.Tag) - 1) = "Idle"
            countGroup(CInt(control.Tag) - 1) = 0

            For Each comp As Control In control.Controls
                If comp.Name.Contains("Button") Then
                    comp.Text = "Start"
                ElseIf comp.Name.Contains("TextBox") Then
                    serialNumber = comp.Text
                    comp.Text = ""
                    comp.Enabled = True
                ElseIf comp.Name.Contains("Linkage") Then
                    linkage = comp.Text
                    comp.Text = ""
                ElseIf comp.Name.Contains("Label") Then
                    comp.Text = countGroup(CInt(control.Tag) - 1)
                ElseIf comp.Name.Contains("Reset") Then
                    comp.Enabled = False
                    comp.Visible = False
                End If
            Next

            WriteLogFile(groupNumber, serialNumber, linkage, warmingUpTime, warmingStatus)
        End If
    End Sub
    Private Sub TextBox_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyUp, TextBox2.KeyUp, TextBox3.KeyUp, TextBox4.KeyUp, TextBox5.KeyUp, TextBox6.KeyUp
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        Dim groupNumber As Integer = textBox.Tag

        If e.KeyCode = Keys.Enter Then
            SearchDatabase(textBox.Text, groupNumber)
        End If
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click
        Dim button As Button = DirectCast(sender, Button)
        Dim groupNumber As Integer = button.Tag
        HandleClickButton(groupNumber)
    End Sub
    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick, Timer2.Tick, Timer3.Tick, Timer4.Tick, Timer5.Tick, Timer6.Tick
        Dim timer As Timer = DirectCast(sender, Timer)
        Dim groupNumber As Integer = timer.Tag
        countGroup(groupNumber - 1) += 1

        Dim label As Label = DirectCast(Controls.Find($"Label{groupNumber}", True)(0), Label)
        label.Text = countGroup(groupNumber - 1)

        If countGroup(groupNumber - 1) >= warmingUpMinimum Then
            Dim control As Control = GetGroupBoxByGroupNumber(groupNumber)
            control.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub Reset_Click(sender As Object, e As EventArgs) Handles Reset1.Click, Reset2.Click, Reset3.Click, Reset4.Click, Reset5.Click, Reset6.Click
        Dim button As Button = DirectCast(sender, Button)
        Dim groupNumber As Integer = button.Tag
        HandleReset(groupNumber)
    End Sub
End Class
