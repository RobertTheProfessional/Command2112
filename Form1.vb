''''''''' ''''''''' ''''''''' ''''''''' ''''''''' 
''''''''' Copyright 2007-2009 Brian Saville
''''''''' Command2112 is free for non-commercial use.
''''''''' Reuse of the source code is free for non-commercial use
''''''''' provided that the source code used is credited to Command2112.
''''''''' ''''''''' ''''''''' ''''''''' '''''''''
Partial Public Class Form1

    Private booShowNotifyBalloonTip As Boolean
    Private booShowNotifyBalloonTip2 As Boolean
    Private Shared objSL As New Object

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not System.IO.File.Exists(AppDomain.CurrentDomain.GetData("DataDirectory").ToString & "\Command2112.sdf") Then
            If MsgBox("Command2112 allows you to control your Denon AVR-2112CI using the telnet protocol. Command2112 sends commands to your Denon and processes any responses. Command2112 is free for non-commercial use. The Command2112 source code is free for non-commercial use provided you cite Command2112 in the derived work. Photos of the Denon AVR-2112CI Photo in this program are Copyright D&M Holdings, Inc. This program is in no way associated with, produced by, or other otherwise endorsed by D&M Holdings, Inc." & vbCrLf & vbCrLf & "THIS PROGRAM AND ITS SOURCE CODE ARE PROVIDED **** AS IS **** AND THERE IS NO WARRANTY EITHER STATED OR IMPLIED. YOU USE THIS PROGRAM AT YOUR OWN RISK." & vbCrLf & vbCrLf & "This is your first time using this program. Do you wish to continue?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.No Then
                End
            End If
        End If

        booFormLoadStarted = True
        booLoading = True
        booShowNotifyBalloonTip = True
        booShowNotifyBalloonTip2 = True
        Button91.Location = Button86.Location

        Try

            SplashScreen1.SetStatus("Initializing")

            AddHandlers()

            LoadSettings()

            DrawTabs()

        Catch ex As Exception
            OutputError(ex)
        Finally
            booLoading = False
        End Try

    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        If Helper.CheckIP(True, False) Then

            Network_MT.Start()

            Try
                StartupQuery()
            Catch ex As Exception
                OutputError(ex)
            End Try

        End If

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        RemoveHandler Network_MT.ResponseReady, AddressOf HandleResponseReceived
        RemoveHandler Network_MT.MessageReady, AddressOf HandleMessageReceived

        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False

        Network_MT.Stop()

        My.Settings.Save()

    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized AndAlso CheckBox8.Checked Then
            Me.Hide()
            Me.NotifyIcon1.Visible = True

            If booShowNotifyBalloonTip Then
                Me.NotifyIcon1.ShowBalloonTip(3)
                booShowNotifyBalloonTip = False
            End If

        End If
    End Sub

#Region "Button Events"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        My.CustomSettings.SavedCommands.Instance.Value.Add(Me.TextBox1.Text)
        FillSavedCommands()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try

            If Helper.CheckIP(False, False) Then
                Commands.SendCommand(TextBox1.Text, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        My.CustomSettings.SavedCommands.Instance.Value.Clear()
        FillSavedCommands()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim strX As String
        Dim ip As Net.IPAddress

        strX = NumericUpDown2.Value.ToString & "."
        strX &= NumericUpDown3.Value.ToString & "."
        strX &= NumericUpDown4.Value.ToString & "."
        strX &= NumericUpDown5.Value.ToString

        If Net.IPAddress.TryParse(strX, ip) Then
            My.CustomSettings.IPAddress.Instance.Value = ip.ToString

            Network_MT.ResetConnection()
        Else
            MsgBox("Invalid IP address.")
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click, Button95.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.Up)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Me.TextBox4.Clear()
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click, Button96.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.Enter)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click, Button98.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.Down)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click, Button100.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.Right)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button44.Click, Button99.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.Left)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button45.Click, Button97.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.Return)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button46.Click, Button93.Click

        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.GUIOff)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Button47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button47.Click, Button94.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMenuCommand(MenuCommand.GUIOn)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button48.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) AndAlso ComboBox52.SelectedIndex >= 0 Then
                Commands.SendThis("SSTPN" & ComboBox52.SelectedValue.ToString & TextBox5.Text.PadRight(9) & ComboBox51.SelectedValue.ToString)
                TextBox5.Clear()
                ComboBox52.SelectedIndex = -1
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button49.Click

        If DataGridView2.SelectedRows.Count = 1 Then

            Dim tpi As TabPageInfo
            Dim intX As Integer

            intX = DataGridView2.SelectedRows(0).Index

            If intX > 0 Then
                tpi = lTabInfo(intX)
                tpi.TabOrder = intX
                lTabInfo(intX - 1).TabOrder = intX + 1
                lTabInfo.RemoveAt(intX)
                lTabInfo.Insert(intX - 1, tpi)
                DataGridView2.Rows(intX - 1).Selected = True
            End If

        End If

    End Sub

    Private Sub Button50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button50.Click

        If DataGridView2.SelectedRows.Count = 1 Then

            Dim tpi As TabPageInfo
            Dim intX As Integer

            intX = DataGridView2.SelectedRows(0).Index

            If intX < lTabInfo.Count - 1 Then
                tpi = lTabInfo(intX)
                tpi.TabOrder = intX + 2
                lTabInfo(intX + 1).TabOrder = intX + 1
                lTabInfo.RemoveAt(intX)
                lTabInfo.Insert(intX + 1, tpi)
                DataGridView2.Rows(intX + 1).Selected = True
            End If

        End If

    End Sub

    Private Sub Button51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button51.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                QueryNet()
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button52.Click
        If MsgBox("This will reset the internal Command2112 database. All of your personal settings and other stored information will be lost. This action does not interact with your Denon in any way. The database in question here only applies to Command2112, so you can't harm your receiver by carrying out this action." & vbCrLf & vbCr & "Do you wish to continue?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
            DB.Reset()
        End If
    End Sub

    Private Sub Button53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button53.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.Up)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button54.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.EnterPlayPause)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button55.Click
        Network_MT.ResetConnection()
    End Sub

    Private Sub Button56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button56.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.Down)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button57.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.Left)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button58.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.Right)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    ' Create new script
    Private Sub Button59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button59.Click

        Dim strX As String

        strX = Helper.LimitedSizeInputBoxInputBox("Please enter the name of the new script.", 100)

        If Not String.IsNullOrEmpty(strX) Then

            DB.Script.Create(strX)

            FillScripts(strX, Nothing, Nothing)

        End If

    End Sub

    Private Sub Button60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button60.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.Stop)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button61.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.PageUp)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button62.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearNet()
                Commands.SendNETCommand(NETCommand.PageDown)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button63.Click
        Try

            If Helper.CheckIP(False, False) Then

                Commands.SendNETCommand(NETCommand.EnterPlayPause)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button64.Click

        Try

            If Helper.CheckIP(False, False) Then

                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.EnterPlayPause)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Button65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button65.Click

        Try

            If Helper.CheckIP(False, False) Then

                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.EnterPlayPause)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Button66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button66.Click

        Try

            If Helper.CheckIP(False, False) Then

                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.EnterPlayPause)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Button67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button67.Click

        Try

            If Helper.CheckIP(False, False) Then

                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.EnterPlayPause)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Button68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button68.Click

        Try

            If Helper.CheckIP(False, False) Then

                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.Down)
                Commands.SendNETCommand(NETCommand.EnterPlayPause)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    ' Delete script
    Private Sub Button69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button69.Click

        Dim strX As String
        Dim drv As DataRowView

        If ListBox1.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox1.SelectedItem, DataRowView)
            strX = drv("name").ToString

            If MsgBox("Are you sure you want to delete " & strX & "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
                DB.Script.CascadeDelete(CInt(drv("ScriptID")))
                FillScripts(Nothing, Nothing, Nothing)
            End If
        End If

    End Sub

    ' Rename script
    Private Sub Button70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button70.Click

        Dim strX As String
        Dim drv As DataRowView

        If ListBox1.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox1.SelectedItem, DataRowView)

            strX = Helper.LimitedSizeInputBoxInputBox("Please enter the new name for this script.", 100, , drv("name").ToString)

            If Not String.IsNullOrEmpty(strX) Then
                DB.Script.UpdateName(CInt(drv("ScriptID")), strX)
                FillScripts(strX, Nothing, Nothing)
            End If

        End If

    End Sub

    ' Move selected script step up
    Private Sub Button71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button71.Click

        Dim drvS, drvD As DataRowView
        Dim strX As String

        If ListBox3.SelectedItems.Count = 1 AndAlso ListBox3.SelectedIndex > 0 Then
            drvS = DirectCast(ListBox3.SelectedItem, DataRowView)
            drvD = DirectCast(ListBox3.Items(ListBox3.SelectedIndex - 1), DataRowView)
            strX = drvS("description").ToString

            DB.ScriptEntry.UpdateSequence(CInt(drvS("ScriptEntryID")), CInt(drvD("ScriptEntryID")), CInt(drvS("sequence")), CInt(drvD("sequence")))

            FillScripts(ListBox1.SelectedValue.ToString, Nothing, strX)

        End If

    End Sub

    Private Sub Button72_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button72.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.PageDown)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button73_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button73.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.PageUp)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button74_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button74.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                QueryIPod()
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button75.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.Up)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button76.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.EnterPlayPause)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button77_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button77.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.Down)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button78.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.Left)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button79.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.Right)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub Button80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button80.Click
        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                ClearIPod()
                Commands.SendIPodCommand(IPodCommand.Stop)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    ' Delete script schedule
    Private Sub Button81_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button81.Click

        Dim strX As String
        Dim drv As DataRowView

        If ListBox2.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox2.SelectedItem, DataRowView)
            strX = drv("name").ToString

            If MsgBox("Are you sure you want to delete " & strX & "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
                DB.ScriptSchedule.Delete(CInt(drv("ScriptScheduleID")))

                FillScripts(ListBox1.SelectedValue.ToString, Nothing, Nothing)

            End If

        End If

    End Sub

    ' Create new script schedule
    Private Sub Button82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button82.Click

        Dim dr As DialogResult
        Dim drv As DataRowView

        If ListBox1.SelectedItems.Count = 1 Then

            drv = DirectCast(ListBox1.SelectedItem, DataRowView)

            ScriptScheduler.ScheduleName = Nothing
            ScriptScheduler.ScheduleCode = Nothing
            ScriptScheduler.ScriptID = CInt(drv("ScriptID"))
            dr = ScriptScheduler.ShowDialog(Me)

            If dr = Windows.Forms.DialogResult.OK Then

                Dim strX As String


                strX = drv("name").ToString


                DB.ScriptSchedule.Create(CInt(drv("ScriptID")), ScriptScheduler.ScheduleCode, ScriptScheduler.ScheduleName)
                FillScripts(ListBox1.SelectedValue.ToString, ScriptScheduler.ScheduleName, Nothing)

            End If

        End If

    End Sub

    ' Move selected script step down
    Private Sub Button83_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button83.Click

        Dim drvS, drvD As DataRowView
        Dim strX As String

        If ListBox3.SelectedItems.Count = 1 AndAlso ListBox3.SelectedIndex < ListBox3.Items.Count - 1 Then
            drvS = DirectCast(ListBox3.SelectedItem, DataRowView)
            drvD = DirectCast(ListBox3.Items(ListBox3.SelectedIndex + 1), DataRowView)
            strX = drvS("description").ToString

            DB.ScriptEntry.UpdateSequence(CInt(drvS("ScriptEntryID")), CInt(drvD("ScriptEntryID")), CInt(drvS("sequence")), CInt(drvD("sequence")))

            FillScripts(ListBox1.SelectedValue.ToString, Nothing, strX)

        End If

    End Sub

    ' Rename selected Script Step
    Private Sub Button84_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button84.Click

        Dim strX As String
        Dim drv As DataRowView

        If ListBox3.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox3.SelectedItem, DataRowView)

            strX = Helper.LimitedSizeInputBoxInputBox("Please enter the new name for this step.", 100, , drv("description").ToString)

            If Not String.IsNullOrEmpty(strX) Then
                DB.ScriptEntry.UpdateDescription(CInt(drv("ScriptEntryID")), strX)
                FillScripts(ListBox1.SelectedValue.ToString, Nothing, strX)
            End If

        End If


    End Sub

    ' Deleted selected script step
    Private Sub Button85_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button85.Click

        Dim strX As String
        Dim drv As DataRowView

        If ListBox3.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox3.SelectedItem, DataRowView)
            strX = drv("description").ToString

            If MsgBox("Are you sure you want to delete " & strX & "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
                DB.ScriptEntry.Delete(CInt(drv("ScriptEntryID")))
                FillScripts(ListBox1.SelectedValue.ToString, Nothing, Nothing)
            End If
        End If

    End Sub

    ' Record script steps
    Private Sub Button86_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button86.Click

        Dim drv As DataRowView

        If ListBox1.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox1.SelectedItem, DataRowView)
            RecordingScriptID = CInt(drv("ScriptID"))
            RecordingScriptName = drv("Name").ToString
            Button86.Visible = False
            Button91.Visible = True
            ListBox1.Enabled = False
            ListBox2.Enabled = False
            ListBox3.Enabled = False
            ToolStripStatusLabel3.Visible = True
        End If

    End Sub

    ' Export script
    Private Sub Button87_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button87.Click

        Dim ds As DataSet
        Dim sr As New IO.StringWriter
        Dim drv As DataRowView
        Dim dt As New DataTable

        If ListBox1.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox1.SelectedItem, DataRowView)

            ds = dsScripts.Clone
            dt.TableName = "Command2112 Version Info"
            dt.Columns.Add("SoftwareVersion", GetType(String))
            dt.Columns.Add("DBVersion", GetType(String))

            dt.Rows.Add(dt.NewRow)
            dt.Rows(0)(0) = My.Application.Info.Version.ToString(4)
            dt.Rows(0)(1) = DB.DBVersion

            ds.Tables(0).ImportRow(drv.Row)

            For Each dr As DataRow In drv.CreateChildView("Script_ScriptSchedule").Table.Rows
                ds.Tables(1).ImportRow(dr)
            Next

            For Each dr As DataRow In drv.CreateChildView("Script_ScriptEntry").Table.Rows
                ds.Tables(2).ImportRow(dr)
            Next

            ds.Tables.Add(dt)

            ds.WriteXml(sr, XmlWriteMode.WriteSchema)

            SaveFileDialog1.FileName = ""
            SaveFileDialog1.DefaultExt = ".c3s"
            SaveFileDialog1.Filter = "Command2112 Script Files|*.c3s|All Files|*.*"

            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                IO.File.WriteAllText(SaveFileDialog1.FileName, sr.ToString)
                MsgBox("Export was successful.", MsgBoxStyle.Information Or MsgBoxStyle.OkOnly)
            End If

        End If

    End Sub

    ' Import script
    Private Sub Button88_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button88.Click

        Dim ds As New DataSet
        Dim sr As IO.StreamReader

        OpenFileDialog1.FileName = ""
        OpenFileDialog1.DefaultExt = ".c3s"
        OpenFileDialog1.Filter = "Command2112 Script Files|*.c3s|All Files|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            sr = IO.File.OpenText(OpenFileDialog1.FileName)
            Try
                ds.ReadXml(sr, XmlReadMode.ReadSchema)
                DB.Script.CreateFromDR(ds.Tables(0).Rows(0))
                FillScripts(ds.Tables(0).Rows(0)("name").ToString, Nothing, Nothing)
                MsgBox("Import was successful.", MsgBoxStyle.Information Or MsgBoxStyle.OkOnly)
            Catch ex As Exception
                OutputError(ex)
            Finally
                sr.Close()
            End Try

        End If

    End Sub

    ' Edit script schedule
    Private Sub Button89_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button89.Click

        Dim dr As DialogResult
        Dim strX As String
        Dim drv As DataRowView

        If ListBox2.SelectedItems.Count = 1 Then
            drv = DirectCast(ListBox2.SelectedItem, DataRowView)
            strX = drv("code").ToString

            ScriptScheduler.ScheduleName = ListBox2.SelectedValue.ToString
            ScriptScheduler.ScheduleCode = strX
            ScriptScheduler.ScriptID = CInt(drv("ScriptID"))
            ScriptScheduler.ScriptScheduleID = CInt(drv("ScriptScheduleID"))

            If ScriptScheduler.ScheduleName = ScriptScheduler.ScheduleCode Then
                ScriptScheduler.ScheduleName = Nothing
            End If

            dr = ScriptScheduler.ShowDialog(Me)

            If dr = Windows.Forms.DialogResult.OK Then
                DB.ScriptSchedule.Update(CInt(drv("ScriptScheduleID")), ScriptScheduler.ScheduleCode, ScriptScheduler.ScheduleName)
                FillScripts(ListBox1.SelectedValue.ToString, ScriptScheduler.ScheduleName, Nothing)
            End If

        End If

    End Sub

    ' Play selected script
    Private Sub Button90_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button90.Click

        Dim drv As DataRowView

        If ListBox1.SelectedItems.Count = 1 AndAlso ListBox3.Items.Count > 0 AndAlso Helper.CheckIP(False, False) Then
            drv = DirectCast(ListBox1.SelectedItem, DataRowView)

            WriteDebug(Helper.GetNowString & " - BEGIN SCRIPT: " & ListBox1.SelectedValue.ToString, True)

            For intX As Integer = 0 To ListBox3.Items.Count - 1
                drv = DirectCast(ListBox3.Items(intX), DataRowView)

                Try
                    Commands.SendCommand(drv("command").ToString, My.CustomSettings.DefaultCommandPauseMS.Instance.Value)
                Catch ex As Exception
                    OutputError(ex)
                End Try

            Next

        End If
    End Sub

    ' Stop recording script steps
    Private Sub Button91_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button91.Click

        RecordingScriptID = 0
        RecordingScriptName = Nothing
        Button91.Visible = False
        Button86.Visible = True
        ListBox1.Enabled = True
        ListBox2.Enabled = True
        ListBox3.Enabled = True
        ToolStripStatusLabel3.Visible = False
    End Sub

#End Region

#Region "Checkbox Events"

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        My.CustomSettings.PreviewCommands.Instance.Value = DirectCast(sender, CheckBox).Checked
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        My.CustomSettings.AutoClearDebug.Instance.Value = DirectCast(sender, CheckBox).Checked
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Try

            Dim cb As CheckBox = DirectCast(sender, CheckBox)

            If Not booLoading Then
                Network_MT.LogRawBytes = cb.Checked
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        My.CustomSettings.AutoRefreshNet.Instance.Value = DirectCast(sender, CheckBox).Checked
        SetNetRefeshTimer()
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged
        My.CustomSettings.MinimizeToTray.Instance.Value = DirectCast(sender, CheckBox).Checked
    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        My.CustomSettings.AutoRefreshIPod.Instance.Value = DirectCast(sender, CheckBox).Checked
        SetIPodRefeshTimer()
    End Sub

#End Region

#Region "ComboBox Events"

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex <> 0 AndAlso Me.ComboBox1.SelectedItem IsNot Nothing Then
            Me.TextBox1.Text = Me.ComboBox1.SelectedItem.ToString
            ComboBox1.SelectedIndex = 0
        End If
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox12.SelectedIndexChanged

        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendVideoSelect(DirectCast(ComboBox12.SelectedItem, Source))
                SetControlWaiting(ComboBox12)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub ComboBox52_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox52.SelectedIndexChanged

        If Not booLoading AndAlso ComboBox52.SelectedIndex >= 0 Then
            TextBox5.Text = lTunerPresets(ComboBox52.SelectedIndex).Name
        End If

    End Sub

    Private Sub ComboBox53_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox53.SelectedIndexChanged

        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendThis("TFAN" & ComboBox53.SelectedValue.ToString)
                SetControlWaiting(ComboBox53)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

#End Region

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick

        Try

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendThis("TPAN" & lTunerPresets(e.RowIndex).Code)
                SetControlWaiting(DataGridView1)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

#Region "Generic Handlers"

    Private Sub HandleChannelVolumeChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, ComboBox7.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox9.SelectedIndexChanged, ComboBox10.SelectedIndexChanged, ComboBox11.SelectedIndexChanged, ComboBox29.SelectedIndexChanged, ComboBox30.SelectedIndexChanged, ComboBox31.SelectedIndexChanged, ComboBox32.SelectedIndexChanged

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then

                Commands.SendChannelVolume(DirectCast(cb.Tag, Channel), DirectCast(cb.SelectedItem, ChannelVolume))
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    ' Events are assigned in iFillRangeControls method
    Private Sub HandleDenonControlValueChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendRange(DirectCast(cb.Tag, RangeControlType), DirectCast(cb.SelectedItem, DenonRangeControl))
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub HandleOnOffRadioButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton4.CheckedChanged, RadioButton3.CheckedChanged, RadioButton5.CheckedChanged, RadioButton6.CheckedChanged, RadioButton7.CheckedChanged, RadioButton8.CheckedChanged, RadioButton9.CheckedChanged, RadioButton10.CheckedChanged, RadioButton11.CheckedChanged, RadioButton12.CheckedChanged, RadioButton13.CheckedChanged, RadioButton14.CheckedChanged, RadioButton15.CheckedChanged, RadioButton16.CheckedChanged, RadioButton17.CheckedChanged, RadioButton18.CheckedChanged, RadioButton19.CheckedChanged, RadioButton20.CheckedChanged, RadioButton21.CheckedChanged, RadioButton22.CheckedChanged, RadioButton23.CheckedChanged, RadioButton24.CheckedChanged, RadioButton25.CheckedChanged, RadioButton26.CheckedChanged, RadioButton32.CheckedChanged, RadioButton33.CheckedChanged, RadioButton34.CheckedChanged, RadioButton35.CheckedChanged, RadioButton36.CheckedChanged, RadioButton37.CheckedChanged, RadioButton38.CheckedChanged, RadioButton39.CheckedChanged, RadioButton40.CheckedChanged, RadioButton41.CheckedChanged

        Try

            Dim rb As RadioButton = DirectCast(sender, RadioButton)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then

                If rb.Checked Then
                    rb.Checked = False
                    Commands.SendOnOff(DirectCast(rb.Tag, OnOffCommand))
                    SetControlWaiting(rb)
                End If

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    ' Events are assigned in AddHandler method
    Private Sub HandleRangeChangeButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim cb As ComboBox
        Dim rcb As RangeChangeButton = DirectCast(sender, RangeChangeButton)

        For Each c As Control In rcb.Parent.Controls

            If TypeOf c Is ComboBox Then
                cb = DirectCast(c, ComboBox)
                Exit For
            End If

        Next

        If Not cb Is Nothing Then

            Select Case rcb.ButtonType

                Case RangeChangeButtonType.MasterVolume, RangeChangeButtonType.ChannelVolume
                    cb.SelectedIndex = CalculateVolumeIndex(cb, rcb)

                Case RangeChangeButtonType.Normal
                    cb.SelectedIndex = CalculateRangeIndex(cb, rcb)

            End Select


        End If

    End Sub

    ' Events are assigned in iConfigureFixedListComboBoxes method
    Private Sub HandleFixedListComboBoxChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendThis(cb.Tag.ToString & cb.SelectedValue.ToString)
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub HandleRemoteAndPanelLockClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton27.CheckedChanged, RadioButton28.CheckedChanged, RadioButton29.CheckedChanged, RadioButton30.CheckedChanged, RadioButton31.CheckedChanged

        Try

            Dim rb As RadioButton = DirectCast(sender, RadioButton)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then

                If rb.Checked Then
                    rb.Checked = False

                    Select Case rb.Name
                        Case "RadioButton27"
                            Commands.SendRemoteAndPanelLock(RemoteAndPanelLock.RemoteLock)
                        Case "RadioButton28"
                            Commands.SendRemoteAndPanelLock(RemoteAndPanelLock.RemoteUnlock)
                        Case "RadioButton29"
                            Commands.SendRemoteAndPanelLock(RemoteAndPanelLock.PanelFullLock)
                        Case "RadioButton30"
                            Commands.SendRemoteAndPanelLock(RemoteAndPanelLock.PanelUnlock)
                        Case "RadioButton31"
                            Commands.SendRemoteAndPanelLock(RemoteAndPanelLock.PanelFullLockExceptMV)
                    End Select

                    SetControlWaiting(rb)
                End If

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub MainMuteChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged, CheckBox10.CheckedChanged

        Try

            Dim cb As CheckBox = DirectCast(sender, CheckBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then

                If cb.Checked Then
                    Commands.SendOnOff(OnOffCommand.MuteOn)
                Else
                    Commands.SendOnOff(OnOffCommand.MuteOff)
                End If
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub MainSourceChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged, ComboBox57.SelectedIndexChanged

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendSource(DirectCast(cb.SelectedItem, Source))
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub MasterVolumeChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged, ComboBox58.SelectedIndexChanged

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendMasterVolume(DirectCast(cb.SelectedItem, MasterVolume))
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub MiniUIToFullUIClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button92.Click, Button101.Click, Button102.Click
        booLoading = True
        MenuStrip1.Visible = True
        StatusStrip1.Visible = True
        TabControl2.Visible = False
        TabControl2.Dock = DockStyle.None
        TabControl2.SendToBack()
        TabControl1.BringToFront()
        TabControl1.Visible = True

        Me.AutoScroll = True
        Me.MinimumSize = New Size(400, 307)
        Me.MaximumSize = New Size(800, 614)
        Me.Size = New Size(800, 614)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        Me.Text = "Command2112"
        booLoading = False
    End Sub

    Private Sub Zone2MuteChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged, CheckBox11.CheckedChanged

        Try

            Dim cb As CheckBox = DirectCast(sender, CheckBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then

                If cb.Checked Then
                    Commands.SendOnOff(OnOffCommand.Zone2MuteOn)
                Else
                    Commands.SendOnOff(OnOffCommand.Zone2MuteOff)
                End If

                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Zone2SourceChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox33.SelectedIndexChanged, ComboBox59.SelectedIndexChanged

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendZoneSource(Zones.Zone2, DirectCast(cb.SelectedItem, Source))
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Zone2VolumneChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox23.SelectedIndexChanged, ComboBox60.SelectedIndexChanged

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendZoneVolume(DirectCast(cb.SelectedItem, MasterVolume), Zones.Zone2)
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Zone3MuteChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged, CheckBox12.CheckedChanged

        Try

            Dim cb As CheckBox = DirectCast(sender, CheckBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then

                If cb.Checked Then
                    Commands.SendOnOff(OnOffCommand.Zone3MuteOn)
                Else
                    Commands.SendOnOff(OnOffCommand.Zone3MuteOff)
                End If
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Zone3SourceChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox36.SelectedIndexChanged, ComboBox61.SelectedIndexChanged

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendZoneSource(Zones.Zone3, DirectCast(cb.SelectedItem, Source))
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Zone3VolumneChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox24.SelectedIndexChanged, ComboBox62.SelectedIndexChanged

        Try

            Dim cb As ComboBox = DirectCast(sender, ComboBox)

            If Not booLoading AndAlso Helper.CheckIP(False, False) Then
                Commands.SendZoneVolume(DirectCast(cb.SelectedItem, MasterVolume), Zones.Zone3)
                SetControlWaiting(cb)
            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

#End Region

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        FillScriptDetail(Nothing, Nothing)
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.NotifyIcon1.Visible = False
        MiniUIToFullUIClick(Nothing, Nothing)
    End Sub

    Private Sub NotifyIcon2_BalloonTipClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon2.BalloonTipClosed
        NotifyIcon2.Visible = False
    End Sub

    Private Sub NotifyIcon2_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon2.MouseClick
        NotifyIcon2.Visible = False
    End Sub

#Region "NumericUpDown Events"

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        If booFormLoadStarted Then
            My.CustomSettings.AutoClearLines.Instance.Value = CInt(NumericUpDown1.Value)
        End If
    End Sub

    Private Sub NumericUpDown2_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown2.GotFocus, NumericUpDown2.Click
        NumericUpDown2.Select(0, 3)
    End Sub

    Private Sub NumericUpDown3_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown3.GotFocus, NumericUpDown3.Click
        NumericUpDown3.Select(0, 3)
    End Sub

    Private Sub NumericUpDown4_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown4.GotFocus, NumericUpDown4.Click
        NumericUpDown4.Select(0, 3)
    End Sub

    Private Sub NumericUpDown5_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown5.GotFocus, NumericUpDown5.Click
        NumericUpDown5.Select(0, 3)
    End Sub

    Private Sub NumericUpDown6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown6.ValueChanged
        If booFormLoadStarted Then
            My.CustomSettings.AutoRefreshNetSecs.Instance.Value = CInt(NumericUpDown6.Value)
            SetNetRefeshTimer()
        End If
    End Sub

    Private Sub NumericUpDown7_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown7.ValueChanged
        If booFormLoadStarted Then
            My.CustomSettings.DefaultCommandPauseMS.Instance.Value = CInt(NumericUpDown7.Value)
        End If
    End Sub

    Private Sub NumericUpDown8_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown8.ValueChanged
        If booFormLoadStarted Then
            My.CustomSettings.AutoRefreshIPodSecs.Instance.Value = CInt(NumericUpDown8.Value)
            SetIPodRefeshTimer()
        End If
    End Sub

#End Region

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        If Not booLoading Then

            For Each tpi As TabPageInfo In lTabInfo

                If TabControl1.SelectedTab.Name = tpi.TabName Then

                    If tpi.QueryOnSelect Then

                        Select Case tpi.TabName
                            Case tabMain.Name
                                Query(QueryType.Main)
                            Case tabSurroundParams.Name
                                Query(QueryType.SurroundParams)
                            Case tabVideoParams.Name
                                Query(QueryType.VideoParams)
                            Case tabOtherParams.Name
                                Query(QueryType.OtherParams)
                            Case tabZone2.Name
                                Query(QueryType.Zone2)
                            Case tabZone3.Name
                                Query(QueryType.Zone3)
                            Case tabTuner.Name
                                Query(QueryType.Tuner)
                            Case tabNet.Name
                                Query(QueryType.NET)
                            Case tabIPod.Name
                                Query(QueryType.iPod)
                        End Select

                    End If

                    SetNetRefeshTimer()
                    SetIPodRefeshTimer()

                    Exit For

                End If

            Next

        End If

    End Sub

    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyUp
        If e.KeyCode = Keys.Enter Then
            Button2.PerformClick()
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyUp
        If e.KeyCode = Keys.A AndAlso e.Control Then
            TextBox4.SelectAll()
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

        Dim strX() As String

        If My.CustomSettings.AutoClearDebug.Instance.Value AndAlso TextBox4.Lines.Length > My.CustomSettings.AutoClearLines.Instance.Value Then

            strX = TextBox4.Lines
            TextBox4.Clear()

            For intX As Integer = strX.Length - 11 To strX.Length - 1
                If Not String.IsNullOrEmpty(strX(intX)) Then
                    WriteDebug(strX(intX), True)
                End If
            Next

        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Button51.PerformClick()
    End Sub

    Private Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Button74.PerformClick()
    End Sub

    Private Sub Timer3_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer3.Tick

        Dim l As New Generic.List(Of GroupBox)
        Dim l2 As New Generic.List(Of GroupBox)
        Dim gb As GroupBox

        SyncLock objSL

            Timer3.Stop()

            For Each gb In lGroupBoxes

                If lControlQueue.ContainsKey(gb.Name) AndAlso lControlQueue(gb.Name).AddSeconds(5) < Now Then
                    ClearControlWaiting(gb)
                    WriteDebug(Helper.GetNowString & " - No response received in at least five seconds for: " & gb.Text & ". This is most likely because the command isn't valid for the Denon's current state. Control color being changed from red to default color.", True)

                    If booShowNotifyBalloonTip2 Then
                        NotifyIcon2.Visible = True
                        NotifyIcon2.ShowBalloonTip(5)
                        booShowNotifyBalloonTip2 = False
                    End If

                End If

            Next

            Timer3.Start()

        End SyncLock

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick

        SyncLock objSL

            Timer4.Stop()

            Dim c As New Schedule.NextRunDateComparer
            Dim datNow As Date = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)
            Dim datThen As Date
            Dim intInterval As Integer
            Dim datRun As Date

            lScriptSchedules.Sort(c)

            For Each s As Schedule In lScriptSchedules

                ' Calling NextRunDate will disable the schedule if the date occurs in the past
                datRun = s.GetNextRunDate(datNow)
                datRun = New Date(datRun.Year, datRun.Month, datRun.Day, datRun.Hour, datRun.Minute, datRun.Second)

                If Not s.IsDisabled Then

                    If datRun = datNow Then
                        WriteDebug(Helper.GetNowString & " - BEGIN SCRIPT: " & DB.Script.ReadScriptName(s.ScriptID) & " (" & s.GetCode & ")", True)
                        s.BeginRun(datNow)
                    ElseIf datRun > datNow Then
                        Exit For
                    End If

                End If

            Next

            datThen = Now.AddSeconds(1)
            datThen = New Date(datThen.Year, datThen.Month, datThen.Day, datThen.Hour, datThen.Minute, datThen.Second).AddMilliseconds(15)
            intInterval = CInt(datThen.Subtract(Now).TotalMilliseconds)
            If intInterval = 0 Then
                intInterval = 50
            End If
            Timer4.Interval = intInterval
            Timer4.Start()

        End SyncLock

    End Sub

#Region "Tool Strip Events"

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub AllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.Full)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub IPodToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IPodToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.iPod)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub MainToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MainToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.Main)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub MiniUIToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MiniUIToolStripMenuItem.Click
        booLoading = True
        Me.tabMain.Focus()
        MenuStrip1.Visible = False
        StatusStrip1.Visible = False
        TabControl1.Visible = False
        TabControl2.Parent = Me
        TabControl2.Location = New Point(0, 0)
        TabControl2.Dock = DockStyle.Fill
        TabControl1.SendToBack()
        TabControl2.BringToFront()
        TabControl2.Visible = True

        Me.AutoScroll = False
        Me.MinimumSize = New Size(414, 320)
        Me.MaximumSize = New Size(414, 320)
        Me.Size = New Size(414, 320)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog

        Me.Text = "Command2112 Mini"
        booLoading = False
    End Sub

    Private Sub MiniUIToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MiniUIToolStripMenuItem1.Click
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.NotifyIcon1.Visible = False
        MiniUIToolStripMenuItem_Click(Nothing, Nothing)
        Me.Location = New Point(MousePosition.X - Me.Width, MousePosition.Y - Me.Height)
    End Sub

    Private Sub NETToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NETToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.NET)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub NotifyExitMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyExitMenuItem.Click
        Me.Close()
    End Sub

    Private Sub NotifyRestoreMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyRestoreMenuItem.Click
        NotifyIcon1_MouseDoubleClick(Nothing, Nothing)
    End Sub

    Private Sub OtherToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OtherToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.OtherParams)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub TunerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TunerToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.Tuner)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try


    End Sub

    Private Sub SourcesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SourcesToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.Sources)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub SurroundParamsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SurroundParamsToolStripMenuItem.Click
        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.SurroundParams)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub VideoParamsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VideoParamsToolStripMenuItem.Click
        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.VideoParams)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try
    End Sub

    Private Sub ZonesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZonesToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.Zone2)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

    Private Sub Zone3ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Zone3ToolStripMenuItem.Click

        Try

            If Helper.CheckIP(False, False) Then

                Query(QueryType.Zone3)

            End If

        Catch ex As Exception
            OutputError(ex)
        End Try

    End Sub

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub GroupBox71_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox71.Enter

    End Sub

    Private Sub GroupBox52_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox52.Enter

    End Sub
End Class

