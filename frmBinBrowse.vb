Public Class frmBinBrowse

    Private Sub frmBinBrowse_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        PrefSave()


    End Sub

    Private Sub frmBinBrowse_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick


    End Sub


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PrefLoad()

        'On Error GoTo ThisErrorSpot
        browser.SelectedPath = txtPath.Text

        If browser.SelectedPath <> "" Then

            For Each file As String In IO.Directory.GetFiles(browser.SelectedPath, txtFilter.Text)
                lstFiles.Items.Add(file)
            Next

            DoIndexScan()

        End If

ThisErrorSpot:
    End Sub



    Private Sub DoBin(n As Integer)
        Dim targetF As System.IO.FileInfo
        Dim FileSuffix As String
        Dim nSuffix As Int32
        Dim nStr As String
        Dim NewFile As String

        targetF = My.Computer.FileSystem.GetFileInfo(txtFile.Text)

        NewFile = ""

        If n = 1 Then
            nSuffix = Val(txtB1.Text) + 1
            FileSuffix = nSuffix.ToString("D4")
            nStr = Replace(Str(n), " ", "")
            txtB1.Text = FileSuffix
            NewFile = txtPath.Text & "\iSloth\" & nStr & "_" & FileSuffix & targetF.Extension
        ElseIf n = 2 Then
            nSuffix = Val(txtB2.Text) + 1
            FileSuffix = nSuffix.ToString("D4")
            nStr = Replace(Str(n), " ", "")
            txtB2.Text = FileSuffix
            NewFile = txtPath.Text & "\iSloth\" & nStr & "_" & FileSuffix & targetF.Extension
        ElseIf n = 3 Then
            nSuffix = Val(txtB3.Text) + 1
            FileSuffix = nSuffix.ToString("D4")
            nStr = Replace(Str(n), " ", "")
            txtB3.Text = FileSuffix
            NewFile = txtPath.Text & "\iSloth\" & nStr & "_" & FileSuffix & targetF.Extension
        ElseIf n = 4 Then
            nSuffix = Val(txtB4.Text) + 1
            FileSuffix = nSuffix.ToString("D4")
            nStr = Replace(Str(n), " ", "")
            txtB4.Text = FileSuffix
            NewFile = txtPath.Text & "\iSloth\" & nStr & "_" & FileSuffix & targetF.Extension
        ElseIf n = 5 Then
            nSuffix = Val(txtB5.Text) + 1
            FileSuffix = nSuffix.ToString("D4")
            nStr = Replace(Str(n), " ", "")
            txtB5.Text = FileSuffix
            NewFile = txtPath.Text & "\iSloth\" & nStr & "_" & FileSuffix & targetF.Extension
        ElseIf n = 6 Then
            nSuffix = Val(txtB6.Text) + 1
            FileSuffix = nSuffix.ToString("D4")
            nStr = Replace(Str(n), " ", "")
            txtB6.Text = FileSuffix
            NewFile = txtPath.Text & "\iSloth\" & nStr & "_" & FileSuffix & targetF.Extension
        End If

        If NewFile <> "" Then
            pBox.Image.Dispose()
            txtNewF.Text = NewFile

            If rbRename.Checked = True Then
                My.Computer.FileSystem.RenameFile(txtFile.Text, nStr & "_" & FileSuffix & targetF.Extension)
            ElseIf rbCopy.Checked = True Then
                My.Computer.FileSystem.MoveFile(txtFile.Text, NewFile, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.ThrowException)
            End If
        End If


    End Sub
    Private Sub ScaleImage()
        On Error GoTo ErrImg
        Dim PicBoxHeight As Integer
        Dim PicBoxWidth As Integer
        Dim ImageHeight As Integer
        Dim ImageWidth As Integer
        Dim TempImage As Image
        Dim scale_factor As Single
        Dim fs As System.IO.FileStream

        pBox.SizeMode = PictureBoxSizeMode.Normal
        PicBoxHeight = pBox.Height
        PicBoxWidth = pBox.Width
        fs = New System.IO.FileStream(txtFile.Text, IO.FileMode.Open, IO.FileAccess.Read)

        TempImage = Image.FromStream(fs)

        ImageHeight = TempImage.Height
        ImageWidth = TempImage.Width

        scale_factor = 1.0 
        If ImageHeight > PicBoxHeight Then
            scale_factor = CSng(PicBoxHeight / ImageHeight)
        End If

        If (ImageWidth * scale_factor) > PicBoxWidth Then
            scale_factor = CSng(PicBoxWidth / ImageWidth)
        End If

        pBox.Image = TempImage

        Dim bm_source As New Bitmap(pBox.Image)

        Dim bm_dest As New Bitmap( _
            CInt(bm_source.Width * scale_factor), _
            CInt(bm_source.Height * scale_factor))

        Dim gr_dest As Graphics = Graphics.FromImage(bm_dest)
        gr_dest.DrawImage(bm_source, 0, 0, _
            bm_dest.Width + 1, _
            bm_dest.Height + 1)
        pBox.Image = bm_dest
        fs.Close()
ErrImg:
    End Sub

    Private Sub cmdBrowse_Click_1(sender As Object, e As EventArgs) Handles cmdBrowse.Click
        Dim nPath As String

        browser.ShowDialog()
        If browser.SelectedPath <> "" Then
            lstFiles.Items.Clear()

            nPath = browser.SelectedPath
            txtPath.Text = nPath

            For Each file As String In IO.Directory.GetFiles(nPath, txtFilter.Text)
                lstFiles.Items.Add(file)
            Next

            DoIndexScan()

        End If
    End Sub
    Private Sub lstFiles_KeyPress(sender As Object, e As KeyPressEventArgs) Handles lstFiles.KeyPress
        DoBin(Val(e.KeyChar))
        lstFiles.Items.Remove(lstFiles.SelectedItem)
    End Sub


    Private Sub lstFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstFiles.SelectedIndexChanged
        txtFile.Text = lstFiles.Text
        ScaleImage()
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        lstFiles.Items.Clear()

        For Each file As String In IO.Directory.GetFiles(txtPath.Text, txtFilter.Text)
            lstFiles.Items.Add(file)
        Next
    End Sub

    Private Sub cmdScan_Click(sender As Object, e As EventArgs) Handles cmdScan.Click
        DoIndexScan()
    End Sub
    Sub DoIndexScan()
        Dim FileSuffix As String
        Dim CheckedFile As String
        Dim nPrefix As Integer
        Dim mCount As Integer

        For nPrefix = 1 To 6
            For mCount = 1 To 9999
                FileSuffix = mCount.ToString("D4")
                If rbCopy.Checked = True Then
                    CheckedFile = txtPath.Text & "\iSloth\" & nPrefix.ToString("D1") & "_" & FileSuffix & LCase(".JPG")
                Else
                    CheckedFile = txtPath.Text & "\" & nPrefix.ToString("D1") & "_" & FileSuffix & LCase(".JPG")
                End If


                If My.Computer.FileSystem.FileExists(CheckedFile) Then
                    If nPrefix = 1 Then
                        txtB1.Text = FileSuffix
                    ElseIf nPrefix = 2 Then
                        txtB2.Text = FileSuffix
                    ElseIf nPrefix = 3 Then
                        txtB3.Text = FileSuffix
                    ElseIf nPrefix = 4 Then
                        txtB4.Text = FileSuffix
                    ElseIf nPrefix = 5 Then
                        txtB5.Text = FileSuffix
                    ElseIf nPrefix = 6 Then
                        txtB6.Text = FileSuffix
                    End If
                Else
                    Exit For
                End If

            Next
        Next
    End Sub

    Private Sub PrefSave()
        Dim SaveString As String
        'Dim nString(5) As String
        Dim nString As String

        'nString(0) = txtPath.Text

        nString = txtPath.Text
        SaveString = ""

        'For n = 0 To 2
        'SaveString = SaveString & nString(n) '& "|"
        'Next

        SaveString = nString

        My.Computer.FileSystem.WriteAllText(Application.StartupPath & "\config.dat", SaveString, False)

    End Sub

    Private Sub PrefLoad()
        Dim LoadData As String
        'Dim nLoad(2) As String

        LoadData = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\config.dat")
        'nLoad = LoadData.Split("|")

        txtPath.Text = LoadData

    End Sub

    Private Sub pBox_Click(sender As Object, e As EventArgs) Handles pBox.Click

    End Sub

    Private Sub pBox_MouseWheel(sender As Object, e As MouseEventArgs) Handles pBox.MouseWheel
        If e.Delta <> 0 Then
            If e.Delta <= 0 Then
                If pBox.Width < 500 Then Exit Sub 'minimum 500?
            Else
                If pBox.Width > 2000 Then Exit Sub 'maximum 2000?
            End If

            pBox.Width += CInt(pBox.Width * e.Delta / 1000)
            pBox.Height += CInt(pBox.Height * e.Delta / 1000)
        End If
    End Sub
End Class
