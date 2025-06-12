Imports System.IO
Imports System.Net.WebRequestMethods
Imports System.Reflection.Metadata.Ecma335

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize all labels

        'Version No set Here
        Label41.Text = "v2.1"

        Dim labels As Label() = {Label19, Label20, Label21, Label22, Label23, Label24,
                         Label25, Label26, Label27, Label28, Label29, Label30,
                         Label31, Label32, Label33, Label34, Label35, Label36}

        For Each lbl As Label In labels
            lbl.Text = ""
        Next
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False

        Timer1.Interval = 1000
        Timer1.Start()
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Clear all label text and set defaults
        Dim allLabels As Label() = {
        Label19, Label20, Label21, Label22, Label23, Label24,
        Label25, Label26, Label27, Label28, Label29, Label30,
        Label31, Label32, Label33, Label34, Label35, Label36
    }

        For Each lbl In allLabels
            lbl.Text = ""
            lbl.Visible = True
            lbl.ForeColor = Color.Black
        Next

        ' Store which checkboxes are selected and their corresponding label offsets
        Dim checkboxPaths As New List(Of Tuple(Of CheckBox, String, Integer))

        If CheckBox1.Checked Then
            'checkboxPaths.Add(Tuple.Create(CheckBox1, "C:\RAPDATA\Prod\SAPtoRAP", 0))
            checkboxPaths.Add(Tuple.Create(CheckBox1, "\\GS0758\d$\RAPDATA\SAPtoRAP", 0))
            CheckBox1.ForeColor = Color.Green
        Else
            CheckBox1.ForeColor = Color.Black
        End If

        If CheckBox2.Checked Then
            'checkboxPaths.Add(Tuple.Create(CheckBox2, "C:\RAPDATA\Test\SAPtoRAP", 6))
            checkboxPaths.Add(Tuple.Create(CheckBox2, "\\GS0757\d$\RAPDATA\SAPtoRAP", 6))
            CheckBox2.ForeColor = Color.Green
        Else
            CheckBox2.ForeColor = Color.Black
        End If

        If CheckBox3.Checked Then
            'checkboxPaths.Add(Tuple.Create(CheckBox3, "C:\RAPDATA\Dev\SAPtoRAP", 12))
            checkboxPaths.Add(Tuple.Create(CheckBox3, "\\GS0756\d$\RAPDATA\SAPtoRAP", 12))
            CheckBox3.ForeColor = Color.Green
        Else
            CheckBox3.ForeColor = Color.Black
        End If

        ' Loop over selected checkboxes and populate their label columns
        For Each item In checkboxPaths
            Dim path As String = item.Item2
            Dim labelOffset As Integer = item.Item3

            If Directory.Exists(path) Then
                Dim folders = Directory.GetDirectories(path)

                For i As Integer = 0 To Math.Min(folders.Length - 1, 5)
                    Dim fullPath As String = folders(i)
                    Dim folderName As String = System.IO.Path.GetFileName(System.IO.Path.TrimEndingDirectorySeparator(fullPath))
                    Dim fileCount As Integer = Directory.GetFiles(fullPath, "*.*").Length

                    Dim targetLabel As Label = allLabels(labelOffset + i)
                    targetLabel.Text = folderName

                    If fileCount > 0 Then
                        targetLabel.ForeColor = Color.Red
                    Else
                        targetLabel.ForeColor = Color.Black
                    End If
                Next



            End If
        Next
    End Sub



    Private Sub CopyFailedFiles()
        Dim locations As String() = {
        "\\GS0758\d$\RAPDATA\SAPtoRAP\Beringen\",
        "\\GS0758\d$\RAPDATA\SAPtoRAP\Kallo\",
        "\\GS0758\d$\RAPDATA\SAPtoRAP\Porvoo\",
        "\\GS0758\d$\RAPDATA\SAPtoRAP\Burghausen\",
        "\\GS0758\d$\RAPDATA\SAPtoRAP\Schwechat\",
        "\\GS0758\d$\RAPDATA\SAPtoRAP\Stenungsund\"
    }

        For Each folderPath In locations
            ProcessFolder(folderPath)
        Next
    End Sub

    Private Sub ProcessFolder(folderPath As String)
        Dim failedPath As String = Path.Combine(folderPath, "Failed")

        If Not My.Computer.FileSystem.DirectoryExists(folderPath) Then Exit Sub

        ' Delete all .error files
        For Each errorFile In Directory.GetFiles(folderPath, "*.error")
            My.Computer.FileSystem.DeleteFile(errorFile)
        Next

        ' Copy .csv files from Failed to main folder
        If My.Computer.FileSystem.DirectoryExists(failedPath) Then
            For Each csvFile In Directory.GetFiles(failedPath, "*.csv")
                Dim fileName = Path.GetFileName(csvFile)
                Dim destPath = Path.Combine(folderPath, fileName)
                My.Computer.FileSystem.CopyFile(csvFile, destPath, True)
            Next

            ' Delete the .csv files from the Failed folder after copying
            For Each copiedFile In Directory.GetFiles(failedPath, "*.csv")
                My.Computer.FileSystem.DeleteFile(copiedFile)
            Next
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CopyFailedFiles()
    End Sub
End Class
