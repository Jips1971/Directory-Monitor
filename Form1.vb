Imports System.IO
Imports System.Reflection.Metadata.Ecma335

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize all labels

        'Version No set Here
        Label41.Text = "v1.4"

        Dim labels As Label() = {Label19, Label20, Label21, Label22, Label23, Label24,
                         Label25, Label26, Label27, Label28, Label29, Label30,
                         Label31, Label32, Label33, Label34, Label35, Label36}

        For Each lbl As Label In labels
            lbl.Text = ""
        Next


        Timer1.Interval = 1000
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Define the directory paths for Prod, Test, and Dev
        Dim basePaths As String() = {
            "\\GS0758\d$\RAPDATA\SAPtoRAP",
            "\\GS0757\d$\RAPDATA\SAPtoRAP",
            "\\GS0756\d$\RAPDATA\SAPtoRAP"
        }

        ' Define locations for each set of directories
        Dim locations As String() = {
            "Beringen", "Burghausen", "Kallo", "Porvoo", "Schwechat", "Stenungsund"
        }

        ' Define a start index for labels (starting from Label19)
        Dim labelIndex As Integer = 19

        ' Iterate through each base path (Prod, Test, Dev)
        For Each basePath As String In basePaths
            ' Iterate through each location
            For Each location As String In locations
                ' Construct the full directory path
                Dim fullPath As String = Path.Combine(basePath, location)

                ' Check if the directory exists and update the corresponding label
                If Directory.Exists(fullPath) Then
                    Me.Controls("Label" & labelIndex).ForeColor = Color.Green

                    Me.Controls("Label" & labelIndex).Text = Directory.GetFiles(fullPath, "*.*").Count.ToString()
                    If Directory.GetFiles(fullPath, "*.*").Count.ToString() > 0 Then
                        Me.Controls("Label" & labelIndex).ForeColor = Color.Red
                        'Flash Me
                        Me.Controls("Label" & labelIndex).Visible = Not Me.Controls("Label" & labelIndex).Visible
                    Else
                        Me.Controls("Label" & labelIndex).ForeColor = Color.Green
                        Me.Controls("Label" & labelIndex).Visible = True
                    End If




                Else
                    Me.Controls("Label" & labelIndex).ForeColor = Color.Red
                    Me.Controls("Label" & labelIndex).Text = "N/A"
                End If

                labelIndex += 1
            Next
        Next
    End Sub





End Class
