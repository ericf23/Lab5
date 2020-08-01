'Name: Eric Fisher
'Date 31/07/2020
'Program: Text Editor
'Program Description: This program is a very basic text editor multi-form

Option Strict On
Imports System.IO
Public Class MainForm

    ''' <summary>
    ''' This event occurs when the user clicks File/New
    ''' </summary>
    Private Sub NewClick(sender As Object, e As EventArgs) Handles mnuFileNew.Click
        'The display textbox clears
        txtDisplay.Text = String.Empty
    End Sub

    ''' <summary>
    ''' This event occurs when the user clicks File/Open
    ''' </summary>
    Private Sub OpenClick(sender As Object, e As EventArgs) Handles mnuFileOpen.Click
        'Declare variable to read data into file
        Dim inputStream As StreamReader

        'Prompt user for file to open then display files text in display textbox
        If OpenFileDialog.ShowDialog() = DialogResult.OK Then
            inputStream = New StreamReader(OpenFileDialog.FileName)
            txtDisplay.Text = inputStream.ReadToEnd()
            inputStream.Close()
        End If
    End Sub
    ''' <summary>
    ''' This event occurs when the user clicks File/SaveAs
    ''' </summary>
    Private Sub SaveAsClick(sender As Object, e As EventArgs) Handles mnuFileSaveAs.Click
        'Declare variable to write text stream as a file
        Dim outputStream As StreamWriter

        'Prompts user where to output text in display textbox into saved file
        If SaveFileDialog.ShowDialog() = DialogResult.OK Then
            outputStream = New StreamWriter(SaveFileDialog.FileName)
            outputStream.Write(txtDisplay.Text)
            outputStream.Close()
        End If
    End Sub

    ''' <summary>
    ''' This event occurs when the user clicks File/Save
    ''' </summary>
    Private Sub SaveClick(sender As Object, e As EventArgs) Handles mnuFileSave.Click

    End Sub

    ''' <summary>
    ''' This event occurs when the user clicks Help/About
    ''' </summary>
    Private Sub AboutClick(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        'Declare variable as form
        Dim aboutModal As New AboutForm
        'Show from contents in dialog
        aboutModal.ShowDialog()
    End Sub
    ''' <summary>
    ''' This event occurs when the user clicks Edit/Copy
    ''' </summary>
    Private Sub CopyClick(sender As Object, e As EventArgs) Handles mnuEditCopy.Click
        Clipboard.SetText(Me.txtDisplay.Text)
    End Sub
    ''' <summary>
    ''' This event occurs when the user clicks Edit/Cut
    ''' </summary>
    Private Sub CutClick(sender As Object, e As EventArgs) Handles mnuEditCut.Click
        Clipboard.SetText(Me.txtDisplay.Text)
        Me.txtDisplay.Text = ""
    End Sub
    ''' <summary>
    ''' This event occurs when the user clicks Edit/Paste
    ''' </summary>
    Private Sub PasteClick(sender As Object, e As EventArgs) Handles mnuEditPaste.Click
        Me.txtDisplay.Text = txtDisplay.Text.Insert(txtDisplay.SelectionStart, Clipboard.GetText)
    End Sub
End Class
