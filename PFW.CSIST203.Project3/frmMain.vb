Public Class frmMain

    Friend persister As PFW.CSIST203.Project3.Persisters.IPersistData

    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        persister = New PFW.CSIST203.Project3.Persisters.Access.AccessPersister()
    End Sub

    Private Sub BtnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        Dim selectedRow = Int32.Parse(txtRow.Text.Trim()) - 1
        Dim maximum = persister.CountRows()
        If selectedRow <= 0 Then
            ' Do Nothing
        Else
            Dim row = persister.GetRow(selectedRow)
            txtFirstname.Text = CType(row("First Name"), String)
            txtLastname.Text = CType(row("Last Name"), String)
            txtEmailAddress.Text = CType(row("E-mail Address"), String)
            txtBusinessPhone.Text = CType(row("Business Phone"), String)
            txtCompany.Text = CType(row("Company"), String)
            txtTitle.Text = CType(row("Job Title"), String)
            txtRow.Text = CType(row("ID"), String)
        End If
    End Sub

    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim maximum = persister.CountRows()
        Dim selectedRow = Int32.Parse(txtRow.Text.Trim()) + 1
        If selectedRow > maximum Then
            ' Do Nothing
        Else
            DisplayRow(selectedRow)
            txtRow.Text = selectedRow.ToString()
        End If
    End Sub

    Friend Sub DisplayRow(selectedRow As Integer)
        Dim table = persister.GetData()
        Dim row = table.Rows(selectedRow - 1)
        txtFirstname.Text = CType(row("First Name"), String)
        txtLastname.Text = CType(row("Last Name"), String)
        txtEmailAddress.Text = CType(row("E-mail Address"), String)
        txtBusinessPhone.Text = CType(row("Business Phone"), String)
        txtCompany.Text = CType(row("Company"), String)
        txtTitle.Text = CType(row("Job Title"), String)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        If Not persister Is Nothing Then
            persister.Dispose()
            persister = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Handle the File -> Open dialog box used for selecting the excel file that is utilized by the front-end
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFileDialog.InitialDirectory = System.Environment.CurrentDirectory
        OpenFileDialog.FileName = String.Empty
        OpenFileDialog.Filter = persister.FileFilter
        Dim result = OpenFileDialog.ShowDialog()
        If result = DialogResult.OK Then
            LoadFile(OpenFileDialog.FileName)
        End If
    End Sub

    Friend Sub LoadFile(selectedFile As String)
        persister.Dispose()
        persister = New PFW.CSIST203.Project3.Persisters.Access.AccessPersister(selectedFile)

        If persister.CountRows() > 0 Then

            ' enable all of the fields for editing
            txtRow.Text = "1" ' reset back to the first item in the data table
            txtFirstname.Enabled = True
            txtLastname.Enabled = True
            txtEmailAddress.Enabled = True
            txtBusinessPhone.Enabled = True
            txtCompany.Enabled = True
            txtTitle.Enabled = True
            btnSave.Enabled = True
            DisplayRow(1)

        Else ' disable all fields from editing

            txtRow.Text = "0" ' reset back to zero
            txtFirstname.Enabled = False
            txtLastname.Enabled = False
            txtEmailAddress.Enabled = False
            txtBusinessPhone.Enabled = False
            txtCompany.Enabled = False
            txtTitle.Enabled = False
            btnSave.Enabled = False

            ' clear out all of the fields
            txtFirstname.Text = String.Empty
            txtLastname.Text = String.Empty
            txtEmailAddress.Text = String.Empty
            txtBusinessPhone.Text = String.Empty
            txtCompany.Text = String.Empty
            txtTitle.Text = String.Empty

        End If
    End Sub

    Friend Function ValidateLength(sender As Control) As Boolean
        If String.IsNullOrWhiteSpace(sender.Text) Then
            ErrorProvider.SetError(sender, "Value must be non-whitespace and non-empty")
            Return False
        Else ' the text entered is non-whitespace
            ErrorProvider.SetError(sender, String.Empty)
            Return True
        End If
    End Function

    Private Sub KeyPress_Handler(sender As Object, e As KeyPressEventArgs) Handles txtFirstname.KeyPress, txtLastname.KeyPress, txtBusinessPhone.KeyPress, txtCompany.KeyPress, txtEmailAddress.KeyPress, txtTitle.KeyPress
        Dim control = CType(sender, Control)
        ValidateLength(control)
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim control = CType(sender, Control)
        If ValidateLength(control) Then

            ' retrieve the existing row from the persistent medium
            Dim row = persister.GetRow(Integer.Parse(txtRow.Text.Trim()))

            ' change the column data of the row
            row("First Name") = txtFirstname.Text
            row("Last Name") = txtLastname.Text
            row("E-mail Address") = txtEmailAddress.Text
            row("Business Phone") = txtBusinessPhone.Text
            row("Company") = txtCompany.Text
            row("Job Title") = txtTitle.Text

            ' propagate the row back to the persister for updating
            persister.StoreRow(row)

        End If

    End Sub

End Class
