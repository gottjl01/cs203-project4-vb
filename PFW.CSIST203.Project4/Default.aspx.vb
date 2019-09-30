Public Class _Default
    Inherits System.Web.UI.Page

    Sub New()
        ' TODO: Implement
    End Sub

    Public Property CurrentRow As Integer
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Integer)
            ' TODO: Implement
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' TODO: Implement page load logic here

    End Sub

    Protected Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click

        ' Implement previous button click here

    End Sub

    Protected Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

        ' implement next button click here

    End Sub

    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        ' Implement save click button logic here

    End Sub

    Protected Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click

        ' Implement reset button click here

    End Sub

    Protected Sub btnNewEntry_Click(sender As Object, e As EventArgs) Handles btnNewEntry.Click

        ' Implement new entry click

    End Sub

    Private Sub UpdateDisplay()

        ' Implement logic to refresh the data in the textboxes

    End Sub

End Class