Public Class RdButton
    Private Vchecked As Boolean
    Public Property Checked() As Boolean
        Get
            Checked = Vchecked
        End Get
        Set(ByVal value As Boolean)
            Vchecked = value
            If Checked Then
                Me.BackgroundImage = Global.MyVisualObjects.My.Resources.RdOn
            Else
                Me.BackgroundImage = Global.MyVisualObjects.My.Resources.RdOff
            End If
        End Set
    End Property

    Private Sub RdButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        If Vchecked Then
            Vchecked = False
            Me.BackgroundImage = Global.MyVisualObjects.My.Resources.RdOff
        Else
            Vchecked = True
            Me.BackgroundImage = Global.MyVisualObjects.My.Resources.RdOn
        End If
        RaiseEvent DoAction()
    End Sub

    Public Event DoAction()


End Class
