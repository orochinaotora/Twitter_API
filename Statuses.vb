Public Class Statuses
    Public Core As CoreTweet.Tokens

    Public Sub New(Core As CoreTweet.Tokens)
        Me.Core = Core
    End Sub
    Public Sub Send(Text As String)
        Core.Statuses.Update(Text)
    End Sub
End Class
