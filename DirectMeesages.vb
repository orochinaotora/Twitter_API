Public Class DirectMeesages
    Public Core As CoreTweet.Tokens

    Public Sub New(Core As CoreTweet.Tokens)
        Me.Core = Core
    End Sub

    Public Function CreateReplyOption(Label As String, Description As String) As CoreTweet.QuickReplyOption
        Dim QRO As New CoreTweet.QuickReplyOption With {
            .Label = Label,
            .Description = Description
        }
        Return QRO
    End Function

    Public Function CreateReply(QRO As CoreTweet.QuickReplyOption) As CoreTweet.QuickReply
        Dim QROS As CoreTweet.QuickReplyOption()
        QROS = {QRO}
        Dim QR As New CoreTweet.QuickReply With {
            .Options = QROS,
            .Type = "options"
        }
        Return QR
    End Function

    Public Sub Send(Text As String, YourID As Long, Optional Reply As CoreTweet.QuickReply = Nothing)
        Core.DirectMessages.Events.[New](Text, YourID, Reply)
    End Sub
End Class
