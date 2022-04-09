Public Class Follows
    Private FollowList As CoreTweet.Cursored(Of CoreTweet.User)
    Public Token As CoreTweet.Tokens
    Public Sub New(Token As CoreTweet.Tokens)
        Me.Token = Token
    End Sub
    Private Sub AutoGetFollows()
        FollowList = Token.Friends.List
    End Sub
End Class
