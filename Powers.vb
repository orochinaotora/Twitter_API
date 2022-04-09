''' <summary>
''' TwitterAPIの接続・切断や独自の電源オプションを操作します
''' </summary>
Public Class Powers
    ''' <summary>
    ''' APIの接続許可の変数です
    ''' <para>True：許可</para>
    ''' <para>False：拒否</para>
    ''' </summary>
    Private Property PowerMode As Boolean
    ''' <summary>
    ''' APIに接続しているかの変数です
    ''' <para>True：接続済み</para>
    ''' <para>False：未接続</para>
    ''' </summary>
    Private Property IsConnected As Boolean

    ''' <summary>
    ''' TwitterAPIのTokenです
    ''' </summary>
    Public Property Token As CoreTweet.OAuth.OAuthSession

    ''' <summary>
    ''' TwitterAPIの操作ができる変数です
    ''' </summary>
    Public Property Core As CoreTweet.Tokens

    Public Enum WhereStart
        GetURI = 0
        PIN = 1
        GetToken = 2
    End Enum

#Region "独自の電源オプション"

    ''' <summary>
    ''' APIの接続許可を取得します
    ''' </summary>
    ''' <returns>返り値の型はBooleanです。
    ''' <para>True：許可</para>
    ''' <para>False：拒否</para></returns>
    Public Function GetPowerMode() As Boolean
        Return PowerMode
    End Function

    ''' <summary>
    ''' APIの接続を切り替えます。例えば、APIの接続が許可だった場合、拒否に変更されます
    ''' </summary>
    Public Sub ChangePowerMode()
        Select Case PowerMode
            Case True
                PowerMode = False
            Case False
                PowerMode = True
            Case Else
                PowerMode = False
        End Select
    End Sub

    ''' <summary>
    ''' APIの接続を許可に変更します。既に許可になっていた場合は例外エラーが発生します
    ''' </summary>
    Public Sub Start()
        Select Case PowerMode
            Case True
                Err.Raise(513, "Twitter_API\Powers\Start", "既にオンになっています")
            Case False
                PowerMode = True
            Case Else
                Err.Raise(513, "Twitter_API\Powers\Start", "状態が不明です")
        End Select
    End Sub

    ''' <summary>
    ''' APIの接続を許可に変更します。「ShowError」の引数で例外エラーの発生を変更できます
    ''' </summary>
    ''' <param name="ShowError">True：発生する、False：発生しない</param>
    Public Sub Start(ShowError As Boolean)
        Select Case PowerMode
            Case True
                If ShowError = False Then Exit Sub
                Err.Raise(513, "Twitter_API\Powers\Start", "既にオンになっています")
            Case False
                PowerMode = True
            Case Else
                If ShowError = False Then Exit Sub
                Err.Raise(513, "Twitter_API\Powers\Start", "状態が不明です")
        End Select
    End Sub

    ''' <summary>
    ''' APIの接続を拒否に変更します。既に拒否になっていた場合は例外エラーが発生します
    ''' </summary>
    Public Sub [Stop]()
        Select Case PowerMode
            Case True
                PowerMode = False
            Case False
                Err.Raise(513, "Twitter_API\Powers\Start", "既にオフになっています")
            Case Else
                Err.Raise(513, "Twitter_API\Powers\Start", "状態が不明です")
        End Select
    End Sub

    ''' <summary>
    ''' APIの接続を拒否に変更します。「ShowError」の引数で例外エラーの発生を変更できます
    ''' </summary>
    ''' <param name="ShowError">True：発生する、False：発生しない</param>
    Public Sub [Stop](ShowError As Boolean)
        Select Case PowerMode
            Case True
                PowerMode = False
            Case False
                Err.Raise(513, "Twitter_API\Powers\Start", "既にオフになっています")
            Case Else
                Err.Raise(513, "Twitter_API\Powers\Start", "状態が不明です")
        End Select
    End Sub

#End Region

    ''' <summary>
    ''' Twitter APIに接続します。
    ''' <para>※接続するアカウントに気を付けてください</para>
    ''' <see href="https://developer.twitter.com/en/portal/dashboard">Twitter Developer Portal</see>を確認してください
    ''' </summary>
    ''' <param name="ConsumerKey">Twitter Developer Portalに表示されているKey</param>
    ''' <param name="ConsumerKeySecret">Twitter Developer Portalに表示されているKey</param>
    Public Sub Connect(ConsumerKey As String, ConsumerKeySecret As String)
        If IsNothing(ConsumerKey) Or IsNothing(ConsumerKeySecret) Then Err.Raise(513, "Twitter_API\Powers\Start", "ConsumerKeysが入力されていません")
        Token = CoreTweet.OAuth.Authorize(ConsumerKey, ConsumerKeySecret)
        Dim URI As String = Token.AuthorizeUri.AbsoluteUri
        Diagnostics.Process.Start(URI)
        Dim PIN As String = InputBox("PINコードを入力してください。", "PINコードが必要です")
        Core = CoreTweet.OAuth.GetTokens(Token, PIN)
    End Sub

    Public Sub Connect(WhereStart As WhereStart, Optional ConsumerKey As String = Nothing, Optional ConsumerKeySecret As String = Nothing, Optional URI As String = Nothing, Optional PIN As String = Nothing)
        Select Case WhereStart
            Case WhereStart.GetURI
                If IsNothing(ConsumerKey) Or IsNothing(ConsumerKeySecret) Then Err.Raise(513, "Twitter_API\Powers\Start", "ConsumerKeysが入力されていません")
                Token = CoreTweet.OAuth.Authorize(ConsumerKey, ConsumerKeySecret)
                URI = Token.AuthorizeUri.AbsoluteUri
                If IsNothing(URI) Then Err.Raise(513, "Twitter_API\Powers\Start", "URIの取得が出来ていません。")
                Diagnostics.Process.Start(URI)
                Core = CoreTweet.OAuth.GetTokens(Token, PIN)
            Case WhereStart.PIN
                Diagnostics.Process.Start(URI)
                Core = CoreTweet.OAuth.GetTokens(Token, PIN)
            Case WhereStart.GetToken
                Core = CoreTweet.OAuth.GetTokens(Token, PIN)
            Case Else
                Err.Raise(513, "Twitter_API\Powers\Connect", "未知な種類が選ばれました")
        End Select

    End Sub
End Class
