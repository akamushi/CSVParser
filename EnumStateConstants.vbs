'/** CsvReader クラス */
Option Explicit


Class CsvReader

    ' プロパティ変数
    Private m_FilePath    'As String
    Private m_HasHeader   'As Boolean
    Private m_IgnoreError 'As Boolean


    ' Private フィールド
    Private mTextStream   'As TextStream
    Private mState        'As EnumState
    Private mHeaders      'As Dictionary


    ' 列挙体の代わり
    Private EnumState     'As EnumStateConstants


    ' コンストラクタ
    Private Sub Class_Initialize()
        Set EnumState = New EnumStateConstants
        mState = EnumState.None
    End Sub


    ' デストラクタ
    Private Sub Class_Terminate()
        Call Me.CloseStream()
    End Sub


    ' FilePath プロパティ - Getter
    Public Property Get FilePath() 'As String
        FilePath = m_FilePath
    End Property


    ' HasHeader プロパティ - Getter
    Public Property Get HasHeader() 'As Boolean
        HasHeader = m_HasHeader
    End Property


    ' IgnoreError プロパティ - Getter
    Public Property Get IgnoreError() 'As Boolean
        IgnoreError = m_IgnoreError
    End Property


    ' IgnoreError プロパティ - Setter
    Public Property Let IgnoreError(ByVal Value) 'As Boolean
        m_IgnoreError = Value
    End Property


    ' EndOfStream プロパティ
    Public Property Get EndOfStream() 'As Boolean
        EndOfStream = mTextStream.AtEndOfStream
    End Property


    ' OpenStream メソッド
    Public Function OpenStream(ByVal stFilePath) 'As Boolean
        On Error Resume Next
        m_FilePath = stFilePath

        Dim cFso 'As FileSystemObject
        Set cFso = WScript.CreateObject("Scripting.FileSystemObject")
        Set mTextStream = cFso.OpenTextFile(Me.FilePath)

        If Err.Number = 0 Then
            OpenStream = True
            Exit Function
        End If

        Call Me.CloseStream()
    End Function


    ' CloseStream メソッド
    Public Sub CloseStream()
        If Not mTextStream Is Nothing Then
            On Error Resume Next
            Call mTextStream.Close()
            On Error GoTo 0
        End If
    End Sub


    ' ReadHeader メソッド
    Public Function ReadHeader() 'As Dictionary
        Set mHeaders = Me.ReadLine()
        m_HasHeader = True
        Set ReadHeader = mHeaders
    End Function


    ' ReadLine メソッド
    Public Function ReadLine() 'As Dictionary
        Do While (True)
            Dim stReadLine 'As String
            stReadLine = stReadLine & mTextStream.ReadLine()

            Dim cRow 'As Dictionary
            Set cRow = ReadLineInternal(stReadLine)

            Select Case mState
                Case EnumState.FindQuote, EnumState.InQuote
                    stReadLine = stReadLine & vbNewLine
                Case Else
                    Exit Do
            End Select
        Loop

        Set ReadLine = cRow
    End Function


    ' ReadToEnd メソッド
    Public Function ReadToEnd() 'As Dictionary
        Dim cTable 'As Dictionary
        Set cTable = WScript.CreateObject("Scripting.Dictionary")

        Dim stReadAll 'As String
        stReadAll = mTextStream.ReadAll()

        Dim stReadLines 'As String
        stReadLines = Split(stReadAll, vbNewLine)

        Dim stReadLine 'As String
        Dim i          'As Integer
        Dim iIndex     'As Integer

        For i = LBound(stReadLines) To UBound(stReadLines)
            stReadLine = stReadLine & stReadLines(i)

            Dim cRow 'As Dictionary
            Set cRow = ReadLineInternal(stReadLine)

            Select Case mState
                Case EnumState.FindQuote, EnumState.InQuote
                    stReadLine = stReadLine & vbNewLine
                Case Else
                    stReadLine = ""
                    iIndex = iIndex + 1
                    Call cTable.Add(iIndex, cRow)
            End Select
        Next

        Set ReadToEnd = cTable
    End Function


    ' 1 行読み込み
    Private Function ReadLineInternal(ByVal stBuffer) 'As Dictionary
        Dim cRow 'As Dictionary
        Set cRow = WScript.CreateObject("Scripting.Dictionary")

        mState = EnumState.Beginning

        Dim stItem 'As String
        Dim iIndex 'As Integer
        Dim iSeek  'As Integer

        For iSeek = 1 To Len(stBuffer)
            Dim chNext 'As String
            chNext = Mid(stBuffer, iSeek, 1)

            Select Case mState
                Case EnumState.Beginning
                    stItem = ReadForStateBeginning(stItem, chNext)
                Case EnumState.WaitInput
                    stItem = ReadForStateWaitInput(stItem, chNext)
                Case EnumState.FindQuote
                    stItem = ReadForStateFindQuote(stItem, chNext)
                Case EnumState.FindQuoteDouble
                    stItem = ReadForStateFindQuoteDouble(stItem, chNext)
                Case EnumState.InQuote
                    stItem = ReadForStateInQuote(stItem, chNext)
                Case EnumState.FindQuoteInQuote
                    stItem = ReadForStateFindQuoteInQuote(stItem, chNext)
            End Select

            Select Case mState
                Case EnumState.FindCrLf
                    mState = EnumState.Beginning
                    Exit For
                Case EnumState.FindComma
                    Call AddRowItem(stItem, cRow, iIndex)

                    mState = EnumState.Beginning
                    stItem = ""
                    iIndex = iIndex + 1
                Case EnumState.Error
                    If Not Me.IgnoreError Then
                        Call Err.Raise(5, "ReadLineInternal", "書式が不正です。")
                    End If

                    mState = EnumState.WaitInput
            End Select
        Next

        If mState = EnumState.FindQuoteDouble Then
            stItem = stItem & """"
        End If

        Call AddRowItem(stItem, cRow, iIndex)
        Set ReadLineInternal = cRow
    End Function


    ' 初回入力待ち状態での Read
    Private Function ReadForStateBeginning(ByVal stItem, ByVal chNext) 'As String
        Select Case chNext
            Case vbCr
                mState = EnumState.FindCr
            Case ","
                mState = EnumState.FindComma
            Case """"
                mState = EnumState.FindQuote
            Case Else
                mState = EnumState.WaitInput
                stItem = stItem & chNext
        End Select

        ReadForStateBeginning = stItem
    End Function


    ' 入力待ち状態での Read
    Private Function ReadForStateWaitInput(ByVal stItem, ByVal chNext) 'As String
        Select Case chNext
            Case vbCr
                mState = EnumState.FindCr
            Case ","
                mState = EnumState.FindComma
            Case """"
                mState = EnumState.FindQuote
            Case Else
                stItem = stItem & chNext
        End Select

        ReadForStateWaitInput = stItem
    End Function


    ' 引用符を発見した状態での Read
    Private Function ReadForStateFindQuote(ByVal stItem, ByVal chNext) 'As String
        Select Case chNext
            Case """"
                mState = EnumState.FindQuoteDouble
            Case Else
                mState = EnumState.InQuote
                stItem = stItem & chNext
        End Select

        ReadForStateFindQuote = stItem
    End Function


    ' 引用符の連続を発見した状態での Read
    Private Function ReadForStateFindQuoteDouble(ByVal stItem, ByVal chNext) 'As String
        Select Case chNext
            Case vbCr
                mState = EnumState.FindCr
                stItem = stItem & """"
            Case ","
                mState = EnumState.FindComma
                stItem = stItem & """"
            Case """"
                mState = EnumState.FindQuote
                stItem = stItem & """"
            Case Else
                mState = EnumState.WaitInput
                stItem = stItem & """" & chNext
        End Select

        ReadForStateFindQuoteDouble = stItem
    End Function


    ' 引用符の中で入力待ち状態での Read
    Private Function ReadForStateInQuote(ByVal stItem, ByVal chNext) 'As String
        Select Case chNext
            Case """"
                mState = EnumState.FindQuoteInQuote
            Case Else
                stItem = stItem & chNext
        End Select

        ReadForStateInQuote = stItem
    End Function


    ' 引用符の中で引用符を発見した状態での Read
    Private Function ReadForStateFindQuoteInQuote(ByVal stItem, ByVal chNext) 'As String
        Select Case chNext
            Case vbCr
                mState = EnumState.FindCr
            Case ","
                mState = EnumState.FindComma
            Case """"
                mState = EnumState.InQuote
                stItem = stItem & """"
            Case Else
                mState = EnumState.Error
        End Select

        ReadForStateFindQuoteInQuote = stItem
    End Function


    ' Row にアイテムを入れる
    Private Sub AddRowItem(ByVal stItem, ByVal cRow, ByVal iIndex)
        If Me.HasHeader Then
            Call cRow.Add(mHeaders(iIndex), stItem)
        Else
            Call cRow.Add(iIndex, stItem)
        End If
    End Sub

End Class