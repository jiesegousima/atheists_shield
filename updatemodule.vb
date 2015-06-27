
Private Declare Function InternetOpenA Lib "wininet.dll" (ByVal sAgent As String, ByVal lAccessType As Integer, ByVal lFlags As String)
Private Declare Function InternetOpenUrlA Lib "wininet.dll" (ByVal hInet As Integer, ByVal lpszUrl As String, ByVal dwflags As Integer)
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hfile As Integer, ByVal lpbuffer As String, ByVal lNumBytesToRead As Integer, ByVal lNumberOfBytesRead As String)

Private Sub Form_Load()
Dim txt As String
Dim hwnd1 As Integer
Dim hwnd2 As Integer
Dim hwnd3 As Integer
Dim hwnd4 As Integer
Dim buffer As String
hwnd1 = InternetOpenA(0, 1, 0)
hwnd2 = InternetOpenUrlA(hwnd1, "https://github.com/jiesegousima/atheists_shield/blob/master/cultshda.txt", 2147483650#)
hwnd3 = InternetReadFile(hwnd2, buffer, 65536, 65536)
txt = CurDir() + "cultshda.txt"
Dim fs As Object
Dim a As Object
Set fs = CreateObject("scripting.filesystemobject")
Set a = fs.createtxtfile(txt, True)
a.writeline buffer
a.Close
Set a = Nothing
hwnd1 = InternetOpenA(0, 1, 0)
hwnd2 = InternetOpenUrlA(hwnd1, "https://github.com/jiesegousima/atheists_shield/blob/master/cultshdc.txt", 2147483650#)
hwnd3 = InternetReadFile(hwnd2, buffer, 65536, 65536)
txt = CurDir() + "cultshdc.txt"
Set fs = CreateObject("scripting.filesystemobject")
Set a = fs.createtxtfile(txt, True)
a.writeline buffer
a.Close
Set a = Nothing
hwnd1 = InternetOpenA(0, 1, 0)
hwnd2 = InternetOpenUrlA(hwnd1, "https://github.com/jiesegousima/atheists_shield/blob/master/cultshdb.txt", 2147483650#)
hwnd3 = InternetReadFile(hwnd2, buffer, 65536, 65536)
txt = CurDir() + "cultshdb.txt"
Set fs = CreateObject("scripting.filesystemobject")
Set a = fs.createtxtfile(txt, True)
a.writeline buffer
a.Close
Set a = Nothing
End Sub

