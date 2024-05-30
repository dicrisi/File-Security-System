Option Strict Off
Option Explicit On
Module Module1
    Public FileOpenMode As String
    Public PASS As String = ""


	
	Public Structure ProjectHeader
		Dim NumberOfFiles As Integer
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=1)> Public IsProjectPasswordFound As String

		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=20)> Public Password As String
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=1)> Public IsNoFilePasswordProtected As String
	End Structure
	'Size 26 bytes
	
	Public Structure FileHeader
		Dim FileSize As Integer
		<VBFixedString(4),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=4)> Public FileExtension As String
		<VBFixedString(255),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=255)> Public FilePath As String
		Dim FileOffSet As Double
		'FileLength As Double
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=1)> Public IsEncrypted As String
		<VBFixedString(1),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=1)> Public IsPasswordProtected As String
        <VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=20)> Public Password As String
        <VBFixedString(7), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=7)> Public Reserved1 As String
	End Structure
	'Size 300
	
	Public p As Module1.ProjectHeader
    Public h() As Module1.FileHeader
    Public Sub PUT_PASS(ByVal VAL As String)
        PASS = VAL
    End Sub
    Public Function GET_PASS() As String
        GET_PASS = PASS
    End Function

End Module