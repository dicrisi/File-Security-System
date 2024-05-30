Option Strict Off
Option Explicit On 
Imports System.Data.SqlClient
Friend Class MDIForm1
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents btnNew As System.Windows.Forms.ToolBarButton
    Friend WithEvents btnView As System.Windows.Forms.ToolBarButton
    Friend WithEvents btnExtract As System.Windows.Forms.ToolBarButton
    Friend WithEvents btnAbout As System.Windows.Forms.ToolBarButton
    Friend WithEvents btnExit As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBar As System.Windows.Forms.ToolBar
    Friend WithEvents ImageList As System.Windows.Forms.ImageList
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MDIForm1))
        Me.ToolBar = New System.Windows.Forms.ToolBar
        Me.btnNew = New System.Windows.Forms.ToolBarButton
        Me.btnView = New System.Windows.Forms.ToolBarButton
        Me.btnExtract = New System.Windows.Forms.ToolBarButton
        Me.btnAbout = New System.Windows.Forms.ToolBarButton
        Me.btnExit = New System.Windows.Forms.ToolBarButton
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.SuspendLayout()
        '
        'ToolBar
        '
        Me.ToolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.btnNew, Me.btnView, Me.btnExtract, Me.btnAbout, Me.btnExit})
        Me.ToolBar.DropDownArrows = True
        Me.ToolBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolBar.ImageList = Me.ImageList
        Me.ToolBar.Location = New System.Drawing.Point(0, 0)
        Me.ToolBar.Name = "ToolBar"
        Me.ToolBar.ShowToolTips = True
        Me.ToolBar.Size = New System.Drawing.Size(555, 58)
        Me.ToolBar.TabIndex = 4
        '
        'btnNew
        '
        Me.btnNew.ImageIndex = 0
        Me.btnNew.Text = "New"
        '
        'btnView
        '
        Me.btnView.ImageIndex = 1
        Me.btnView.Text = "View"
        '
        'btnExtract
        '
        Me.btnExtract.ImageIndex = 2
        Me.btnExtract.Text = "Extract"
        '
        'btnAbout
        '
        Me.btnAbout.ImageIndex = 3
        Me.btnAbout.Text = "About"
        '
        'btnExit
        '
        Me.btnExit.ImageIndex = 4
        Me.btnExit.Text = "Exit"
        '
        'ImageList
        '
        Me.ImageList.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'MDIForm1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(555, 448)
        Me.Controls.Add(Me.ToolBar)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.IsMdiContainer = True
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "MDIForm1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "File Protector"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
#End Region

    Dim fso As Scripting.FileSystemObject
    Dim NumFiles As String
    Dim pw As String
    Dim fp As String
    Private Sub MDIForm1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.Visible = False
        Dim f As New frmLogin
        f.ShowDialog()
        If LoginSucceed = False Then
            End
        End If
        If LoginSucceed = True Then
            Me.Visible = True
        End If
    End Sub
    Private Sub ToolBar_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar.ButtonClick
        If (e.Button.Text = "New") Then
            Module1.FileOpenMode = "New"
            frmAdd.DefInstance.Show()
        ElseIf (e.Button.Text = "Extract") Then
            fso = New Scripting.FileSystemObject
            Module1.FileOpenMode = "Extract"
            Dim ofd As New OpenFileDialog
            ofd.InitialDirectory = "C:\"
            ofd.Filter = "File Protector (*.FPP)|*.FPP"
            ofd.ShowDialog()
            If ofd.FileName <> "" Then
                If fso.FileExists(ofd.FileName) Then
                    FileOpen(1, ofd.FileName, OpenMode.Binary)
                    FileGet(1, p)
                    NumFiles = CStr(p.NumberOfFiles)
                    If p.IsProjectPasswordFound = "1" Then
                        Dim PAA As New PASSWORD
                        PAA.ShowDialog()
                        pw = GET_PASS()
                        If pw <> Trim(p.Password) Then
                            MsgBox("Wrong Password", MsgBoxStyle.Critical)
                            FileClose(1)
                            fso = Nothing
                            Exit Sub
                        End If
                    End If
                    frmAdd.DefInstance.Show()
                    Dim cnt As Integer
                    ReDim h(CInt(NumFiles) - 1)
                    For cnt = 1 To CInt(NumFiles)
                        FileGet(1, h(cnt - 1))
                        With frmAdd.DefInstance
                            .Text2.Text = Trim(p.Password)
                            .Text3.Text = ofd.FileName
                            fp = Trim(h(cnt - 1).FilePath)
                            .List1.Items.Add(fp)
                            .Combo1.Items.Add(fp)
                            .Combo2.Items.Add(fp)
                            .Combo4.Items.Add(h(cnt - 1).IsEncrypted)
                            .Combo5.Items.Add(h(cnt - 1).IsPasswordProtected)
                            .Combo3.Items.Add(Trim(h(cnt - 1).Password))
                        End With
                    Next cnt
                    FileClose(1)
                    frmAdd.DefInstance.extractall()
                    fso = Nothing
                End If
            End If
        ElseIf (e.Button.Text = "View") Then
            Dim cnt As Integer
            fso = New Scripting.FileSystemObject
            Dim ofd1 As New OpenFileDialog
            ofd1.InitialDirectory = "C:\"
            ofd1.Filter = "File Protector (*.FPP)|*.FPP"
            ofd1.ShowDialog()
            If ofd1.FileName <> "" Then
                If fso.FileExists(ofd1.FileName) Then
                    FileOpen(1, ofd1.FileName, OpenMode.Binary)
                    FileGet(1, p)
                    NumFiles = CStr(p.NumberOfFiles)
                    If p.IsProjectPasswordFound = "1" Then
                        Dim PAA As New PASSWORD
                        PAA.ShowDialog()
                        pw = GET_PASS()
                        If pw <> Trim(p.Password) Then
                            MsgBox("Wrong Password")
                            FileClose(1)
                            fso = Nothing
                            Exit Sub
                        End If
                    End If
                    frmAdd.DefInstance.Show()
                    If p.IsNoFilePasswordProtected = "1" Then
                        frmAdd.DefInstance.Command9.Enabled = True
                    Else
                        frmAdd.DefInstance.Command9.Enabled = False
                    End If
                    ReDim h(CInt(NumFiles) - 1)
                    For cnt = 1 To CInt(NumFiles)
                        FileGet(1, h(cnt - 1))
                        With frmAdd.DefInstance
                            .Text2.Text = Trim(p.Password)
                            .Text3.Text = ofd1.FileName
                            fp = Trim(h(cnt - 1).FilePath)
                            .List1.Items.Add(fp)
                            .Combo1.Items.Add(fp)
                            .Combo2.Items.Add(fp)
                            .Combo4.Items.Add(h(cnt - 1).IsEncrypted)

                            .Combo5.Items.Add(h(cnt - 1).IsPasswordProtected)

                            .Combo3.Items.Add(Trim(h(cnt - 1).Password))
                        End With
                    Next cnt
                    FileClose(1)
                    fso = Nothing
                    Dim con As New SqlConnection("server=.\sqlexpress;integrated security=true;database=fileprotector")
                    con.Open()
                    Dim cmd1 As New SqlCommand("delete from datatable")
                    cmd1.Connection = con
                    cmd1.ExecuteNonQuery()
                    For cnt = 1 To CInt(NumFiles)
                        Try
                            Dim str As String
                            str = h(cnt - 1).FilePath
                            Dim ss As String
                            ss = str.Trim
                            Dim cmd As New SqlCommand("insert into datatable values(" & h(cnt - 1).FileSize & ",'" & h(cnt - 1).FileExtension & "','" & ss & "', " & h(cnt - 1).FileOffSet & " ," & h(cnt - 1).IsEncrypted & "," & h(cnt - 1).IsPasswordProtected & ")")
                            cmd.Connection = con
                            cmd.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message.ToString())
                        End Try
                    Next
                    frmView.DefInstance.Show()
                    Dim da As New SqlDataAdapter("select * from datatable", con)
                    Dim ds As New DataSet
                    da.Fill(ds)
                    frmAdd.DefInstance.Close()
                End If
            End If
        ElseIf (e.Button.Text = "About") Then
            Dim ff As New frmAbout
            ff.ShowDialog()
        ElseIf (e.Button.Text = "Exit") Then
            End
        End If
    End Sub
End Class