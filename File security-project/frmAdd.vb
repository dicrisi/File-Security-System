Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAdd
	Inherits System.Windows.Forms.Form

#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            If m_InitializingDefInstance Then
                m_vb6FormDefInstance = Me
            Else
                Try
                    'For the start-up form, the first instance created is the default instance.
                    If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
                        m_vb6FormDefInstance = Me
                    End If
                Catch
                End Try
            End If
        End If
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
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents Command9 As System.Windows.Forms.Button
    Public WithEvents Command8 As System.Windows.Forms.Button
    Public WithEvents Command7 As System.Windows.Forms.Button
    Public WithEvents Text3 As System.Windows.Forms.TextBox
    Public WithEvents Combo5 As System.Windows.Forms.ComboBox
    Public WithEvents Combo4 As System.Windows.Forms.ComboBox
    Public WithEvents Combo3 As System.Windows.Forms.ComboBox
    Public WithEvents Text2 As System.Windows.Forms.TextBox
    Public WithEvents Command6 As System.Windows.Forms.Button
    Public WithEvents Check1 As System.Windows.Forms.CheckBox
    Public WithEvents Command5 As System.Windows.Forms.Button
    Public WithEvents Text1 As System.Windows.Forms.TextBox
    Public WithEvents Combo2 As System.Windows.Forms.ComboBox
    Public WithEvents Combo1 As System.Windows.Forms.ComboBox
    Public WithEvents Command4 As System.Windows.Forms.Button
    Public WithEvents List1 As System.Windows.Forms.ListBox
    Public WithEvents Command1 As System.Windows.Forms.Button
    Public WithEvents Command2 As System.Windows.Forms.Button
    Public WithEvents Command3 As System.Windows.Forms.Button
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label8 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAdd))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command9 = New System.Windows.Forms.Button
        Me.Command8 = New System.Windows.Forms.Button
        Me.Command7 = New System.Windows.Forms.Button
        Me.Text3 = New System.Windows.Forms.TextBox
        Me.Combo5 = New System.Windows.Forms.ComboBox
        Me.Combo4 = New System.Windows.Forms.ComboBox
        Me.Combo3 = New System.Windows.Forms.ComboBox
        Me.Text2 = New System.Windows.Forms.TextBox
        Me.Command6 = New System.Windows.Forms.Button
        Me.Check1 = New System.Windows.Forms.CheckBox
        Me.Command5 = New System.Windows.Forms.Button
        Me.Text1 = New System.Windows.Forms.TextBox
        Me.Combo2 = New System.Windows.Forms.ComboBox
        Me.Combo1 = New System.Windows.Forms.ComboBox
        Me.Command4 = New System.Windows.Forms.Button
        Me.List1 = New System.Windows.Forms.ListBox
        Me.Command1 = New System.Windows.Forms.Button
        Me.Command2 = New System.Windows.Forms.Button
        Me.Command3 = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Label8 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Command9
        '
        Me.Command9.BackColor = System.Drawing.SystemColors.Control
        Me.Command9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command9.Font = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command9.Image = CType(resources.GetObject("Command9.Image"), System.Drawing.Image)
        Me.Command9.Location = New System.Drawing.Point(640, 336)
        Me.Command9.Name = "Command9"
        Me.Command9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command9.Size = New System.Drawing.Size(129, 69)
        Me.Command9.TabIndex = 25
        Me.Command9.Text = "Extract All"
        Me.Command9.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command9.UseVisualStyleBackColor = False
        '
        'Command8
        '
        Me.Command8.BackColor = System.Drawing.SystemColors.Control
        Me.Command8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command8.Font = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command8.Image = CType(resources.GetObject("Command8.Image"), System.Drawing.Image)
        Me.Command8.Location = New System.Drawing.Point(536, 336)
        Me.Command8.Name = "Command8"
        Me.Command8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command8.Size = New System.Drawing.Size(97, 69)
        Me.Command8.TabIndex = 24
        Me.Command8.Text = "Extract"
        Me.Command8.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command8.UseVisualStyleBackColor = False
        '
        'Command7
        '
        Me.Command7.BackColor = System.Drawing.SystemColors.Control
        Me.Command7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command7.Location = New System.Drawing.Point(32, 424)
        Me.Command7.Name = "Command7"
        Me.Command7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command7.Size = New System.Drawing.Size(120, 48)
        Me.Command7.TabIndex = 23
        Me.Command7.Text = "Create Project File"
        Me.Command7.UseVisualStyleBackColor = False
        '
        'Text3
        '
        Me.Text3.AcceptsReturn = True
        Me.Text3.BackColor = System.Drawing.SystemColors.Window
        Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text3.Enabled = False
        Me.Text3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text3.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.Text3.Location = New System.Drawing.Point(8, 376)
        Me.Text3.MaxLength = 0
        Me.Text3.Name = "Text3"
        Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text3.Size = New System.Drawing.Size(177, 20)
        Me.Text3.TabIndex = 21
        '
        'Combo5
        '
        Me.Combo5.BackColor = System.Drawing.SystemColors.Window
        Me.Combo5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo5.Location = New System.Drawing.Point(552, 40)
        Me.Combo5.Name = "Combo5"
        Me.Combo5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo5.Size = New System.Drawing.Size(216, 22)
        Me.Combo5.TabIndex = 19
        '
        'Combo4
        '
        Me.Combo4.BackColor = System.Drawing.SystemColors.Window
        Me.Combo4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo4.Location = New System.Drawing.Point(552, 256)
        Me.Combo4.Name = "Combo4"
        Me.Combo4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo4.Size = New System.Drawing.Size(216, 22)
        Me.Combo4.TabIndex = 18
        '
        'Combo3
        '
        Me.Combo3.BackColor = System.Drawing.SystemColors.Window
        Me.Combo3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo3.Location = New System.Drawing.Point(408, 8)
        Me.Combo3.Name = "Combo3"
        Me.Combo3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo3.Size = New System.Drawing.Size(360, 22)
        Me.Combo3.TabIndex = 17
        Me.Combo3.Visible = False
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(8, 304)
        Me.Text2.MaxLength = 20
        Me.Text2.Name = "Text2"
        Me.Text2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(177, 20)
        Me.Text2.TabIndex = 15
        '
        'Command6
        '
        Me.Command6.BackColor = System.Drawing.SystemColors.Control
        Me.Command6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command6.Font = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command6.Image = CType(resources.GetObject("Command6.Image"), System.Drawing.Image)
        Me.Command6.Location = New System.Drawing.Point(224, 336)
        Me.Command6.Name = "Command6"
        Me.Command6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command6.Size = New System.Drawing.Size(305, 69)
        Me.Command6.TabIndex = 14
        Me.Command6.Text = "Create Project"
        Me.Command6.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command6.UseVisualStyleBackColor = False
        '
        'Check1
        '
        Me.Check1.BackColor = System.Drawing.SystemColors.Control
        Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Check1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check1.Location = New System.Drawing.Point(408, 208)
        Me.Check1.Name = "Check1"
        Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check1.Size = New System.Drawing.Size(81, 25)
        Me.Check1.TabIndex = 13
        Me.Check1.Text = "Encrypt"
        Me.Check1.UseVisualStyleBackColor = False
        '
        'Command5
        '
        Me.Command5.BackColor = System.Drawing.SystemColors.Control
        Me.Command5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command5.Image = CType(resources.GetObject("Command5.Image"), System.Drawing.Image)
        Me.Command5.Location = New System.Drawing.Point(712, 104)
        Me.Command5.Name = "Command5"
        Me.Command5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command5.Size = New System.Drawing.Size(57, 65)
        Me.Command5.TabIndex = 12
        Me.Command5.Text = "Save"
        Me.Command5.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command5.UseVisualStyleBackColor = False
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(528, 104)
        Me.Text1.MaxLength = 20
        Me.Text1.Name = "Text1"
        Me.Text1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(177, 20)
        Me.Text1.TabIndex = 11
        '
        'Combo2
        '
        Me.Combo2.BackColor = System.Drawing.SystemColors.Window
        Me.Combo2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo2.Location = New System.Drawing.Point(408, 176)
        Me.Combo2.Name = "Combo2"
        Me.Combo2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo2.Size = New System.Drawing.Size(361, 22)
        Me.Combo2.TabIndex = 7
        '
        'Combo1
        '
        Me.Combo1.BackColor = System.Drawing.SystemColors.Window
        Me.Combo1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo1.Location = New System.Drawing.Point(408, 72)
        Me.Combo1.Name = "Combo1"
        Me.Combo1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo1.Size = New System.Drawing.Size(361, 22)
        Me.Combo1.TabIndex = 6
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Image = CType(resources.GetObject("Command4.Image"), System.Drawing.Image)
        Me.Command4.Location = New System.Drawing.Point(352, 192)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(49, 45)
        Me.Command4.TabIndex = 5
        Me.Command4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command4.UseVisualStyleBackColor = False
        '
        'List1
        '
        Me.List1.BackColor = System.Drawing.SystemColors.Window
        Me.List1.Cursor = System.Windows.Forms.Cursors.Default
        Me.List1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.List1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.List1.ItemHeight = 14
        Me.List1.Location = New System.Drawing.Point(8, 40)
        Me.List1.Name = "List1"
        Me.List1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.List1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.List1.Size = New System.Drawing.Size(337, 228)
        Me.List1.TabIndex = 3
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Image = CType(resources.GetObject("Command1.Image"), System.Drawing.Image)
        Me.Command1.Location = New System.Drawing.Point(352, 48)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(49, 41)
        Me.Command1.TabIndex = 2
        Me.Command1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Image = CType(resources.GetObject("Command2.Image"), System.Drawing.Image)
        Me.Command2.Location = New System.Drawing.Point(352, 96)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(49, 41)
        Me.Command2.TabIndex = 1
        Me.Command2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command2.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Image = CType(resources.GetObject("Command3.Image"), System.Drawing.Image)
        Me.Command3.Location = New System.Drawing.Point(352, 144)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(49, 45)
        Me.Command3.TabIndex = 0
        Me.Command3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(8, 344)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(89, 25)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Save Project"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(224, 416)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(297, 17)
        Me.Label6.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(8, 272)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(209, 25)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Enter Project Password (Optional)"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(408, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(113, 25)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Enter Password"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(408, 144)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(129, 25)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Encrypt"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(408, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(129, 25)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "PASSWORD STATUS"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(129, 25)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Files"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(408, 256)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(136, 24)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "ENCRYPT STATUS"
        '
        'frmAdd
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(776, 486)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Command9)
        Me.Controls.Add(Me.Command8)
        Me.Controls.Add(Me.Command7)
        Me.Controls.Add(Me.Text3)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Combo5)
        Me.Controls.Add(Me.Combo4)
        Me.Controls.Add(Me.Combo3)
        Me.Controls.Add(Me.Command6)
        Me.Controls.Add(Me.Check1)
        Me.Controls.Add(Me.Command5)
        Me.Controls.Add(Me.Combo2)
        Me.Controls.Add(Me.Combo1)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.List1)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmAdd"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Project"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region

#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmAdd
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmAdd
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmAdd
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As frmAdd)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region
    Dim PH As ProjectHeader
    Dim FH() As FileHeader
    Dim si As Short

    Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged
        If Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
            VB6.SetItemString(Combo4, Combo2.SelectedIndex, "1")
        Else
            VB6.SetItemString(Combo4, Combo2.SelectedIndex, "0")
        End If
    End Sub

    Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
        Text1.Text = VB6.GetItemString(Combo3, Combo1.SelectedIndex)
    End Sub

    Private Sub Combo2_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo2.SelectedIndexChanged
        If VB6.GetItemString(Combo4, Combo2.SelectedIndex) = "1" Then
            Check1.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
    End Sub

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        ChangeOrder(List1, -1)
        ChangeOrder(Combo1, -1)
        ChangeOrder(Combo2, -1)
        ChangeOrder(Combo3, -1)
        ChangeOrder(Combo4, -1)
        ChangeOrder(Combo5, -1)
    End Sub
    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        ChangeOrder(List1, 1)
        ChangeOrder(Combo1, 1)
        ChangeOrder(Combo2, 1)
        ChangeOrder(Combo3, 1)
        ChangeOrder(Combo4, 1)
        ChangeOrder(Combo5, 1)
    End Sub

    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Dim i As Object
        Dim ofd As New OpenFileDialog
        ofd.InitialDirectory = "C:\"
        ofd.Filter = "All Files (*.*)|*.*"
        ofd.ShowDialog()
        If ofd.FileName <> "" Then
            For i = 0 To List1.Items.Count - 1
                If VB6.GetItemString(List1, i) = ofd.FileName Then
                    MsgBox("Already Added", MsgBoxStyle.Information)
                    Exit Sub
                End If
            Next
            List1.Items.Add(ofd.FileName)
            Combo3.Items.Add("")
            Combo4.Items.Add("")
            Combo5.Items.Add("")
            RefillComboBoxes()
        End If
    End Sub

    Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
        Try
            Dim index As Integer
            List1.Items.RemoveAt(index)
            Combo1.Items.RemoveAt(index)
            Combo2.Items.RemoveAt(index)
        Catch ex As Exception
            MsgBox("Please Select any file in the ListBox to Delete")
        End Try
    End Sub

    Public Sub ChangeOrder(ByRef List1 As ListBox, ByVal mode As Short)
        On Error GoTo down
        Dim arr() As String
        Dim i As Short
        Dim x As String
        Dim tmp As String
        If si >= 0 Then
            ReDim arr(List1.Items.Count - 1)
            For i = 0 To List1.Items.Count - 1
                arr(i) = VB6.GetItemString(List1, i)
            Next
            tmp = arr(si)
            arr(si) = arr(si + mode)
            arr(si + mode) = tmp
            List1.Items(si) = arr(si)
            List1.Items(si + mode) = arr(si + mode)
            List1.SelectedIndex = si
            List1.Focus()
        End If
down:
    End Sub

    Public Sub ChangeOrder(ByRef List1 As ComboBox, ByVal mode As Short)
        On Error GoTo down
        Dim arr() As String
        Dim i As Short
        Dim x As String
        Dim tmp As String
        If si >= 0 Then
            ReDim arr(List1.Items.Count - 1)
            For i = 0 To List1.Items.Count - 1
                arr(i) = VB6.GetItemString(List1, i)
            Next
            tmp = arr(si)
            arr(si) = arr(si + mode)
            arr(si + mode) = tmp
            List1.Items(si) = arr(si)
            List1.Items(si + mode) = arr(si + mode)
            List1.SelectedIndex = si
            List1.Focus()
        End If
down:
    End Sub

    Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
        Dim result As Integer
        If Text1.Text = "" Then
            result = MsgBox("Are you sure you want to save the data without the password", MsgBoxStyle.OKCancel)
        Else
            result = MsgBox("Are you sure you want to save the data with password", MsgBoxStyle.OKCancel)
        End If
        If result = 1 Then
            If Text1.Text = "" Then
                VB6.SetItemString(Combo5, Combo1.SelectedIndex, "0")
                VB6.SetItemString(Combo3, Combo1.SelectedIndex, "")
                If Check1.Checked = True Then
                    MsgBox("Successfully Encrypting the file=" & Combo2.Text & " without password")
                Else
                    MsgBox("Do not Encrypting the file=" & Combo2.Text & " without password")
                End If
            Else
                VB6.SetItemString(Combo5, Combo1.SelectedIndex, "1")
                VB6.SetItemString(Combo3, Combo1.SelectedIndex, Text1.Text)
                If Check1.Checked = True Then
                    MsgBox("Successfully Encrypting the file=" & Combo2.Text & " with password")
                Else
                    MsgBox("Do not Encrypting the file=" & Combo2.Text & " with password")
                End If
            End If
        End If
    End Sub

    Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
        Dim revext As String
        Dim revfn As String
        If List1.Items.Count = 0 Then
            MsgBox("Please Select Files", MsgBoxStyle.Information)
            Command1.Focus()
            Exit Sub
        End If
        If Text3.Text = "" Then
            MsgBox("Please Select Project File Path")
            Command7.Focus()
            Exit Sub
        End If
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Label6.Text = "Processing..."
        PH.NumberOfFiles = List1.Items.Count
        If Text2.Text = "" Then
            PH.IsProjectPasswordFound = "0"
            PH.Password = ""
        Else
            PH.IsProjectPasswordFound = "1"
            PH.Password = Text2.Text
        End If
        Dim PWFound As Boolean
        Dim i As Short
        For i = 0 To Combo5.Items.Count - 1
            If VB6.GetItemString(Combo5, i) = "1" Then
                PWFound = True
                Exit For
            End If
        Next
        If Not PWFound Then
            PH.IsNoFilePasswordProtected = "1"
        Else
            PH.IsNoFilePasswordProtected = "0"
        End If
        Dim fso As New Scripting.FileSystemObject
        ReDim FH(List1.Items.Count - 1)
        Dim fn As String
        Dim fil As Scripting.File
        For i = 0 To List1.Items.Count - 1
            fn = VB6.GetItemString(List1, i)
            fil = fso.GetFile(fn)
            FH(i).FileSize = fil.Size
            revfn = StrReverse(fn)
            revext = Mid(revfn, 1, InStr(1, revfn, ".") - 1)
            FH(i).FileExtension = StrReverse(revext)
            FH(i).FilePath = fil.Path
            If VB6.GetItemString(Combo5, i) = "1" Then
                FH(i).IsPasswordProtected = "1"
                FH(i).Password = VB6.GetItemString(Combo3, i)
            Else
                FH(i).IsPasswordProtected = "0"
                FH(i).Password = ""
            End If

            If VB6.GetItemString(Combo4, i) = "1" Then
                FH(i).IsEncrypted = "1"
            Else
                FH(i).IsEncrypted = "0"
            End If
        Next
        Dim Offset As Double
        If fso.FileExists(Text3.Text) Then
            fso.DeleteFile(Text3.Text, True)
        End If
        FileOpen(1, Text3.Text, OpenMode.Binary)
        FilePut(1, PH)
        For i = 0 To List1.Items.Count - 1
            FH(i).FileOffSet = 26 + (List1.Items.Count * 300) + Offset
            Offset = Offset + FH(i).FileSize
            FilePut(1, FH(i))
        Next
        Dim ts As Scripting.TextStream
        Dim c As String
        For i = 0 To List1.Items.Count - 1
            Label6.Text = "Saving " & VB6.GetItemString(List1, i)
            If FH(i).IsEncrypted = "1" Then
                EncFile("abcdefgh", VB6.GetItemString(List1, i), VB6.GetItemString(List1, i) & ".csd")
                ts = fso.OpenTextFile(VB6.GetItemString(List1, i) & ".csd", Scripting.IOMode.ForReading, False)
                fil = fso.GetFile(VB6.GetItemString(List1, i) & ".csd")
                FH(i).FileExtension = ".csd"
                FH(i).FileSize = fil.Size
            Else
                ts = fso.OpenTextFile(VB6.GetItemString(List1, i), Scripting.IOMode.ForReading, False)
                fil = fso.GetFile(VB6.GetItemString(List1, i))
                FH(i).FileSize = fil.Size
            End If
            Do While Not ts.AtEndOfStream
                c = ts.Read(1)
                FilePut(1, c)
            Loop
            ts.Close()
        Next
        Offset = 0
        For i = 0 To List1.Items.Count - 1
            FH(i).FileOffSet = 26 + List1.Items.Count * 300 + Offset
            Offset = Offset + FH(i).FileSize
            FilePut(1, FH(i), 27 + (i * 300))
        Next
        FileClose(1)
        For i = 0 To List1.Items.Count - 1
            System.IO.File.Delete(VB6.GetItemString(List1, i))
        Next
        Label6.Text = "Finished..."
        Me.Cursor = System.Windows.Forms.Cursors.Default
        MsgBox("Project Created", MsgBoxStyle.Information)
ErrInfo:
        If Err.Number <> 0 Then
            MsgBox(Err.Description)
        End If
    End Sub

    Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command7.Click
        Dim sfd As New SaveFileDialog
        sfd.FileName = "C:\Project1.fpp"
        sfd.Filter = "File Protector (*.FPP)|*.FPP"
        sfd.ShowDialog()
        Dim fso As New Scripting.FileSystemObject
        If sfd.FileName <> "" Then
            If UCase(VB.Right(sfd.FileName, 4)) <> ".FPP" Then
                Text3.Text = sfd.FileName & ".FPP"
            Else
                Text3.Text = sfd.FileName
            End If
            fso = Nothing
        End If
    End Sub

    Private Sub Command8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command8.Click
        Dim k As Integer
        Dim revext As String
        Dim revfn As String
        Dim i As Long
        Dim fn As String
        For i = 0 To List1.Items.Count - 1
            If List1.GetSelected(i) = True Then
                fn = List1.Text
                Exit For
            End If
        Next
        Dim pw As String
        Dim j As Long
        Dim c As String
        Dim fso As New Scripting.FileSystemObject
        If fn <> "" Then
            If h(i).IsPasswordProtected = "1" Then
                Dim PAA As New PASSWORD
                PAA.ShowDialog()
                pw = GET_PASS()
                If pw <> Trim(h(i).Password) Then
                    MsgBox("Wrong Password")
                    Exit Sub
                End If
            End If
            revfn = StrReverse(fn)
            revext = Mid(revfn, 1, InStr(1, revfn, ".") - 1)
            Dim sfd1 As New SaveFileDialog
            sfd1.Filter = "Files |*." & StrReverse(revext) & ""
            sfd1.ShowDialog()
            If sfd1.FileName <> "" Then
                fn = Mid(fn, InStrRev(fn, "\") + 1)
                FileOpen(1, Text3.Text, OpenMode.Binary)
                For j = 1 To h(List1.SelectedIndex).FileOffSet
                    c = Space(1)
                    FileGet(1, c, j)
                Next j
                If fso.FileExists(sfd1.FileName) Then
                    fso.DeleteFile(sfd1.FileName, True)
                End If
                fso = Nothing
                FileOpen(2, sfd1.FileName, OpenMode.Binary)
                FileOpen(3, sfd1.FileName & "1", OpenMode.Binary)
                For k = j To j + h(List1.SelectedIndex).FileSize - 1
                    c = Space(1)
                    FileGet(1, c, k)
                    If h(List1.SelectedIndex).IsEncrypted = "1" Then
                        FilePut(3, c)
                    Else
                        FilePut(2, c, k - j + 1)
                    End If
                Next k
                FileClose(2)
                FileClose(3)
                If h(List1.SelectedIndex).IsEncrypted = "1" Then
                    DecFile("abcdefgh", sfd1.FileName & "1", sfd1.FileName)
                End If
                FileClose(1)
                MsgBox("File Extracted To " & sfd1.FileName, MsgBoxStyle.Information)
            End If
        End If
ErrInfo:
        If Err.Number <> 0 Then
            MsgBox(Err.Description)
            On Error GoTo down
            FileClose(1)
            FileClose(2)
        End If
down:
    End Sub

    Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click
        Dim i As Integer
        Dim k As Integer
        Dim fn As String
        Dim cnt As Integer
        On Error GoTo ErrInfo
        Dim fso As New Scripting.FileSystemObject
        Dim FolderName As String
        Dim pw As String
        If h(0).IsPasswordProtected = "1" Then
            Dim PAA As New PASSWORD
            PAA.ShowDialog()
            pw = GET_PASS()
            If pw <> Trim(h(0).Password) Then
                MsgBox("Wrong Password")
                Exit Sub
            End If
        End If
        FolderBrowserDialog1.ShowDialog()
        FolderName = FolderBrowserDialog1.SelectedPath
        Dim j As Long
        Dim c As String
        If fso.FolderExists(FolderName) Then
            FileOpen(1, Text3.Text, OpenMode.Binary)
            For cnt = 0 To List1.Items.Count - 1
                fn = VB6.GetItemString(List1, cnt)
                fn = Mid(fn, InStrRev(fn, "\") + 1)
                For j = 1 To h(cnt).FileOffSet
                    c = Space(1)
                    FileGet(1, c, j)
                Next j
                On Error GoTo ErrInfo
                FileOpen(2, FolderName & "\" & fn, OpenMode.Binary)
                FileOpen(3, FolderName & "\" & fn & "1", OpenMode.Binary)
                For k = j To j + h(cnt).FileSize - 1
                    c = Space(1)
                    FileGet(1, c, k)
                    If h(cnt).IsEncrypted = "1" Then
                        FilePut(3, c) ', k - j + 1)
                    Else
                        FilePut(2, c, k - j + 1)
                    End If
                Next k
                FileClose(2)
                FileClose(3)
                If h(cnt).IsEncrypted = "1" Then
                    DecFile("abcdefgh", FolderName & "\" & fn & "1", FolderName & "\" & fn)
                End If
                Label6.Text = "Extracting " & VB6.GetItemString(List1, i)
            Next cnt
            FileClose(1)
            Label6.Text = "Finished"
            MsgBox("File(s) Extracted To " & FolderName, MsgBoxStyle.Information)
        End If
ErrInfo:
        If Err.Number <> 0 Then
            MsgBox(Err.Description)
            On Error GoTo down
            FileClose(1)
            FileClose(2)
        End If
down:
    End Sub

    Private Sub frmAdd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim x As System.Windows.Forms.Control
        If Module1.FileOpenMode = "Extract" Then
            For Each x In Me.Controls
                Select Case TypeName(x)
                    Case "TextBox", "ComboBox", "CommandButton", "CheckBox"
                        x.Enabled = False
                End Select
            Next x
            Command8.Visible = True
            Command9.Visible = True
            Command8.Enabled = True
        Else
            Command8.Visible = False
            Command9.Visible = False
        End If
        Label6.Text = ""
    End Sub
    Public Sub extractall()
        Dim aa As Integer
        Dim pw As String
        Dim last As String = ""
        Dim prev As String = ""
        For aa = 0 To List1.Items.Count - 1
            If aa = 0 Then
                prev = Trim(h(aa).Password)
                Command9.Enabled = True
            Else
                last = Trim(h(aa).Password)
                If last = prev Then
                    Command9.Enabled = True
                Else
                    Command9.Enabled = False
                    Exit For
                End If
                prev = last
            End If
        Next
    End Sub

    Private Sub List1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles List1.SelectedIndexChanged
        If Module1.FileOpenMode <> "Extract" Then
            si = List1.SelectedIndex
            If si = 0 Then
                Command1.Enabled = False
            Else
                Command1.Enabled = True
            End If

            If si = List1.Items.Count - 1 Then
                Command2.Enabled = False
            Else
                Command2.Enabled = True
            End If
        End If
    End Sub

    Sub RefillComboBoxes()
        Dim i As Short
        Combo1.Items.Clear()
        Combo2.Items.Clear()
        For i = 0 To List1.Items.Count - 1
            Combo1.Items.Add(VB6.GetItemString(List1, i))
            Combo2.Items.Add(VB6.GetItemString(List1, i))
        Next
    End Sub
End Class