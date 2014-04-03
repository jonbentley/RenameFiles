Option Strict Off
Option Explicit On 

Imports System.IO
Imports VB = Microsoft.VisualBasic

Public Class RenameFiles
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
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdThumb As System.Windows.Forms.Button
    Public WithEvents cmdMVI As System.Windows.Forms.Button
    Public WithEvents cmdRunFile As System.Windows.Forms.Button
    Public WithEvents txtLiteral3 As System.Windows.Forms.TextBox
    Public WithEvents cmdCreateFile As System.Windows.Forms.Button
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents chkSequenceNumber As System.Windows.Forms.CheckBox
    Public WithEvents txtSeqInc As System.Windows.Forms.TextBox
    Public WithEvents txtLiteral2 As System.Windows.Forms.TextBox
    Public WithEvents txtSeqPad As System.Windows.Forms.TextBox
    Public WithEvents txtSeqStartAt As System.Windows.Forms.TextBox
    Public WithEvents txtLiteral1 As System.Windows.Forms.TextBox
    Public WithEvents txtReplaceTo3 As System.Windows.Forms.TextBox
    Public WithEvents txtReplaceFrom3 As System.Windows.Forms.TextBox
    Public WithEvents txtReplaceTo2 As System.Windows.Forms.TextBox
    Public WithEvents txtReplaceFrom2 As System.Windows.Forms.TextBox
    Public WithEvents txtReplaceTo1 As System.Windows.Forms.TextBox
    Public WithEvents txtReplaceFrom1 As System.Windows.Forms.TextBox
    Public WithEvents chkAppendOriginalName As System.Windows.Forms.CheckBox
    Public WithEvents Line5 As System.Windows.Forms.Label
    Public WithEvents Line4 As System.Windows.Forms.Label
    Public WithEvents Line3 As System.Windows.Forms.Label
    Public WithEvents Line2 As System.Windows.Forms.Label
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents cboDirectory As System.Windows.Forms.ComboBox
    Public WithEvents Label18 As System.Windows.Forms.Label
    Public WithEvents txtSearchPattern As System.Windows.Forms.TextBox
    Public WithEvents Label19 As System.Windows.Forms.Label
    Public WithEvents txtDateTimeTakenFormat As System.Windows.Forms.TextBox
    Friend WithEvents btnTestPattern As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteRenameFiles As System.Windows.Forms.Button
    Friend WithEvents chkAddDatePictureTaken As System.Windows.Forms.CheckBox
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents txtDeletePattern As System.Windows.Forms.TextBox
    Public WithEvents cmdDeletePattern As System.Windows.Forms.Button
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents txtLiteral0 As System.Windows.Forms.TextBox
    Public WithEvents Label15 As System.Windows.Forms.Label
    Public WithEvents Label16 As System.Windows.Forms.Label
    Public WithEvents Label17 As System.Windows.Forms.Label
    Public WithEvents Label20 As System.Windows.Forms.Label
    Public WithEvents Label21 As System.Windows.Forms.Label
    Public WithEvents txtAdjustYears As System.Windows.Forms.TextBox
    Public WithEvents txtAdjustMonths As System.Windows.Forms.TextBox
    Public WithEvents txtAdjustDays As System.Windows.Forms.TextBox
    Public WithEvents Label22 As System.Windows.Forms.Label
    Public WithEvents txtAdjustSeconds As System.Windows.Forms.TextBox
    Public WithEvents Label23 As System.Windows.Forms.Label
    Public WithEvents txtAdjustMinutes As System.Windows.Forms.TextBox
    Public WithEvents txtAdjustHours As System.Windows.Forms.TextBox
    Public WithEvents Label24 As System.Windows.Forms.Label
    Public WithEvents Label25 As System.Windows.Forms.Label
    Public WithEvents Label26 As System.Windows.Forms.Label
    Public WithEvents txtNewExtension As System.Windows.Forms.TextBox
    Friend WithEvents chkExtensionAppendToOld As System.Windows.Forms.CheckBox
    Friend WithEvents cmdReformatFileName As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdThumb = New System.Windows.Forms.Button
        Me.cmdMVI = New System.Windows.Forms.Button
        Me.cmdRunFile = New System.Windows.Forms.Button
        Me.txtLiteral3 = New System.Windows.Forms.TextBox
        Me.cmdCreateFile = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.chkSequenceNumber = New System.Windows.Forms.CheckBox
        Me.txtSeqInc = New System.Windows.Forms.TextBox
        Me.txtLiteral2 = New System.Windows.Forms.TextBox
        Me.txtSeqPad = New System.Windows.Forms.TextBox
        Me.txtSeqStartAt = New System.Windows.Forms.TextBox
        Me.txtLiteral1 = New System.Windows.Forms.TextBox
        Me.txtReplaceTo3 = New System.Windows.Forms.TextBox
        Me.txtReplaceFrom3 = New System.Windows.Forms.TextBox
        Me.txtReplaceTo2 = New System.Windows.Forms.TextBox
        Me.txtReplaceFrom2 = New System.Windows.Forms.TextBox
        Me.txtReplaceTo1 = New System.Windows.Forms.TextBox
        Me.txtReplaceFrom1 = New System.Windows.Forms.TextBox
        Me.chkAppendOriginalName = New System.Windows.Forms.CheckBox
        Me.Line5 = New System.Windows.Forms.Label
        Me.Line4 = New System.Windows.Forms.Label
        Me.Line3 = New System.Windows.Forms.Label
        Me.Line2 = New System.Windows.Forms.Label
        Me.Line1 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.label1 = New System.Windows.Forms.Label
        Me.cboDirectory = New System.Windows.Forms.ComboBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtSearchPattern = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtDateTimeTakenFormat = New System.Windows.Forms.TextBox
        Me.btnTestPattern = New System.Windows.Forms.Button
        Me.cmdDeleteRenameFiles = New System.Windows.Forms.Button
        Me.chkAddDatePictureTaken = New System.Windows.Forms.CheckBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtDeletePattern = New System.Windows.Forms.TextBox
        Me.cmdDeletePattern = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtLiteral0 = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.txtAdjustYears = New System.Windows.Forms.TextBox
        Me.txtAdjustMonths = New System.Windows.Forms.TextBox
        Me.txtAdjustDays = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtAdjustSeconds = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtAdjustMinutes = New System.Windows.Forms.TextBox
        Me.txtAdjustHours = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.txtNewExtension = New System.Windows.Forms.TextBox
        Me.chkExtensionAppendToOld = New System.Windows.Forms.CheckBox
        Me.cmdReformatFileName = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'cmdThumb
        '
        Me.cmdThumb.BackColor = System.Drawing.SystemColors.Control
        Me.cmdThumb.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdThumb.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdThumb.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdThumb.Location = New System.Drawing.Point(72, 704)
        Me.cmdThumb.Name = "cmdThumb"
        Me.cmdThumb.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdThumb.Size = New System.Drawing.Size(33, 17)
        Me.cmdThumb.TabIndex = 42
        Me.cmdThumb.Text = "THM"
        Me.ToolTip1.SetToolTip(Me.cmdThumb, "Replace .THM extensions")
        '
        'cmdMVI
        '
        Me.cmdMVI.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMVI.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMVI.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMVI.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMVI.Location = New System.Drawing.Point(24, 704)
        Me.cmdMVI.Name = "cmdMVI"
        Me.cmdMVI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMVI.Size = New System.Drawing.Size(33, 17)
        Me.cmdMVI.TabIndex = 41
        Me.cmdMVI.Text = "MVI"
        Me.ToolTip1.SetToolTip(Me.cmdMVI, "Replace _MVI and _IMG")
        '
        'cmdRunFile
        '
        Me.cmdRunFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRunFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRunFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRunFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRunFile.Location = New System.Drawing.Point(256, 696)
        Me.cmdRunFile.Name = "cmdRunFile"
        Me.cmdRunFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRunFile.Size = New System.Drawing.Size(129, 25)
        Me.cmdRunFile.TabIndex = 40
        Me.cmdRunFile.Text = "&Run Last Rename File"
        '
        'txtLiteral3
        '
        Me.txtLiteral3.AcceptsReturn = True
        Me.txtLiteral3.AutoSize = False
        Me.txtLiteral3.BackColor = System.Drawing.SystemColors.Window
        Me.txtLiteral3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLiteral3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLiteral3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLiteral3.Location = New System.Drawing.Point(376, 584)
        Me.txtLiteral3.MaxLength = 0
        Me.txtLiteral3.Name = "txtLiteral3"
        Me.txtLiteral3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLiteral3.Size = New System.Drawing.Size(129, 19)
        Me.txtLiteral3.TabIndex = 23
        Me.txtLiteral3.Text = " (J)"
        '
        'cmdCreateFile
        '
        Me.cmdCreateFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCreateFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCreateFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCreateFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCreateFile.Location = New System.Drawing.Point(128, 696)
        Me.cmdCreateFile.Name = "cmdCreateFile"
        Me.cmdCreateFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCreateFile.Size = New System.Drawing.Size(120, 25)
        Me.cmdCreateFile.TabIndex = 22
        Me.cmdCreateFile.Text = "&Create Rename File"
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(544, 696)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(88, 25)
        Me.cmdCancel.TabIndex = 21
        Me.cmdCancel.Text = "Cancel"
        '
        'chkSequenceNumber
        '
        Me.chkSequenceNumber.BackColor = System.Drawing.SystemColors.Control
        Me.chkSequenceNumber.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSequenceNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSequenceNumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSequenceNumber.Location = New System.Drawing.Point(16, 312)
        Me.chkSequenceNumber.Name = "chkSequenceNumber"
        Me.chkSequenceNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSequenceNumber.Size = New System.Drawing.Size(160, 17)
        Me.chkSequenceNumber.TabIndex = 20
        Me.chkSequenceNumber.Text = "Add a Sequence Number"
        '
        'txtSeqInc
        '
        Me.txtSeqInc.AcceptsReturn = True
        Me.txtSeqInc.AutoSize = False
        Me.txtSeqInc.BackColor = System.Drawing.SystemColors.Window
        Me.txtSeqInc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSeqInc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSeqInc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSeqInc.Location = New System.Drawing.Point(368, 352)
        Me.txtSeqInc.MaxLength = 0
        Me.txtSeqInc.Name = "txtSeqInc"
        Me.txtSeqInc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSeqInc.Size = New System.Drawing.Size(129, 19)
        Me.txtSeqInc.TabIndex = 16
        Me.txtSeqInc.Text = "10"
        '
        'txtLiteral2
        '
        Me.txtLiteral2.AcceptsReturn = True
        Me.txtLiteral2.AutoSize = False
        Me.txtLiteral2.BackColor = System.Drawing.SystemColors.Window
        Me.txtLiteral2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLiteral2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLiteral2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLiteral2.Location = New System.Drawing.Point(368, 416)
        Me.txtLiteral2.MaxLength = 0
        Me.txtLiteral2.Name = "txtLiteral2"
        Me.txtLiteral2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLiteral2.Size = New System.Drawing.Size(129, 19)
        Me.txtLiteral2.TabIndex = 18
        Me.txtLiteral2.Text = ""
        '
        'txtSeqPad
        '
        Me.txtSeqPad.AcceptsReturn = True
        Me.txtSeqPad.AutoSize = False
        Me.txtSeqPad.BackColor = System.Drawing.SystemColors.Window
        Me.txtSeqPad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSeqPad.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSeqPad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSeqPad.Location = New System.Drawing.Point(368, 376)
        Me.txtSeqPad.MaxLength = 0
        Me.txtSeqPad.Name = "txtSeqPad"
        Me.txtSeqPad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSeqPad.Size = New System.Drawing.Size(129, 19)
        Me.txtSeqPad.TabIndex = 17
        Me.txtSeqPad.Text = "4"
        '
        'txtSeqStartAt
        '
        Me.txtSeqStartAt.AcceptsReturn = True
        Me.txtSeqStartAt.AutoSize = False
        Me.txtSeqStartAt.BackColor = System.Drawing.SystemColors.Window
        Me.txtSeqStartAt.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSeqStartAt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSeqStartAt.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSeqStartAt.Location = New System.Drawing.Point(368, 328)
        Me.txtSeqStartAt.MaxLength = 0
        Me.txtSeqStartAt.Name = "txtSeqStartAt"
        Me.txtSeqStartAt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSeqStartAt.Size = New System.Drawing.Size(129, 19)
        Me.txtSeqStartAt.TabIndex = 15
        Me.txtSeqStartAt.Text = "10"
        '
        'txtLiteral1
        '
        Me.txtLiteral1.AcceptsReturn = True
        Me.txtLiteral1.AutoSize = False
        Me.txtLiteral1.BackColor = System.Drawing.SystemColors.Window
        Me.txtLiteral1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLiteral1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLiteral1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLiteral1.Location = New System.Drawing.Point(192, 280)
        Me.txtLiteral1.MaxLength = 0
        Me.txtLiteral1.Name = "txtLiteral1"
        Me.txtLiteral1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLiteral1.Size = New System.Drawing.Size(64, 19)
        Me.txtLiteral1.TabIndex = 14
        Me.txtLiteral1.Text = ""
        '
        'txtReplaceTo3
        '
        Me.txtReplaceTo3.AcceptsReturn = True
        Me.txtReplaceTo3.AutoSize = False
        Me.txtReplaceTo3.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplaceTo3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplaceTo3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReplaceTo3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplaceTo3.Location = New System.Drawing.Point(376, 552)
        Me.txtReplaceTo3.MaxLength = 0
        Me.txtReplaceTo3.Name = "txtReplaceTo3"
        Me.txtReplaceTo3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplaceTo3.Size = New System.Drawing.Size(169, 19)
        Me.txtReplaceTo3.TabIndex = 13
        Me.txtReplaceTo3.Text = ""
        '
        'txtReplaceFrom3
        '
        Me.txtReplaceFrom3.AcceptsReturn = True
        Me.txtReplaceFrom3.AutoSize = False
        Me.txtReplaceFrom3.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplaceFrom3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplaceFrom3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReplaceFrom3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplaceFrom3.Location = New System.Drawing.Point(184, 552)
        Me.txtReplaceFrom3.MaxLength = 0
        Me.txtReplaceFrom3.Name = "txtReplaceFrom3"
        Me.txtReplaceFrom3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplaceFrom3.Size = New System.Drawing.Size(161, 19)
        Me.txtReplaceFrom3.TabIndex = 12
        Me.txtReplaceFrom3.Text = ""
        '
        'txtReplaceTo2
        '
        Me.txtReplaceTo2.AcceptsReturn = True
        Me.txtReplaceTo2.AutoSize = False
        Me.txtReplaceTo2.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplaceTo2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplaceTo2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReplaceTo2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplaceTo2.Location = New System.Drawing.Point(376, 528)
        Me.txtReplaceTo2.MaxLength = 0
        Me.txtReplaceTo2.Name = "txtReplaceTo2"
        Me.txtReplaceTo2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplaceTo2.Size = New System.Drawing.Size(169, 19)
        Me.txtReplaceTo2.TabIndex = 11
        Me.txtReplaceTo2.Text = ""
        '
        'txtReplaceFrom2
        '
        Me.txtReplaceFrom2.AcceptsReturn = True
        Me.txtReplaceFrom2.AutoSize = False
        Me.txtReplaceFrom2.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplaceFrom2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplaceFrom2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReplaceFrom2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplaceFrom2.Location = New System.Drawing.Point(184, 528)
        Me.txtReplaceFrom2.MaxLength = 0
        Me.txtReplaceFrom2.Name = "txtReplaceFrom2"
        Me.txtReplaceFrom2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplaceFrom2.Size = New System.Drawing.Size(161, 19)
        Me.txtReplaceFrom2.TabIndex = 10
        Me.txtReplaceFrom2.Text = "img_"
        '
        'txtReplaceTo1
        '
        Me.txtReplaceTo1.AcceptsReturn = True
        Me.txtReplaceTo1.AutoSize = False
        Me.txtReplaceTo1.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplaceTo1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplaceTo1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReplaceTo1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplaceTo1.Location = New System.Drawing.Point(376, 504)
        Me.txtReplaceTo1.MaxLength = 0
        Me.txtReplaceTo1.Name = "txtReplaceTo1"
        Me.txtReplaceTo1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplaceTo1.Size = New System.Drawing.Size(169, 19)
        Me.txtReplaceTo1.TabIndex = 9
        Me.txtReplaceTo1.Text = ""
        '
        'txtReplaceFrom1
        '
        Me.txtReplaceFrom1.AcceptsReturn = True
        Me.txtReplaceFrom1.AutoSize = False
        Me.txtReplaceFrom1.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplaceFrom1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplaceFrom1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReplaceFrom1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplaceFrom1.Location = New System.Drawing.Point(184, 504)
        Me.txtReplaceFrom1.MaxLength = 0
        Me.txtReplaceFrom1.Name = "txtReplaceFrom1"
        Me.txtReplaceFrom1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplaceFrom1.Size = New System.Drawing.Size(161, 19)
        Me.txtReplaceFrom1.TabIndex = 8
        Me.txtReplaceFrom1.Text = "mvi_"
        '
        'chkAppendOriginalName
        '
        Me.chkAppendOriginalName.BackColor = System.Drawing.SystemColors.Control
        Me.chkAppendOriginalName.Checked = True
        Me.chkAppendOriginalName.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAppendOriginalName.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAppendOriginalName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAppendOriginalName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAppendOriginalName.Location = New System.Drawing.Point(24, 448)
        Me.chkAppendOriginalName.Name = "chkAppendOriginalName"
        Me.chkAppendOriginalName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAppendOriginalName.Size = New System.Drawing.Size(257, 17)
        Me.chkAppendOriginalName.TabIndex = 6
        Me.chkAppendOriginalName.Text = "Append with the Original Name"
        '
        'Line5
        '
        Me.Line5.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line5.Location = New System.Drawing.Point(0, 304)
        Me.Line5.Name = "Line5"
        Me.Line5.Size = New System.Drawing.Size(640, 1)
        Me.Line5.TabIndex = 43
        '
        'Line4
        '
        Me.Line4.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line4.Location = New System.Drawing.Point(0, 576)
        Me.Line4.Name = "Line4"
        Me.Line4.Size = New System.Drawing.Size(640, 1)
        Me.Line4.TabIndex = 44
        '
        'Line3
        '
        Me.Line3.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line3.Location = New System.Drawing.Point(0, 440)
        Me.Line3.Name = "Line3"
        Me.Line3.Size = New System.Drawing.Size(640, 1)
        Me.Line3.TabIndex = 45
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Line2.Location = New System.Drawing.Point(0, 408)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(640, 1)
        Me.Line2.TabIndex = 46
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(0, 120)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(640, 1)
        Me.Line1.TabIndex = 47
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(16, 56)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(57, 17)
        Me.Label12.TabIndex = 29
        Me.Label12.Text = "Directory"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(312, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(297, 49)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "This Utility will take all the files in the current directory and create a batch " & _
        """rename"" job to allow easy renaming of all the files.  The options below allow f" & _
        "or tayloring of the ""new"" name"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(32, 584)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(241, 17)
        Me.Label10.TabIndex = 24
        Me.Label10.Text = "Another Literal"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(16, 352)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(273, 17)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "                                                     Increment By (+ / -)"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(40, 472)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(505, 25)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "If Original name included then                     REPLACE                       " & _
        "                           WITH"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 416)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(241, 17)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Another Literal"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(32, 376)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(225, 25)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Pad to what length ? (eg ""4"" means ""10"" becomes ""0010"")"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(16, 328)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(233, 17)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "                                                     Start At"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 280)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(88, 17)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Add a literal ...."
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(257, 17)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Format of the New Name...."
        '
        'label1
        '
        Me.label1.BackColor = System.Drawing.SystemColors.Control
        Me.label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.label1.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label1.Location = New System.Drawing.Point(80, 8)
        Me.label1.Name = "label1"
        Me.label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.label1.Size = New System.Drawing.Size(201, 41)
        Me.label1.TabIndex = 0
        Me.label1.Text = "Rename Files"
        '
        'cboDirectory
        '
        Me.cboDirectory.Location = New System.Drawing.Point(96, 56)
        Me.cboDirectory.Name = "cboDirectory"
        Me.cboDirectory.Size = New System.Drawing.Size(488, 22)
        Me.cboDirectory.TabIndex = 48
        Me.cboDirectory.Text = "C:\Photos\Add Times\Keith"
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(16, 88)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(176, 24)
        Me.Label18.TabIndex = 49
        Me.Label18.Text = "Search pattern (e.g. IMG*.jpg)"
        '
        'txtSearchPattern
        '
        Me.txtSearchPattern.AcceptsReturn = True
        Me.txtSearchPattern.AutoSize = False
        Me.txtSearchPattern.BackColor = System.Drawing.SystemColors.Window
        Me.txtSearchPattern.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSearchPattern.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchPattern.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSearchPattern.Location = New System.Drawing.Point(216, 88)
        Me.txtSearchPattern.MaxLength = 0
        Me.txtSearchPattern.Name = "txtSearchPattern"
        Me.txtSearchPattern.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSearchPattern.Size = New System.Drawing.Size(120, 19)
        Me.txtSearchPattern.TabIndex = 50
        Me.txtSearchPattern.Text = "*.jpg"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Location = New System.Drawing.Point(56, 200)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(160, 17)
        Me.Label19.TabIndex = 51
        Me.Label19.Text = "Format of the date/time"
        '
        'txtDateTimeTakenFormat
        '
        Me.txtDateTimeTakenFormat.AcceptsReturn = True
        Me.txtDateTimeTakenFormat.AutoSize = False
        Me.txtDateTimeTakenFormat.BackColor = System.Drawing.SystemColors.Window
        Me.txtDateTimeTakenFormat.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDateTimeTakenFormat.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDateTimeTakenFormat.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDateTimeTakenFormat.Location = New System.Drawing.Point(192, 200)
        Me.txtDateTimeTakenFormat.MaxLength = 0
        Me.txtDateTimeTakenFormat.Name = "txtDateTimeTakenFormat"
        Me.txtDateTimeTakenFormat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDateTimeTakenFormat.Size = New System.Drawing.Size(240, 19)
        Me.txtDateTimeTakenFormat.TabIndex = 52
        Me.txtDateTimeTakenFormat.Text = "yyyyMMdd HH-mm-ss / [d ddd HH-mm] "
        '
        'btnTestPattern
        '
        Me.btnTestPattern.Location = New System.Drawing.Point(424, 88)
        Me.btnTestPattern.Name = "btnTestPattern"
        Me.btnTestPattern.Size = New System.Drawing.Size(136, 16)
        Me.btnTestPattern.TabIndex = 53
        Me.btnTestPattern.Text = "Test the Pattern"
        '
        'cmdDeleteRenameFiles
        '
        Me.cmdDeleteRenameFiles.Location = New System.Drawing.Point(400, 696)
        Me.cmdDeleteRenameFiles.Name = "cmdDeleteRenameFiles"
        Me.cmdDeleteRenameFiles.Size = New System.Drawing.Size(128, 24)
        Me.cmdDeleteRenameFiles.TabIndex = 54
        Me.cmdDeleteRenameFiles.Text = "&Delete Rename Files"
        '
        'chkAddDatePictureTaken
        '
        Me.chkAddDatePictureTaken.Location = New System.Drawing.Point(8, 184)
        Me.chkAddDatePictureTaken.Name = "chkAddDatePictureTaken"
        Me.chkAddDatePictureTaken.Size = New System.Drawing.Size(440, 16)
        Me.chkAddDatePictureTaken.TabIndex = 55
        Me.chkAddDatePictureTaken.Text = "For Images with a ""Date Picture Taken"" attribute, add the date in the specified f" & _
        "ormat"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Label11.Location = New System.Drawing.Point(1, 608)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(640, 1)
        Me.Label11.TabIndex = 56
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(32, 672)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(176, 17)
        Me.Label13.TabIndex = 57
        Me.Label13.Text = "Delete all files matching pattern..."
        '
        'txtDeletePattern
        '
        Me.txtDeletePattern.AcceptsReturn = True
        Me.txtDeletePattern.AutoSize = False
        Me.txtDeletePattern.BackColor = System.Drawing.SystemColors.Window
        Me.txtDeletePattern.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDeletePattern.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeletePattern.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDeletePattern.Location = New System.Drawing.Point(216, 672)
        Me.txtDeletePattern.MaxLength = 0
        Me.txtDeletePattern.Name = "txtDeletePattern"
        Me.txtDeletePattern.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDeletePattern.Size = New System.Drawing.Size(64, 19)
        Me.txtDeletePattern.TabIndex = 58
        Me.txtDeletePattern.Text = "*.thm"
        '
        'cmdDeletePattern
        '
        Me.cmdDeletePattern.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDeletePattern.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDeletePattern.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeletePattern.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDeletePattern.Location = New System.Drawing.Point(288, 672)
        Me.cmdDeletePattern.Name = "cmdDeletePattern"
        Me.cmdDeletePattern.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDeletePattern.Size = New System.Drawing.Size(80, 16)
        Me.cmdDeletePattern.TabIndex = 59
        Me.cmdDeletePattern.Text = "Delete"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(24, 152)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(88, 17)
        Me.Label14.TabIndex = 60
        Me.Label14.Text = "Add a literal ...."
        '
        'txtLiteral0
        '
        Me.txtLiteral0.AcceptsReturn = True
        Me.txtLiteral0.AutoSize = False
        Me.txtLiteral0.BackColor = System.Drawing.SystemColors.Window
        Me.txtLiteral0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLiteral0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLiteral0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLiteral0.Location = New System.Drawing.Point(192, 152)
        Me.txtLiteral0.MaxLength = 0
        Me.txtLiteral0.Name = "txtLiteral0"
        Me.txtLiteral0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLiteral0.Size = New System.Drawing.Size(64, 19)
        Me.txtLiteral0.TabIndex = 61
        Me.txtLiteral0.Text = ""
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Label15.Location = New System.Drawing.Point(1, 272)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(640, 1)
        Me.Label15.TabIndex = 62
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.Label16.Location = New System.Drawing.Point(1, 176)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(640, 1)
        Me.Label16.TabIndex = 63
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(112, 224)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(56, 24)
        Me.Label17.TabIndex = 64
        Me.Label17.Text = "Adjust the Date/Time"
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(224, 224)
        Me.Label20.Name = "Label20"
        Me.Label20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label20.Size = New System.Drawing.Size(40, 17)
        Me.Label20.TabIndex = 65
        Me.Label20.Text = "Years"
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.SystemColors.Control
        Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label21.Location = New System.Drawing.Point(326, 224)
        Me.Label21.Name = "Label21"
        Me.Label21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label21.Size = New System.Drawing.Size(40, 17)
        Me.Label21.TabIndex = 66
        Me.Label21.Text = "Months"
        '
        'txtAdjustYears
        '
        Me.txtAdjustYears.AcceptsReturn = True
        Me.txtAdjustYears.AutoSize = False
        Me.txtAdjustYears.BackColor = System.Drawing.SystemColors.Window
        Me.txtAdjustYears.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdjustYears.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdjustYears.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdjustYears.Location = New System.Drawing.Point(272, 224)
        Me.txtAdjustYears.MaxLength = 0
        Me.txtAdjustYears.Name = "txtAdjustYears"
        Me.txtAdjustYears.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdjustYears.Size = New System.Drawing.Size(48, 19)
        Me.txtAdjustYears.TabIndex = 67
        Me.txtAdjustYears.Text = "+0"
        '
        'txtAdjustMonths
        '
        Me.txtAdjustMonths.AcceptsReturn = True
        Me.txtAdjustMonths.AutoSize = False
        Me.txtAdjustMonths.BackColor = System.Drawing.SystemColors.Window
        Me.txtAdjustMonths.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdjustMonths.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdjustMonths.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdjustMonths.Location = New System.Drawing.Point(384, 224)
        Me.txtAdjustMonths.MaxLength = 0
        Me.txtAdjustMonths.Name = "txtAdjustMonths"
        Me.txtAdjustMonths.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdjustMonths.Size = New System.Drawing.Size(48, 19)
        Me.txtAdjustMonths.TabIndex = 68
        Me.txtAdjustMonths.Text = "+0"
        '
        'txtAdjustDays
        '
        Me.txtAdjustDays.AcceptsReturn = True
        Me.txtAdjustDays.AutoSize = False
        Me.txtAdjustDays.BackColor = System.Drawing.SystemColors.Window
        Me.txtAdjustDays.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdjustDays.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdjustDays.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdjustDays.Location = New System.Drawing.Point(504, 224)
        Me.txtAdjustDays.MaxLength = 0
        Me.txtAdjustDays.Name = "txtAdjustDays"
        Me.txtAdjustDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdjustDays.Size = New System.Drawing.Size(48, 19)
        Me.txtAdjustDays.TabIndex = 70
        Me.txtAdjustDays.Text = "+0"
        '
        'Label22
        '
        Me.Label22.BackColor = System.Drawing.SystemColors.Control
        Me.Label22.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label22.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label22.Location = New System.Drawing.Point(440, 224)
        Me.Label22.Name = "Label22"
        Me.Label22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label22.Size = New System.Drawing.Size(32, 17)
        Me.Label22.TabIndex = 69
        Me.Label22.Text = "Days"
        '
        'txtAdjustSeconds
        '
        Me.txtAdjustSeconds.AcceptsReturn = True
        Me.txtAdjustSeconds.AutoSize = False
        Me.txtAdjustSeconds.BackColor = System.Drawing.SystemColors.Window
        Me.txtAdjustSeconds.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdjustSeconds.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdjustSeconds.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdjustSeconds.Location = New System.Drawing.Point(504, 248)
        Me.txtAdjustSeconds.MaxLength = 0
        Me.txtAdjustSeconds.Name = "txtAdjustSeconds"
        Me.txtAdjustSeconds.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdjustSeconds.Size = New System.Drawing.Size(48, 19)
        Me.txtAdjustSeconds.TabIndex = 76
        Me.txtAdjustSeconds.Text = "+0"
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.SystemColors.Control
        Me.Label23.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label23.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label23.Location = New System.Drawing.Point(440, 248)
        Me.Label23.Name = "Label23"
        Me.Label23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label23.Size = New System.Drawing.Size(48, 17)
        Me.Label23.TabIndex = 75
        Me.Label23.Text = "Seconds"
        '
        'txtAdjustMinutes
        '
        Me.txtAdjustMinutes.AcceptsReturn = True
        Me.txtAdjustMinutes.AutoSize = False
        Me.txtAdjustMinutes.BackColor = System.Drawing.SystemColors.Window
        Me.txtAdjustMinutes.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdjustMinutes.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdjustMinutes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdjustMinutes.Location = New System.Drawing.Point(384, 248)
        Me.txtAdjustMinutes.MaxLength = 0
        Me.txtAdjustMinutes.Name = "txtAdjustMinutes"
        Me.txtAdjustMinutes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdjustMinutes.Size = New System.Drawing.Size(48, 19)
        Me.txtAdjustMinutes.TabIndex = 74
        Me.txtAdjustMinutes.Text = "+0"
        '
        'txtAdjustHours
        '
        Me.txtAdjustHours.AcceptsReturn = True
        Me.txtAdjustHours.AutoSize = False
        Me.txtAdjustHours.BackColor = System.Drawing.SystemColors.Window
        Me.txtAdjustHours.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdjustHours.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdjustHours.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdjustHours.Location = New System.Drawing.Point(272, 248)
        Me.txtAdjustHours.MaxLength = 0
        Me.txtAdjustHours.Name = "txtAdjustHours"
        Me.txtAdjustHours.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdjustHours.Size = New System.Drawing.Size(48, 19)
        Me.txtAdjustHours.TabIndex = 73
        Me.txtAdjustHours.Text = "+0"
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.SystemColors.Control
        Me.Label24.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label24.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label24.Location = New System.Drawing.Point(328, 248)
        Me.Label24.Name = "Label24"
        Me.Label24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label24.Size = New System.Drawing.Size(48, 17)
        Me.Label24.TabIndex = 72
        Me.Label24.Text = "Minutes"
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.SystemColors.Control
        Me.Label25.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label25.Location = New System.Drawing.Point(224, 248)
        Me.Label25.Name = "Label25"
        Me.Label25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label25.Size = New System.Drawing.Size(40, 17)
        Me.Label25.TabIndex = 71
        Me.Label25.Text = "Hours"
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.Control
        Me.Label26.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label26.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label26.Location = New System.Drawing.Point(32, 616)
        Me.Label26.Name = "Label26"
        Me.Label26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label26.Size = New System.Drawing.Size(144, 17)
        Me.Label26.TabIndex = 77
        Me.Label26.Text = "Change the extension to ..."
        '
        'txtNewExtension
        '
        Me.txtNewExtension.AcceptsReturn = True
        Me.txtNewExtension.AutoSize = False
        Me.txtNewExtension.BackColor = System.Drawing.SystemColors.Window
        Me.txtNewExtension.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNewExtension.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewExtension.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNewExtension.Location = New System.Drawing.Point(376, 616)
        Me.txtNewExtension.MaxLength = 0
        Me.txtNewExtension.Name = "txtNewExtension"
        Me.txtNewExtension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNewExtension.Size = New System.Drawing.Size(129, 19)
        Me.txtNewExtension.TabIndex = 78
        Me.txtNewExtension.Text = ""
        '
        'chkExtensionAppendToOld
        '
        Me.chkExtensionAppendToOld.Location = New System.Drawing.Point(208, 616)
        Me.chkExtensionAppendToOld.Name = "chkExtensionAppendToOld"
        Me.chkExtensionAppendToOld.Size = New System.Drawing.Size(144, 16)
        Me.chkExtensionAppendToOld.TabIndex = 79
        Me.chkExtensionAppendToOld.Text = "Append to existing Ext"
        '
        'cmdReformatFileName
        '
        Me.cmdReformatFileName.Location = New System.Drawing.Point(512, 664)
        Me.cmdReformatFileName.Name = "cmdReformatFileName"
        Me.cmdReformatFileName.Size = New System.Drawing.Size(120, 23)
        Me.cmdReformatFileName.TabIndex = 80
        Me.cmdReformatFileName.Text = "Reformat Filename"
        '
        'RenameFiles
        '
        Me.AcceptButton = Me.cmdCreateFile
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(642, 726)
        Me.Controls.Add(Me.cmdReformatFileName)
        Me.Controls.Add(Me.chkExtensionAppendToOld)
        Me.Controls.Add(Me.txtNewExtension)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.txtAdjustSeconds)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.txtAdjustMinutes)
        Me.Controls.Add(Me.txtAdjustHours)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.txtAdjustDays)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.txtAdjustMonths)
        Me.Controls.Add(Me.txtAdjustYears)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.txtLiteral0)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.cmdDeletePattern)
        Me.Controls.Add(Me.txtDeletePattern)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.chkAddDatePictureTaken)
        Me.Controls.Add(Me.cmdDeleteRenameFiles)
        Me.Controls.Add(Me.btnTestPattern)
        Me.Controls.Add(Me.txtDateTimeTakenFormat)
        Me.Controls.Add(Me.txtSearchPattern)
        Me.Controls.Add(Me.txtLiteral3)
        Me.Controls.Add(Me.txtSeqInc)
        Me.Controls.Add(Me.txtLiteral2)
        Me.Controls.Add(Me.txtSeqPad)
        Me.Controls.Add(Me.txtSeqStartAt)
        Me.Controls.Add(Me.txtLiteral1)
        Me.Controls.Add(Me.txtReplaceTo3)
        Me.Controls.Add(Me.txtReplaceFrom3)
        Me.Controls.Add(Me.txtReplaceTo2)
        Me.Controls.Add(Me.txtReplaceFrom2)
        Me.Controls.Add(Me.txtReplaceTo1)
        Me.Controls.Add(Me.txtReplaceFrom1)
        Me.Controls.Add(Me.chkAppendOriginalName)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.cboDirectory)
        Me.Controls.Add(Me.cmdThumb)
        Me.Controls.Add(Me.cmdMVI)
        Me.Controls.Add(Me.cmdRunFile)
        Me.Controls.Add(Me.cmdCreateFile)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.chkSequenceNumber)
        Me.Controls.Add(Me.Line5)
        Me.Controls.Add(Me.Line4)
        Me.Controls.Add(Me.Line3)
        Me.Controls.Add(Me.Line2)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "RenameFiles"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ResumeLayout(False)

    End Sub
#End Region

    Const gcstrRegistry_AppName As String = "RenameFiles"
    Private mstrRenameFile As String


    Private Sub RenameFiles_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Populate fields
        cmdRunFile.Enabled = False
        mstrRenameFile = ""
        InitializeFields()
        SetUpFields_MVI_IMG()

    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdCreateFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCreateFile.Click
        Try
            Dim colFileNames_Old As New Collection
            Dim colFileNames_New As New Collection

            ' Must be a directory
            If System.IO.Directory.Exists(cboDirectory.Text) = False Then
                MsgBox("The ""Directory"" box must contain a valid directory")
                cboDirectory.Focus()
                Exit Sub
            End If

            A_GetListOfFilesInDirectory(colFileNames_Old)
            ' to manually reformat the file names then set a breakpoint here
            If 1 = 1 Then
                B_GenerateListOfNewFileNames(colFileNames_Old, colFileNames_New)
            Else
                B2_GenerateListOfNewFileNames_ManualOnly(colFileNames_Old, colFileNames_New)
            End If
            C_CreateBatchFile(colFileNames_Old, colFileNames_New)
        Catch ex As Exception
            FileOperations.String_To_Tempfile(ex.ToString, "RenameFilesError", "txt", True)
        End Try

    End Sub

    Private Function A_GetListOfFilesInDirectory(ByRef rcolFileNames_Old As Collection) As Integer
        ' get a list of all the files in the passed directory
        Dim strFile As String

        rcolFileNames_Old = New Collection

        Dim strFileList As String() = GetTheFiles(txtSearchPattern.Text)

        For Each strFilePathName As String In strFileList
            ' ignore hidden files
            If (GetAttr(strFilePathName) And FileAttribute.Hidden) = 0 Then
                ' Check out selection criteria
                ' Get the file without the path
                strFile = System.IO.Path.GetFileName(strFilePathName)
                If strFile.StartsWith("RenameFiles") = False _
                Or strFile.EndsWith(".bat") = False Then
                    rcolFileNames_Old.Add(strFile)
                End If
            End If
        Next

    End Function

    Private Function GetTheFiles(ByVal rstrPattern As String) As String()
        Dim strPattern As String = rstrPattern
        If strPattern = "" Then strPattern = "*"

        If System.IO.Directory.Exists(cboDirectory.Text) = False Then
            MsgBox("Directory does not exist : " & cboDirectory.Text)
            Return Nothing
        End If
        Return System.IO.Directory.GetFiles(cboDirectory.Text, strPattern)

    End Function

    Private Function B_GenerateListOfNewFileNames(ByRef rcolFileNames_Old As Collection, _
                                                  ByRef rcolFileNames_New As Collection) As Object
        ' For each passed file name, generate it's new file name based on the values
        ' set on the form
        Dim ix As Integer
        Dim strExtension As String = ""
        Dim strName As String = ""
        Dim strOldName As String = ""
        Dim strNewName As String = ""
        Dim intSeq As Short
        Dim strSeq As String
        rcolFileNames_New = New Collection

        '--------------------------------------------------------------
        ' if sequence number required - set up the start value
        If chkSequenceNumber.CheckState Then
            intSeq = CShort(txtSeqStartAt.Text)
        End If

        ' Loop round all the files
        For ix = 1 To rcolFileNames_Old.Count()
            strNewName = ""
            strExtension = GetTextString(1, rcolFileNames_Old.Item(ix), ".", basTextStringFuncs.GETSTRING.TextAfterLastFindString)
            strOldName = GetTextString(1, rcolFileNames_Old.Item(ix), ".", basTextStringFuncs.GETSTRING.TextBeforeLastFindString)

            '--------------------------------------------------------------
            ' Off we go with the 1st literal
            strNewName &= txtLiteral0.Text

            '--------------------------------------------------------------
            ' Set Date Time the picture was taken
            If chkAddDatePictureTaken.CheckState Then
                If txtDateTimeTakenFormat.Text <> "" Then
                    Dim dtDateTimeTaken As Date = ExifWrapper_DateTimeTaken(cboDirectory.Text & "\" & rcolFileNames_Old.Item(ix))
                    ' Adjust the date/time
                    ' really should have some proper validation - ho hum.
                    Try
                        dtDateTimeTaken = dtDateTimeTaken.AddYears(txtAdjustYears.Text)
                        dtDateTimeTaken = dtDateTimeTaken.AddMonths(txtAdjustMonths.Text)
                        dtDateTimeTaken = dtDateTimeTaken.AddDays(txtAdjustDays.Text)
                        dtDateTimeTaken = dtDateTimeTaken.AddHours(txtAdjustHours.Text)
                        dtDateTimeTaken = dtDateTimeTaken.AddMinutes(txtAdjustMinutes.Text)
                        dtDateTimeTaken = dtDateTimeTaken.AddSeconds(txtAdjustSeconds.Text)
                    Catch ex As Exception
                        MsgBox("You've probably put a non numeric value is teh dat adjusting fields" & vbCrLf & vbCrLf _
                        & ex.ToString)

                    End Try

                    strNewName &= Format(dtDateTimeTaken, txtDateTimeTakenFormat.Text)
                End If
            End If
            '--------------------------------------------------------------
            ' Now another literal
            strNewName &= txtLiteral1.Text
            '--------------------------------------------------------------
            ' if sequence number required - append it
            If chkSequenceNumber.CheckState Then
                ' increment the sequence
                If ix > 1 Then intSeq = intSeq + CShort(txtSeqInc.Text)

                strSeq = CStr(intSeq)
                ' pad with spaces if required
                If txtSeqPad.Text <> "" Then
                    strSeq = VB.Right("0000000000" & strSeq, CShort(txtSeqPad.Text))
                End If

                ' append the new sequence no
                strNewName = strNewName & strSeq
            End If

            '--------------------------------------------------------------
            ' Now another literal
            strNewName = strNewName & txtLiteral2.Text

            '--------------------------------------------------------------
            ' append the original name
            If chkAppendOriginalName.CheckState Then
                strName = strOldName
                ' If there's a "replace from" specified then replace the FROM string with the TO string
                If txtReplaceFrom1.Text <> "" Then
                    ReplaceTextString(strName, txtReplaceFrom1.Text, txtReplaceTo1.Text)
                End If
                If txtReplaceFrom2.Text <> "" Then
                    ReplaceTextString(strName, txtReplaceFrom2.Text, txtReplaceTo2.Text)
                End If
                If txtReplaceFrom3.Text <> "" Then
                    ReplaceTextString(strName, txtReplaceFrom3.Text, txtReplaceTo3.Text)
                End If

                ''''''--------------------------------------------------------------
                '''''' Mickey Hack
                '''''' Change 1st 5 chars - it's a time, add 1 hour, 3 minutes
                '''''Dim strTime As String
                '''''Dim strTimeNew As String
                '''''Dim dteTime As Date

                '''''strTime = strOldName.Substring(0, 5)
                '''''dteTime = CDate(strTime.Replace("-", ":"))

                '''''strTimeNew = Format(dteTime.AddHours(1).AddMinutes(3), "HH-mm")

                '''''strName = strName.Replace(strTime, strTimeNew)
                '--------------------------------------------------------------

                strNewName = strNewName & strName
            End If

            '--------------------------------------------------------------
            ' Now another literal
            strNewName = strNewName & txtLiteral3.Text

            '--------------------------------------------------------------
            ' add the old extension back on or a new one if specified - or both
            If txtNewExtension.Text <> "" Then
                If chkExtensionAppendToOld.CheckState = CheckState.Checked Then
                    strNewName = strNewName & "." & strExtension & "." & txtNewExtension.Text
                Else
                    strNewName = strNewName & "." & txtNewExtension.Text
                End If
            Else
                strNewName = strNewName & "." & strExtension
            End If


            rcolFileNames_New.Add(strNewName)

        Next ix
    End Function

    Private Function B2_GenerateListOfNewFileNames_ManualOnly(ByRef rcolFileNames_Old As Collection, _
                                                             ByRef rcolFileNames_New As Collection) As Object
        ' For each passed file name, generate it's new file name based on the values
        ' set on the form
        Dim ix As Integer
        Dim strExtension As String = ""
        Dim strName As String = ""
        Dim strOldName As String = ""
        Dim strNewName As String = ""
        Dim intSeq As Short
        Dim strSeq As String
        rcolFileNames_New = New Collection


        ' Loop round all the files
        For ix = 1 To rcolFileNames_Old.Count()
            strNewName = ""
            strExtension = GetTextString(1, rcolFileNames_Old.Item(ix), ".", basTextStringFuncs.GETSTRING.TextAfterLastFindString)
            strOldName = GetTextString(1, rcolFileNames_Old.Item(ix), ".", basTextStringFuncs.GETSTRING.TextBeforeLastFindString)

            Dim str1 As String
            Dim str2 As String
            Dim str3 As String
            Dim str4 As String
            Dim str5 As String
            Dim strLastBit As String
            StringFunctions.ParseVar(strOldName, str1, "RB-2011", str2, "s-", str3, strLastBit)

            strNewName = str1 & str3 & strLastBit & "." & strExtension.ToLower

            rcolFileNames_New.Add(strNewName)

        Next ix
    End Function

    Private Function C_CreateBatchFile(ByRef rcolFileNames_Old As Collection, _
                                       ByRef rcolFileNames_New As Collection) As Integer
        Dim strRenameCommands As String
        Dim ix As Object

        mstrRenameFile = cboDirectory.Text & "\" & "RenameFiles_" & Format(Now, "yyMMdd_hh-mm-ss") & ".bat"

        For ix = 1 To rcolFileNames_Old.Count()
            strRenameCommands &= vbCrLf & "rename """ & rcolFileNames_Old.Item(ix) & """" & vbTab & vbTab & """" & rcolFileNames_New.Item(ix) & """"
        Next
        strRenameCommands &= vbCrLf & "pause"

        ' Create a Batch file from the string
        ' Do NOT "open" then file because "open" will "run" with a file type of .bat !
        FileOperations.String_To_File(strRenameCommands, mstrRenameFile, False)
        cmdRunFile.Enabled = True

        Dim strError As String
        If FileOperations.OpenFileinNotepad(mstrRenameFile, strError) > 0 Then MsgBox(strError)

    End Function


    Private Sub cmdMVI_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMVI.Click
        SetUpFields_MVI_IMG()
    End Sub

    Private Sub cmdRunFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRunFile.Click
        Dim rc As Integer
        Dim strError As String

        ' "open" the command file
        If FileOperations.OpenFile(mstrRenameFile, "", strError) > 0 Then
            MsgBox(strError)
        End If

    End Sub

    Private Sub cmdThumb_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdThumb.Click
        SetUpFields_Thumb()
    End Sub

    Private Sub txtSeqInc_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSeqInc.Leave
        If IsNumeric(txtSeqStartAt.Text) = False Then
            MsgBox("The Sequence Number Increment value needs to be numeric (positive or negative)")
            txtSeqStartAt.Focus()
        End If
    End Sub

    Private Sub txtSeqStartAt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSeqStartAt.Leave
        If IsNumeric(txtSeqStartAt.Text) = False Then
            MsgBox("The Start At Sequence Number needs to be numeric")
            txtSeqStartAt.Focus()
        End If
    End Sub

    Private Function InitializeFields() As Object
        txtLiteral1.Text = ""
        chkSequenceNumber.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtSeqStartAt.Text = "10"
        txtSeqInc.Text = "10"
        txtSeqPad.Text = "4"
        txtLiteral2.Text = ""
        chkAppendOriginalName.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtReplaceFrom1.Text = ""
        txtReplaceTo1.Text = ""
        txtReplaceFrom2.Text = ""
        txtReplaceTo2.Text = ""
        txtReplaceFrom3.Text = ""
        txtReplaceTo3.Text = ""
        txtLiteral3.Text = ""
    End Function


    Private Function SetUpFields_MVI_IMG() As Object
        InitializeFields()
        txtLiteral1.Text = ""
        chkSequenceNumber.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtLiteral2.Text = ""
        chkAppendOriginalName.CheckState = System.Windows.Forms.CheckState.Checked
        txtReplaceFrom1.Text = "mvi_"
        txtReplaceTo1.Text = ""
        txtReplaceFrom2.Text = "img_"
        txtReplaceTo2.Text = ""
        txtReplaceFrom3.Text = ""
        txtReplaceTo3.Text = ""
        txtLiteral3.Text = ""
    End Function

    Private Function SetUpFields_Thumb() As Object
        InitializeFields()
        txtLiteral1.Text = ""
        chkSequenceNumber.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtLiteral2.Text = ""
        chkAppendOriginalName.CheckState = System.Windows.Forms.CheckState.Checked
        txtReplaceFrom1.Text = "_MVI"
        txtReplaceTo1.Text = "_Thumb"
        txtReplaceFrom2.Text = ""
        txtReplaceTo2.Text = ""
        txtReplaceFrom3.Text = ""
        txtReplaceTo3.Text = ""
        txtLiteral3.Text = ""
    End Function

    Public Function ExifWrapper_DateTimeTaken(ByVal rstrFile As String) As Date

        Dim R As ExifReader
        Try

            ''-- Load image
            'Dim I As System.Drawing.Bitmap
            'Try
            '    I = New System.Drawing.Bitmap(rstrFile)
            'Catch ex As Exception
            '    MsgBox("Error: Can't load image, error: " & ex.Message)
            'End Try
            'If I.PropertyIdList.Length = 0 Then MsgBox("This image does not contain any EXIF data")

            R = New ExifReader(rstrFile)

            Return R.DateTimeOriginal

        Catch ex As Exception
            Throw New Exception("Error in ExifWrapper_DateTimeTaken" & vbCrLf & ex.ToString)
        Finally
            R = Nothing
            GC.Collect()
        End Try

    End Function

    Private Sub btnTestPattern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestPattern.Click
        Dim strFileList As String() = GetTheFiles(txtSearchPattern.Text)
        FileOperations.Array_To_Tempfile(strFileList, "", "txt", True)

    End Sub

    Private Sub cmdDeleteRenameFiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteRenameFiles.Click
        Dim strRenameFileList As String() = System.IO.Directory.GetFiles(cboDirectory.Text, "RenameFiles*.bat")

        For Each strFile As String In strRenameFileList
            System.IO.File.Delete(strFile)
        Next

    End Sub

    Private Sub cboDirectory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDirectory.SelectedIndexChanged

    End Sub

    Private Sub cmdDeletePattern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeletePattern.Click
        ' Delete all files that match the pattern
        Dim strFileList As String() = GetTheFiles(txtDeletePattern.Text)
        Dim strFiles As String
        FileOperations.Array_To_String(strFileList, strFiles)

        ' Prompt - are you sure?
        If MsgBox("Are you sure you want to delete the follwoing files..." & vbCrLf & strFiles, MsgBoxStyle.YesNo, "Delete these files?") = MsgBoxResult.Yes Then
            For Each strFileToDelete As String In strFileList
                System.IO.File.Delete(strFileToDelete)
            Next
        End If

    End Sub

    Private Sub cmdReformatFileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReformatFileName.Click
        MessageBox.Show("To reformat the existing file name then edit the source, put a breakpoint in section B and do it manually.  It is not worth tryng to create a generic solution")
    End Sub
End Class