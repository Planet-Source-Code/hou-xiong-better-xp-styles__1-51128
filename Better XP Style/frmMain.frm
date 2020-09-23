VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Better XP Styles"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   120
      Max             =   100
      TabIndex        =   13
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Frame grp1 
      Caption         =   "Group 1"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   4335
         TabIndex        =   1
         Top             =   240
         Width           =   4335
         Begin VB.Frame Frame1 
            Caption         =   "Child Group"
            Height          =   1335
            Left            =   1200
            TabIndex        =   11
            Top             =   600
            Width           =   3015
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   240
               TabIndex        =   14
               Text            =   "Combo List"
               Top             =   720
               Width           =   2535
            End
            Begin VB.Label Label1 
               Caption         =   "Nothing is in here! Or maybe not."
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Chess"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   1800
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Checker"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   1560
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Option 4"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option 3"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option 2"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option 1"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdMsgBox 
            Caption         =   "MsgBox"
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdInputBOx 
            Caption         =   "InputBox"
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox txtSample 
            Height          =   375
            Left            =   2400
            TabIndex        =   2
            Text            =   "empty"
            Top             =   0
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================
'
' Better XP Styles
'
' Hou Xiong
'
' In this new version, you don't have to carry an
' external XML file, which makes your app look
' unprofessional and you don't have to deal with all
' the resource files that you have to configure
' so accurately to make XP styles work.  This is
' very simple and straight to the point.  Enjoy
' and thanks for downloading.
'
'===================================================

Option Explicit

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200
Private Const XMLFile = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
                        "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
                        "<assemblyIdentity" & vbCrLf & _
                        "    version=""1.0.0.0""" & vbCrLf & _
                        "    processorArchitecture=""X86""" & vbCrLf & _
                        "    name=""WinXp""" & vbCrLf & _
                        "    type=""win32""" & vbCrLf & _
                        "/>" & vbCrLf & _
                        "<description>XP Style</description>" & vbCrLf & _
                        "<dependency>" & vbCrLf & _
                        "    <dependentAssembly>" & vbCrLf & _
                        "        <assemblyIdentity" & vbCrLf & _
                        "            type=""win32""" & vbCrLf & _
                        "            name=""Microsoft.Windows.Common-Controls""" & vbCrLf & _
                        "            version=""6.0.0.0""" & vbCrLf & _
                        "            processorArchitecture=""X86""" & vbCrLf & _
                        "            publicKeyToken=""6595b64144ccf1df""" & vbCrLf & _
                        "            language=""*""" & vbCrLf & _
                        "        />" & vbCrLf & _
                        "    </dependentAssembly>" & vbCrLf & _
                        "</dependency>" & vbCrLf & _
                        "</assembly>" & vbCrLf & vbCrLf

Private Const Compiled = False 'make this true before compilation

Private Sub InitCommControls()
   Dim iccex As tagInitCommonControlsEx
   iccex.lngSize = LenB(iccex)
   iccex.lngICC = ICC_USEREX_CLASSES
   InitCommonControlsEx iccex
End Sub

Private Sub Form_Initialize()
    Dim XMLFileName As String
    
    If Compiled Then
        XMLFileName = App.Path & "\" & App.EXEName & ".exe.manifest"
        'Does XML File exist
        If Dir(XMLFileName) = "" Then
            'XML does not exist, create new one
            Open XMLFileName For Binary As #1
                Put #1, , XMLFile
            Close #1
            'Run a new instance of this app so the new XML file takes effect
            Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
            'End this instance
            End
        Else
            'XML does exist
            'call function to make sure XP styles are used
            InitCommControls
            'Delete the XML since it's not necessary anymore
            Kill XMLFileName
        End If
    Else
        MsgBox "Must compile the app for XP Styles to work." & vbCrLf & "Remember to switch 'Compiled' to 'True' before compilation to turn this message off."
    End If
End Sub

'===================================================
'
' Misc
'
'===================================================

Private Sub cmdInputBOx_Click()
    txtSample = InputBox("Enter new value:", , txtSample)
End Sub

Private Sub cmdMsgBox_Click()
    MsgBox "MsgBox test!!!" & vbCrLf & vbCrLf & txtSample, vbOKCancel Or vbInformation
End Sub

Private Sub Form_Load()
    Dim Ctrl As Control
    For Each Ctrl In Controls
        Combo1.AddItem Ctrl.Name & ": " & TypeName(Ctrl)
    Next
End Sub
