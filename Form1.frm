VERSION 5.00
Object = "*\AProject1.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Load different menu"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   615
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   1200
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   8280
      Picture         =   "Form1.frx":0006
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":504A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":77FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CFEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D340
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D692
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D9E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DD36
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E088
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E3DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E72C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EA7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EDD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F122
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.XPsidemenu XPsidemenu1 
      Align           =   3  'Links ausrichten
      Height          =   7560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   13335
      ShowBorder      =   0   'False
      Speed           =   100
      Resizable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












Private Sub Command1_Click()
With Me.XPsidemenu1
.RemoveAll
Set .Pimagelist = Me.ImageList2
Set .hImageList = Me.ImageList1
    'set speed to 0 befor building menu speeds up process
    .Speed = 0
    .Addpanel "P1", "Caption 1", Opened, False, 1
    .Addpanel "P4", "Caption 4", Closed, True, , Picture1
   
    .AddHyper "H1", "P1", "Hyperlink 1", True, Hyperlink, 2, "This is tooltip 1"
    .AddHyper "H2", "P1", "Hyperlink 1", True, Hyperlink, 5, "This is tooltip 2"
    
    
    
    .AddHyper "H8", "P4", "Hyperlink 1", True, Label, 1, "This is tooltip 8"
    .AddHyper "H9", "P4", "Hyperlink 2", True, Label, 1, "This is tooltip 9"
    .AddHyper "H10", "P4", "Hyperlink 3", True, Label, 1, "This is tooltip 10"
    .AddHyper "H11", "P4", "Hyperlink 4", True, Label, 1, "This is tooltip 11"
    'set speed to desired value higher means faster
    .Speed = 15
End With
End Sub

Private Sub Form_Load()
With Me.XPsidemenu1
Set .Pimagelist = Me.ImageList2
Set .hImageList = Me.ImageList1
    'set speed to 0 befor building menu speeds up process
    .Speed = 0
    .Addpanel "P1", "Caption 1", Opened, False, 1
    .Addpanel "P2", "Caption 2", Closed, False, 2
    .Addpanel "P3", "Caption 3", Opened, True, 3, , App.Path & "\sample.jpg"
    .Addpanel "P4", "Caption 4", Fixed, True, , Picture1
   
    .AddHyper "H1", "P1", "Hyperlink 1", True, Hyperlink, 2, "This is tooltip 1"
    .AddHyper "H2", "P1", "Hyperlink 1", True, Hyperlink, 5, "This is tooltip 2"
    
    .AddHyper "H3", "P2", "Hyperlink 1", True, Hyperlink, 4, "This is tooltip 3"
    .AddHyper "H4", "P2", "Hyperlink 2", True, Hyperlink, 3, "This is tooltip 4"
    .AddHyper "H5", "P2", "Hyperlink 3", True, Hyperlink, 6, "This is tooltip 5"
    
    .AddHyper "H6", "P3", "Hyperlink 1", True, Hyperlink, 7, "This is tooltip 6"
    .AddHyper "H7", "P3", "Hyperlink 2", True, Hyperlink, 8, "This is tooltip 7"
    
    .AddHyper "H8", "P4", "Hyperlink 1", True, Label, 1, "This is tooltip 8"
    .AddHyper "H9", "P4", "Hyperlink 2", True, Label, 1, "This is tooltip 9"
    .AddHyper "H10", "P4", "Hyperlink 3", True, Label, 1, "This is tooltip 10"
    .AddHyper "H11", "P4", "Hyperlink 4", True, Label, 1, "This is tooltip 11"
    'set speed to desired value higher means faster
    .Speed = 50
End With
End Sub


Private Sub XPsidemenu1_HyperClick(key As String)
Label1.Caption = "You have Clicked on Hyperlink " & key
End Sub




Private Sub XPsidemenu1_PictureClick(key As String)
Label1.Caption = "You have Clicked on Picture " & key
End Sub


