VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A56D3A&
   Caption         =   "�F�����~"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   FillStyle       =   0  '���
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11505
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame frame1 
      Appearance      =   0  '����
      BackColor       =   &H00A56D3A&
      Caption         =   "�n �J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   -120
      TabIndex        =   40
      Top             =   5640
      Width           =   3015
      Begin VB.CommandButton Command7 
         Caption         =   "�n �J �t ��"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   42
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox pass 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  '�Ȥ�
         Left            =   720
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label messfail 
         BackColor       =   &H00A56D3A&
         Caption         =   "�K�X���~!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   840
         TabIndex        =   48
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H00A56D3A&
         Caption         =   "�� �J �K �X :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   43
         Top             =   720
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   10841402
      TabCaption(0)   =   "�s�W"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Line3"
      Tab(0).Control(5)=   "Line1"
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Line2"
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(12)=   "Line4"
      Tab(0).Control(13)=   "Label10"
      Tab(0).Control(14)=   "Line6"
      Tab(0).Control(15)=   "Line7"
      Tab(0).Control(16)=   "Label16"
      Tab(0).Control(17)=   "Line5"
      Tab(0).Control(18)=   "Label17"
      Tab(0).Control(19)=   "Label24"
      Tab(0).Control(20)=   "Label7"
      Tab(0).Control(21)=   "Label29"
      Tab(0).Control(22)=   "Adodc2"
      Tab(0).Control(23)=   "Adodc1"
      Tab(0).Control(24)=   "Command1"
      Tab(0).Control(25)=   "fdata"
      Tab(0).Control(26)=   "sn"
      Tab(0).Control(27)=   "pdate"
      Tab(0).Control(28)=   "cname"
      Tab(0).Control(29)=   "ctype"
      Tab(0).Control(30)=   "tel"
      Tab(0).Control(31)=   "address"
      Tab(0).Control(32)=   "comm"
      Tab(0).Control(33)=   "km"
      Tab(0).Control(34)=   "Option1(1)"
      Tab(0).Control(35)=   "Option1(2)"
      Tab(0).Control(36)=   "money"
      Tab(0).Control(37)=   "Command3"
      Tab(0).Control(38)=   "Text1"
      Tab(0).Control(39)=   "Check1"
      Tab(0).Control(40)=   "Command8"
      Tab(0).Control(41)=   "Command9"
      Tab(0).Control(42)=   "cmark"
      Tab(0).Control(43)=   "tel2"
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "�d��"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(4)=   "Label15"
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(6)=   "Command2"
      Tab(1).Control(7)=   "putdata"
      Tab(1).Control(8)=   "ser_address"
      Tab(1).Control(9)=   "ser_tel"
      Tab(1).Control(10)=   "ser_ctype"
      Tab(1).Control(11)=   "ser_name"
      Tab(1).Control(12)=   "Command4"
      Tab(1).Control(13)=   "ser_cmark"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "�t�κ��@"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "����d��"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label26"
      Tab(3).Control(1)=   "Label27"
      Tab(3).Control(2)=   "ser_sdate"
      Tab(3).Control(3)=   "Command13"
      Tab(3).Control(4)=   "ser_edate"
      Tab(3).Control(5)=   "re_ser"
      Tab(3).Control(6)=   "Command6"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "�D�e��"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label22"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label21"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label19"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Picture1"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.TextBox tel2 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71760
         TabIndex        =   75
         Top             =   2450
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "�M   ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71520
         TabIndex        =   74
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox re_ser 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   -74400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '�������b
         TabIndex        =   71
         Top             =   2280
         Width           =   9855
      End
      Begin VB.TextBox ser_edate 
         Appearance      =   0  '����
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73440
         MaxLength       =   10
         TabIndex        =   70
         Top             =   1635
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         Caption         =   "�d   ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71520
         TabIndex        =   69
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox ser_sdate 
         Appearance      =   0  '����
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73440
         MaxLength       =   10
         TabIndex        =   68
         Top             =   915
         Width           =   1335
      End
      Begin VB.TextBox ser_cmark 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -67200
         MaxLength       =   10
         TabIndex        =   66
         Top             =   960
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         DragMode        =   1  '�۰�
         Enabled         =   0   'False
         FillStyle       =   0  '���
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         Left            =   1080
         MousePointer    =   5  '���V�|��
         Picture         =   "Form1.frx":008C
         ScaleHeight     =   4297.521
         ScaleMode       =   0  '�ϥΪ̦ۭq
         ScaleWidth      =   4195.402
         TabIndex        =   62
         Top             =   1320
         Width           =   8820
      End
      Begin VB.TextBox cmark 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -66240
         MaxLength       =   10
         TabIndex        =   60
         Top             =   600
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "�t�Τu��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -72840
         TabIndex        =   53
         Top             =   5040
         Width           =   6855
         Begin VB.CommandButton Command5 
            Caption         =   "�� �w"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4200
            TabIndex        =   55
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton Command10 
            Caption         =   "�� �}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4200
            TabIndex        =   54
            Top             =   1440
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�ƥ��٭�u��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -72840
         TabIndex        =   49
         Top             =   1080
         Width           =   6855
         Begin VB.DriveListBox Drive1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   480
            TabIndex        =   59
            Top             =   1200
            Width           =   1815
         End
         Begin VB.DirListBox Dir1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1890
            Left            =   480
            TabIndex        =   58
            Top             =   1560
            Width           =   1815
         End
         Begin VB.FileListBox File2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   2280
            Pattern         =   "*.mdb"
            TabIndex        =   57
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command12 
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4200
            TabIndex        =   56
            Top             =   2520
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   960
            TabIndex        =   51
            Text            =   "D:\BACKUP"
            Top             =   600
            Width           =   5535
         End
         Begin VB.CommandButton Command11 
            Caption         =   "��  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4200
            TabIndex        =   50
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label23 
            Caption         =   "���|"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   52
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "�M   ��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67680
         TabIndex        =   45
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "�C   �L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69240
         TabIndex        =   44
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72840
         TabIndex        =   38
         Top             =   600
         Value           =   1  '�֨�
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '����
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "Form1.frx":BCC2
         Top             =   6240
         Width           =   6975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�M   ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66240
         TabIndex        =   35
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox ser_name 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73800
         TabIndex        =   29
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox ser_ctype 
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox ser_tel 
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox ser_address 
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1920
         Width           =   5055
      End
      Begin VB.TextBox putdata 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   -74640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '�������b
         TabIndex        =   25
         Top             =   2400
         Width           =   10335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�d�߰򥻸��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70800
         TabIndex        =   24
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�d   ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68040
         TabIndex        =   23
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox money 
         Alignment       =   1  '�a�k���
         Appearance      =   0  '����
         BorderStyle     =   0  '�S���ؽu
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0123456789"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66360
         MaxLength       =   9
         TabIndex        =   13
         Top             =   6360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�^��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -68880
         TabIndex        =   12
         Top             =   1250
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -68880
         TabIndex        =   11
         Top             =   1080
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox km 
         Alignment       =   1  '�a�k���
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69720
         MaxLength       =   7
         TabIndex        =   10
         Top             =   1150
         Width           =   735
      End
      Begin VB.TextBox comm 
         Appearance      =   0  '����
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -66720
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox address 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70200
         TabIndex        =   8
         Top             =   2450
         Width           =   1815
      End
      Begin VB.TextBox tel 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   7
         Top             =   2450
         Width           =   1215
      End
      Begin VB.TextBox ctype 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -73920
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1900
         Width           =   5535
      End
      Begin VB.TextBox cname 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73920
         TabIndex        =   5
         Top             =   1600
         Width           =   2895
      End
      Begin VB.TextBox pdate 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65280
         TabIndex        =   4
         Top             =   1150
         Width           =   855
      End
      Begin VB.TextBox sn 
         Appearance      =   0  '����
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1150
         Width           =   975
      End
      Begin VB.TextBox fdata 
         Appearance      =   0  '����
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  '�������b
         TabIndex        =   2
         Top             =   3000
         Width           =   10455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�x  �s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72360
         TabIndex        =   1
         Top             =   7440
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   -74880
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Lv\test.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Lv\test.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "������"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   -73680
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Lv\test.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Lv\test.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "������"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   77
         Top             =   2445
         Width           =   135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʹq��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   -72600
         TabIndex        =   76
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72840
         TabIndex        =   73
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   72
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "���P"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68040
         TabIndex        =   67
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "�F �� �� �~"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   615
         Left            =   1080
         TabIndex        =   65
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label21 
         Caption         =   "�x�n���ñd�Ϥj�W�F�� 422 ��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   4560
         TabIndex        =   64
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label22 
         Caption         =   "TEL : 06 - 2716993"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   4560
         TabIndex        =   63
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label24 
         Caption         =   "���P"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66840
         TabIndex        =   61
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Caption         =   "�Ȥ�ñ�p"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67800
         TabIndex        =   39
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Line Line5 
         X1              =   -66840
         X2              =   -66840
         Y1              =   1560
         Y2              =   2760
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Caption         =   "�b���`�B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67800
         TabIndex        =   37
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "�F �� �� �~ �� �� �� ��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74640
         TabIndex        =   34
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label Label14 
         Caption         =   "���D�W��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   33
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "�����˦�"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   32
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "�p���q��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   31
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "�p���a�}"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   30
         Top             =   1920
         Width           =   735
      End
      Begin VB.Line Line7 
         X1              =   -66480
         X2              =   -66480
         Y1              =   6240
         Y2              =   7320
      End
      Begin VB.Line Line6 
         X1              =   -67920
         X2              =   -64200
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label Label10 
         Caption         =   "���{"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70200
         TabIndex        =   22
         Top             =   1200
         Width           =   375
      End
      Begin VB.Line Line4 
         X1              =   -74880
         X2              =   -64200
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label Label9 
         Caption         =   "�F�����~"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ĳ���פ����`�N�ƶ�"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -68160
         TabIndex        =   20
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   -68280
         X2              =   -68280
         Y1              =   1560
         Y2              =   2760
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�p���q��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�����˦�"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���D�W��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   1650
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   -74880
         X2              =   -64200
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line3 
         X1              =   -74880
         X2              =   -64200
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label3 
         Caption         =   "�L����"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66120
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "�u�@�渹"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "�e  ��  ��"
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70560
         TabIndex        =   14
         Top             =   480
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  '���
         Height          =   5775
         Left            =   -74880
         Top             =   1560
         Width           =   10695
      End
   End
   Begin VB.Label logo 
      BackColor       =   &H00A56D3A&
      Caption         =   "Design by sLab - Ver.1.1 (20130118)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   47
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Label Label20 
      BackColor       =   &H00A56D3A&
      Caption         =   "�F �� �� �~"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   46
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim Rs As New ADODB.Recordset
Dim dbpath, sql, column_name, column_value As String
Dim a(100)
Dim b(100)
Dim c As String
Dim click_num As Integer
Dim mk_sl As String
Dim pwd As String
Dim exe_path As String




Private Sub Command1_Click()

If cname.Text <> "" And fdata.Text <> "" And tel.Text <> "" And ctype.Text <> "" And cmark.Text <> "" And money.Text <> "" And km.Text <> "" Then
  click_num = MsgBox("�T�w�s�W?", 1, "�w�����")
  If click_num = 1 Then
    For I = 0 To 99
      a(I) = ""
      b(I) = ""
    Next I
    a(1) = "�u�@�渹"
    a(2) = "�p���a�}"
    a(3) = "�p���q��"
    a(4) = "�����˦�"
    a(5) = "���D�W��"
    a(6) = "���{"
    a(7) = "���{���"
    a(8) = "�L����"
    a(9) = "��ĳ���פ����`�N�ƶ�"
    a(10) = "���e"
    a(11) = "�`�B"
    a(12) = "���P"
    b(1) = sn.Text
    b(2) = address.Text
    b(3) = tel.Text + " , " + tel2.Text
    b(4) = ctype.Text
    b(5) = cname.Text
    b(6) = km.Text
    b(7) = mk_sl
    b(8) = pdate.Text
    b(9) = comm.Text
    b(10) = fdata.Text
    b(11) = money.Text
    b(12) = cmark.Text
    c = "�F�����~"
    Call Query_sql
    Call initdata
  End If
Else
  click_num = MsgBox("���|����g", vbOKOnly, "")
End If



End Sub

Private Sub Query_sql()

  column_name = ""
  column_value = ""
    For I = 1 To 12
        If b(I) <> "" And a(I) <> "" Then
            If column_name = "" Then
              column_name = a(I)
              column_value = "'" & b(I) & "'"
            Else
              column_name = column_name & "," & a(I)
              column_value = column_value & ",'" & b(I) & "'"
            End If
        End If
    Next I
  
    sql = "insert into " & c & "(" & column_name & ") values (" & column_value & ")"
    cmd.CommandText = sql
    Set Rs = cmd.Execute
End Sub
Private Sub initdata()

  Dim nowdate As String
  Dim nowsn As String
  Dim tmp() As String
  nowdate = Date
  tmp = Split(nowdate, "/")
  If tmp(1) < 10 Then
  tmp(1) = "0" & tmp(1)
  End If
  If tmp(2) < 10 Then
    tmp(2) = "0" & tmp(2)
  End If
  
  pdate.Text = nowdate
  nowsn = tmp(0) & tmp(1) & tmp(2)
  
  mk_sl = "����"
  
  sql = "select �u�@�渹 from " & c & " where �L����='" & nowdate & "' order by �u�@�渹 desc"

  cmd.CommandText = sql
  Set Rs = cmd.Execute
  If Rs.EOF = False Then
    nowsn = Rs.Fields("�u�@�渹").Value + 1
  Else
    nowsn = nowsn & "01"
  End If
  sn.Text = nowsn
  
End Sub

Private Sub Command10_Click()
End
End Sub

Private Sub Command11_Click()
  Dim nowdate As String
  Dim nowsn As String
  Dim tmp() As String
  nowdate = Date
  tmp = Split(nowdate, "/")
  If tmp(1) < 10 Then
    tmp(1) = "0" & tmp(1)
  End If
  If tmp(2) < 10 Then
    tmp(2) = "0" & tmp(2)
  End If
  nowsn = tmp(0) & tmp(1) & tmp(2)

If Text2.Text <> "" Then
  conn.Close
  If Right(Text2.Text, 1) = "\" Then
    FileCopy "dbfile.mdb", Text2.Text & nowsn & ".mdb"
  Else
    FileCopy "dbfile.mdb", Text2.Text + "\" + nowsn + ".mdb"
  End If
  Call LoadDb
End If



End Sub

Private Sub Command12_Click()
Dim getpwd As String
If Right(Text2.Text, 4) = ".mdb" Or Right(Text2.Text, 4) = ".MDB" Then
  getpwd = InputBox("�п�J�K�X:", "�٭���")
  If getpwd = pwd Then
    conn.Close
    FileCopy Text2.Text, exe_path + "\dbfile.mdb"
    Call LoadDb
  Else
    MsgBox "�K�X���~"
  End If
Else
  MsgBox "�п���ɮ�"
End If

End Sub

Private Sub Command13_Click()
If ser_sdate.Text <> "" And ser_edate.Text <> "" Then

  Dim tmp() As String
  tmp = Split(ser_sdate.Text, "/")
  If tmp(1) < 10 Then
    tmp(1) = "0" & tmp(1)
  End If
  If tmp(2) < 10 Then
    tmp(2) = "0" & tmp(2)
  End If
  ssd = tmp(0) & tmp(1) & tmp(2)
  
  tmp = Split(ser_edate.Text, "/")
  If tmp(1) < 10 Then
    tmp(1) = "0" & tmp(1)
  End If
  If tmp(2) < 10 Then
    tmp(2) = "0" & tmp(2)
  End If
  sed = tmp(0) & tmp(1) & tmp(2)

  sql = "select * from " & c & " where �u�@�渹 > " + ssd + "00 and �u�@�渹 < " + sed + "99 order by �u�@�渹"
  re_ser.Text = ""

  cmd.CommandText = sql
  Set Rs = cmd.Execute
  
  While (Rs.EOF = False)
    
    re_ser.Text = re_ser.Text + "<" + Rs.Fields("�L����").Value + ">    ���P: " + Rs.Fields("���P").Value + "    �u�@�渹: " + Str(Rs.Fields("�u�@�渹").Value) + "    ���{: " + Rs.Fields("���{").Value + "(" + Rs.Fields("���{���").Value + ") " + vbCrLf + "���D�W��: " + Rs.Fields("���D�W��").Value + "    �p���q��: " + Rs.Fields("�p���q��").Value + vbCrLf + Rs.Fields("���e").Value + vbCrLf + "�b���`�B: " + Rs.Fields("�`�B").Value + vbCrLf + vbCrLf
    Rs.MoveNext
    
  Wend

End If

End Sub

Private Sub Command2_Click()

If ser_name.Text <> "" Or ser_cmark.Text <> "" Then

  If ser_name.Text <> "" Then
    tmpsql = " ���D�W�� like '%" + ser_name.Text + "%' "
  End If
  
  If ser_cmark.Text <> "" And tmpsql <> "" Then
    tmpsql = tmpsql + " or ���P = '" + ser_cmark.Text + "' "
  ElseIf ser_cmark.Text <> "" Then
    tmpsql = tmpsql + " ���P = '" + ser_cmark.Text + "' "
  End If
  
  sql = "select * from " & c & " where " + tmpsql + " order by �u�@�渹"

  cmd.CommandText = sql
  Set Rs = cmd.Execute
  
  Call cleardata
  
  While (Rs.EOF = False)
  
    If Rs.Fields("���D�W��").Value <> "" Then
      If ser_name.Text <> Rs.Fields("���D�W��").Value And ser_name.Text <> "" Then
        ser_name.Text = ser_name.Text + " , " + Rs.Fields("���D�W��").Value
      Else
        ser_name.Text = Rs.Fields("���D�W��").Value
      End If
    End If

    If Rs.Fields("�p���q��").Value <> "" Then
      If ser_tel.Text <> Rs.Fields("�p���q��").Value And ser_tel.Text <> "" Then
        ser_tel.Text = ser_tel.Text + " , " + Rs.Fields("�p���q��").Value
      Else
        ser_tel.Text = Rs.Fields("�p���q��").Value
      End If
    End If

    If Rs.Fields("�����˦�").Value <> "" Then
      If ser_ctype.Text <> Rs.Fields("�����˦�").Value And ser_ctype.Text <> "" Then
        ser_ctype.Text = ser_ctype.Text + " , " + Rs.Fields("�����˦�").Value
      Else
        ser_ctype.Text = Rs.Fields("�����˦�").Value
      End If
    End If
    
    If Rs.Fields("�p���a�}").Value <> "" Then
      If ser_address.Text <> Rs.Fields("�p���a�}").Value And ser_address.Text <> "" Then
        ser_address.Text = ser_address.Text + " , " + Rs.Fields("�p���a�}").Value
      Else
        ser_address.Text = Rs.Fields("�p���a�}").Value
      End If
    End If
    
    'putdata.Text = putdata.Text + "=" + Rs.Fields("�L����").Value + "=  ���P:" + Rs.Fields("���P").Value + "  �u�@�渹:" + Str(Rs.Fields("�u�@�渹").Value) + "    ���{:" + Rs.Fields("���{").Value + "(" + Rs.Fields("���{���").Value + ") " + vbCrLf + Rs.Fields("���e").Value + vbCrLf + vbCrLf
    putdata.Text = putdata.Text + "<" + Rs.Fields("�L����").Value + ">    ���P: " + Rs.Fields("���P").Value + "    �u�@�渹: " + Str(Rs.Fields("�u�@�渹").Value) + "    ���{: " + Rs.Fields("���{").Value + "(" + Rs.Fields("���{���").Value + ") " + vbCrLf + "���D�W��: " + Rs.Fields("���D�W��").Value + "    �p���q��: " + Rs.Fields("�p���q��").Value + vbCrLf + Rs.Fields("���e").Value + vbCrLf + "�b���`�B: " + Rs.Fields("�`�B").Value + vbCrLf + vbCrLf
    
    Rs.MoveNext
    
  Wend

End If



End Sub

Private Sub Command3_Click()
Dim tmpsql As String
If cname.Text <> "" Or cmark.Text <> "" Then

  If cname.Text <> "" Then
    tmpsql = " where ���D�W��='" + cname.Text + "' "
  End If
  
  If cmark.Text <> "" And tmpsql <> "" Then
    tmpsql = tmpsql + " and ���P='" + cmark.Text + "' "
  ElseIf cmark.Text <> "" Then
    tmpsql = " where ���P='" + cmark.Text + "' "
  End If
  
  sql = "select * from " & c & tmpsql & " order by �u�@�渹"
  
  cmd.CommandText = sql
  Set Rs = cmd.Execute


    While (Rs.EOF = False)
      If Rs.Fields("���D�W��").Value <> "" Then
        cname.Text = Rs.Fields("���D�W��").Value
      End If
      If Rs.Fields("�����˦�").Value <> "" Then
        ctype.Text = Rs.Fields("�����˦�").Value
      End If
      If Rs.Fields("�p���q��").Value <> "" Then
        tel.Text = Rs.Fields("�p���q��").Value
      End If
      If Rs.Fields("�p���a�}").Value <> "" Then
        address.Text = Rs.Fields("�p���a�}").Value
      End If
      If Rs.Fields("���{").Value <> "" Then
        km.Text = Rs.Fields("���{").Value
      End If
      If Rs.Fields("���P").Value <> "" Then
        cmark.Text = Rs.Fields("���P").Value
      End If
      If Rs.Fields("���{���").Value <> "" Then
        If Rs.Fields("���{���").Value = "����" Then
          Option1(1).Value = True
          Option1(2).Value = False
        ElseIf Rs.Fields("���{���").Value = "�^��" Then
          Option1(1).Value = False
          Option1(2).Value = True
        End If
      End If
      Rs.MoveNext
    Wend
    
    
End If
End Sub

Private Sub Command4_Click()
Call cleardata
End Sub
Private Sub cleardata()

  ser_name.Text = ""
  ser_tel.Text = ""
  ser_ctype.Text = ""
  ser_address.Text = ""
  putdata.Text = ""
  ser_cmark = ""
  
End Sub

Private Sub print_table()
    Load dr1
    
    Set dr1.DataSource = Rs
    dr1.Sections("Section2").Controls("Label1").Caption = Label1.Caption
    dr1.Sections("Section2").Controls("Label3").Caption = Text1.Text
    dr1.Sections("Section2").Controls("Label4").Caption = fdata.Text
    dr1.Sections("Section2").Controls("Label5").Caption = cname.Text
    dr1.Sections("Section2").Controls("Label6").Caption = ctype.Text
    dr1.Sections("Section2").Controls("Label7").Caption = tel.Text
    dr1.Sections("Section2").Controls("tel22").Caption = tel2.Text
    dr1.Sections("Section2").Controls("Label8").Caption = address.Text
    dr1.Sections("Section2").Controls("Label9").Caption = comm.Text
    dr1.Sections("Section2").Controls("Label10").Caption = pdate.Text
    dr1.Sections("Section2").Controls("Label11").Caption = sn.Text
    dr1.Sections("Section2").Controls("Label12").Caption = km.Text
    dr1.Sections("Section2").Controls("Label20").Caption = mk_sl
    dr1.Sections("Section2").Controls("Label21").Caption = money.Text
    dr1.Sections("Section2").Controls("Label26").Caption = cmark.Text
    
    If Check1.Value = 0 Then
      dr1.Sections("Section2").Controls("Label2").Caption = ""
    End If
    
    dr1.Show

End Sub

Private Sub Command5_Click()
SSTab1.Visible = False
frame1.Visible = True
Command7.Default = True
End Sub

Private Sub Command6_Click()
re_ser = ""
End Sub

Private Sub Command7_Click()

If pass.Text = pwd Then
pass.Text = ""
SSTab1.Visible = True
frame1.Visible = False
Command7.Default = False
messfail.Visible = False
Else
messfail.Visible = True
pass.Text = ""
End If


End Sub

Private Sub Command8_Click()
    Call print_table
End Sub

Private Sub Command9_Click()
  cname.Text = ""
  tel.Text = ""
  tel2.Text = ""
  ctype.Text = ""
  address.Text = ""
  km.Text = ""
  comm.Text = ""
  fdata.Text = ""
  money.Text = ""
  cmark.Text = ""
  Call initdata
End Sub

Private Sub Dir1_Change()
File2.Path = Dir1.Path
Text2.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub



Private Sub File2_Click()
Text2.Text = Dir1.Path & "\" & File2.FileName

End Sub

Private Sub Form_Load()
    ser_sdate.Text = Date
    ser_edate.Text = Date
    exe_path = File2.Path
    Text2.Text = Dir1.Path
    pwd = "2j/ rm4"
    c = "�F�����~"

    Call LoadDb
    Call initdata
End Sub

Private Sub LoadDb()

    dbpath = "Provider=Microsoft.Jet.OLEDB.4.0; Persist Security Info=False;"
    dbpath = dbpath & "Data Source=" & App.Path & "\dbfile.mdb"
    conn.Open dbpath
    Rs.CursorLocation = adUseClient
    Set cmd.ActiveConnection = conn
    
End Sub

Private Sub Form_Resize()
  frame1.Top = (Form1.Height - frame1.Height) / 2
  frame1.Left = (Form1.Width - frame1.Width) / 2
  SSTab1.Top = (Form1.Height - SSTab1.Height) / 2 - 250
  SSTab1.Left = (Form1.Width - SSTab1.Width) / 2 - 50
  logo.Top = (Form1.Height - logo.Height) - 500
  logo.Left = (Form1.Width - logo.Width) - 400
End Sub

Private Sub Option1_Click(Index As Integer)
  If Index = 1 Then
    mk_sl = "����"
  ElseIf Index = 2 Then
    mk_sl = "�^��"
  End If
End Sub

Private Sub pass_Change()
SSTab1.Visible = False
End Sub



