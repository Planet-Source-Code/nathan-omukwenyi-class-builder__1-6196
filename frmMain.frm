VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Object Builder"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPicture 
      Height          =   6135
      Left            =   0
      TabIndex        =   28
      Top             =   120
      Width           =   3135
      Begin VB.Label lblHelpText 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2895
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgBanner 
         BorderStyle     =   1  'Fixed Single
         Height          =   2985
         Left            =   120
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2865
      End
   End
   Begin VB.Frame frmNavigation 
      Height          =   855
      Left            =   3240
      TabIndex        =   6
      Top             =   5400
      Width           =   7215
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Previous"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   495
         Left            =   2000
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3880
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "Finish"
         Height          =   495
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmInterfaces 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Interfaces"
      Height          =   5295
      Left            =   3240
      TabIndex        =   19
      Top             =   120
      Width           =   7215
      Begin VB.CheckBox Check1 
         Caption         =   "Collection to be created"
         Height          =   375
         Left            =   4680
         TabIndex        =   65
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton cmdAddEvent 
         Caption         =   "Event"
         Height          =   615
         Left            =   5040
         TabIndex        =   23
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddMethod 
         Caption         =   "Method"
         Height          =   615
         Left            =   5040
         TabIndex        =   22
         Top             =   2460
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddProperty 
         Caption         =   "Property"
         Height          =   615
         Left            =   5040
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin MSComctlLib.TreeView tvwTables 
         Height          =   4455
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7858
         _Version        =   393217
         Indentation     =   706
         Style           =   7
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lvlBackGround 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Interface"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   4560
         TabIndex        =   24
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame frmTables 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tables"
      Height          =   5295
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdFields 
         Caption         =   "Fields"
         Height          =   285
         Left            =   6120
         TabIndex        =   37
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   2720
         Width           =   375
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   1960
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtPropertyName 
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   " "
         Top             =   4800
         Width           =   2415
      End
      Begin VB.ListBox lstSelectedFields 
         Height          =   3570
         ItemData        =   "frmMain.frx":1F202
         Left            =   3480
         List            =   "frmMain.frx":1F204
         TabIndex        =   13
         Top             =   840
         Width           =   3495
      End
      Begin VB.ListBox lstTables 
         Height          =   4155
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label txtClassName 
         Caption         =   "Class Name"
         Height          =   255
         Left            =   3480
         TabIndex        =   36
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Tables"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame frmFinished 
      Height          =   5295
      Left            =   3240
      TabIndex        =   30
      Top             =   120
      Width           =   7215
      Begin MSComCtl2.Animation animProgress 
         Height          =   975
         Left            =   240
         TabIndex        =   57
         Top             =   2880
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1720
         _Version        =   393216
         Center          =   -1  'True
         FullWidth       =   449
         FullHeight      =   65
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   4440
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblProgressText 
         Caption         =   "Label2"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   2040
         Width           =   6615
      End
      Begin VB.Label lblFinishText 
         Caption         =   $"frmMain.frx":1F206
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame frmDirectory 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Project Directory"
      Height          =   5295
      Left            =   3240
      TabIndex        =   25
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdChangeDir 
         Caption         =   "Change Default Directory"
         Height          =   375
         Left            =   4080
         TabIndex        =   27
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtDirectory 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   3855
      End
      Begin VB.Frame frmSelectDirectory 
         Caption         =   "Select Folder"
         Height          =   3495
         Left            =   240
         TabIndex        =   51
         Top             =   1560
         Width           =   6375
         Begin VB.CommandButton cmdFolderCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   4320
            TabIndex        =   56
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton cmdNewFolder 
            Caption         =   "New Folder"
            Height          =   375
            Left            =   4320
            TabIndex        =   55
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Select Folder"
            Height          =   375
            Left            =   4320
            TabIndex        =   54
            Top             =   360
            Width           =   1695
         End
         Begin VB.DriveListBox DrvDrives 
            Height          =   315
            Left            =   480
            TabIndex        =   53
            Top             =   3000
            Width           =   3135
         End
         Begin VB.DirListBox DirFolders 
            Height          =   2565
            Left            =   480
            TabIndex        =   52
            Top             =   360
            Width           =   3135
         End
      End
   End
   Begin VB.Frame frmSQLServer 
      Caption         =   "Choose SQL Server database"
      Height          =   6135
      Left            =   3240
      TabIndex        =   58
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdSQLCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3000
         TabIndex        =   64
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CommandButton cmdSQLApply 
         Caption         =   "&Apply"
         Height          =   375
         Left            =   5160
         TabIndex        =   63
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox txtIntialCatalog 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   60
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   59
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label lblcatalog 
         Caption         =   "Initial catalog"
         Height          =   255
         Left            =   600
         TabIndex        =   62
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblServer 
         Caption         =   "SQL Server Name"
         Height          =   255
         Left            =   600
         TabIndex        =   61
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame frmDatabase 
      BackColor       =   &H00C0C0C0&
      Caption         =   "The database"
      Height          =   5295
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin MSComDlg.CommonDialog cdlgDatabase 
         Left            =   3240
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtDatabaseName 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3120
         Width           =   4815
      End
      Begin VB.CommandButton cmdOPenDatabase 
         Caption         =   "Open Database"
         Height          =   405
         Left            =   5280
         TabIndex        =   4
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Frame frmChooseDBType 
         Caption         =   "Choose Database type"
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6975
         Begin VB.OptionButton optSQL 
            Caption         =   "SQL Server Database"
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton optAccess 
            Caption         =   "Access Database"
            Height          =   375
            Left            =   480
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Label lblEstablishingConnection 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   360
         TabIndex        =   35
         Top             =   4200
         Width           =   105
      End
   End
   Begin VB.Frame frmFields 
      Caption         =   "Fields in"
      Height          =   6135
      Left            =   3240
      TabIndex        =   38
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   5640
         TabIndex        =   41
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelFields 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   40
         Top             =   5520
         Width           =   1335
      End
      Begin MSComctlLib.ListView LvwFields 
         Height          =   4695
         Left            =   240
         TabIndex        =   39
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8281
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Field Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data Type"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame frmProperties 
      Caption         =   "Properties"
      Height          =   5295
      Left            =   3240
      TabIndex        =   42
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtArguments 
         Height          =   1095
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   3000
         Width           =   6255
      End
      Begin VB.CommandButton cmdPropApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   5520
         TabIndex        =   48
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdPropCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   47
         Top             =   4680
         Width           =   1215
      End
      Begin VB.ComboBox cmbDataType 
         Height          =   315
         ItemData        =   "frmMain.frx":1F2E1
         Left            =   480
         List            =   "frmMain.frx":1F315
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtNewPropertyName 
         Height          =   285
         Left            =   480
         TabIndex        =   44
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label lblArguments 
         Caption         =   "Arguments"
         Height          =   255
         Left            =   480
         TabIndex        =   50
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label lblDataType 
         Caption         =   "Data Type"
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label lblPropertyName 
         Caption         =   "Property name"
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intState As Integer
Private Const state1 = "Choose the database type and then select the database that you want to use to create your business objects"
Private Const state2 = "Select fields from the tables in the database.You can also put other names for the fields"
Private Const state3 = "You can add properties,methods and events to the classes.Click on the class(table) then click the interface button to add an interface"
Private Const state4 = "Choose the directory where your classes and project will be saved. You may also elect to use the default folder"
Private Const state5 = "Click the 'Finish' Button to start generating your classes"

Private Const ACCESSPROVIDER = "Microsoft.Jet.OLEDB.3.51"
Private Const SQLPROVIDER = "SQLOLEDB.1"

Private strConnectionString As String
Private CurrentTablename As String
Private CurrentTableIndex As Integer
Private CurrentKey  As String

Private myTables As Tables
Private strDirectory As String

Private Enum InterfaceAdded
    interProperty = 1
    interMethod = 2
    intEvent = 3
End Enum

Private enumInterfaceAdded As InterfaceAdded

Dim v As New Collection
Dim v1 As New Collection

Dim strServer As String
Dim strInitCatalog As String

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Check1_Click()
    Dim i As Long
    
    For i = 1 To myTables.Count
    
        If myTables(i).Key = CurrentKey Then
            myTables(i).HasCollection = Check1.Value
        End If
        
    Next
End Sub

Private Sub cmdAdd_Click()
    
    Dim i As Integer
    Dim bfound As Boolean
    
    
    If Len(lstTables.List(lstTables.ListIndex)) < 1 Then Exit Sub
    
    
    ClearCollections
    
    If v.Count > 0 Then
        For i = 1 To v.Count
            If lstTables.List(lstTables.ListIndex) = v.Item(i) Then
                bfound = True
                Exit For
            End If
        Next i
    
    End If

    If Not bfound Then
        If v1.Count > 0 Then
            For i = 1 To v1.Count
                If lstTables.List(lstTables.ListIndex) = v1.Item(i) Then
                
                    bfound = True
                    Exit For
                End If
            Next i
        Else
            bfound = False
        End If
    
       If Not bfound Then
                       
            If Len(lstTables.List(lstTables.ListIndex)) > 0 Then
                v.Add lstTables.List(lstTables.ListIndex)
                v1.Add lstTables.List(lstTables.ListIndex)
                lstSelectedFields.AddItem lstTables.List(lstTables.ListIndex)
                
                For i = 1 To myTables.Count
                    If lstTables.List(lstTables.ListIndex) = myTables(i).tableName Then
                        myTables(i).Isincluded = True
                        Exit For
                    End If
                Next
            End If
        End If
        
    End If
    
End Sub
Private Sub ClearCollections()
    
    Dim i As Integer
    
    If lstSelectedFields.ListCount = 0 Then
    
        For i = 1 To v.Count
            v.Remove v.Count
        Next i
        
        For i = 1 To v1.Count
            v1.Remove v1.Count
        Next i
        
    End If
End Sub

Private Sub cmdAddAll_Click()
    
    Dim i As Integer
    Dim j As Integer
    
    Dim bfound As Boolean
    
    ClearCollections
    
    For j = 1 To lstTables.ListCount
    
        If v1.Count > 0 Then
            For i = 1 To v1.Count
                If lstTables.List(j - 1) = v1.Item(i) Then
                    bfound = True
                    Exit For
                End If
            Next i
        
       End If
    
     
       If Not bfound Then
            If v.Count > 0 Then
                For i = 1 To v.Count
                    If lstTables.List(j - 1) = v.Item(i) Then
                        bfound = True
                        Exit For
                    End If
                Next i
            Else
                bfound = False
           End If
       End If
       
       If Not bfound Then
            v.Add lstTables.List(j - 1)
            v1.Add lstTables.List(j - 1)
            lstSelectedFields.AddItem lstTables.List(j - 1)
       End If
    
    Next j
    
    For i = 1 To myTables.Count
        myTables(i).Isincluded = True
        
    Next
   
   
End Sub

Private Sub cmdAddEvent_Click()
    
    frmProperties.Caption = "New Event"
    frmProperties.Visible = True
    
    lblPropertyName = "Event Name"
    txtNewPropertyName.Text = ""
    txtArguments.Visible = True
    txtArguments.Text = ""
    lblArguments.Visible = True
    lblDataType.Visible = False
    cmbDataType.Visible = False
    
    enumInterfaceAdded = intEvent
    
    frmInterfaces.Visible = False
End Sub

Private Sub cmdAddMethod_Click()
    
    frmProperties.Caption = "New Method"
    frmProperties.Visible = True
    
    lblPropertyName = "Method Name"
    txtNewPropertyName.Text = ""
    txtArguments.Text = ""
    txtArguments.Visible = True
    lblArguments.Visible = True
    
    lblDataType.Caption = "Return Data Type"
    lblDataType.Visible = True
    cmbDataType.Visible = True
    cmbDataType.ListIndex = 0
    
    enumInterfaceAdded = interMethod
    
    frmInterfaces.Visible = False
    
End Sub

Private Sub cmdAddProperty_Click()
    
    
    
    frmProperties.Caption = "New Property"
    frmProperties.Visible = True
    
    txtNewPropertyName.Text = ""
    txtArguments.Visible = False
    lblArguments.Visible = False
    cmbDataType.ListIndex = 14
    lblDataType.Caption = "Data Type"
    lblDataType.Visible = True
    cmbDataType.Visible = True
    enumInterfaceAdded = interProperty
    
    frmInterfaces.Visible = False
    
End Sub

Private Sub cmdApply_Click()
    
    Dim i As Integer
    Dim oFields As Fields
    Dim j As Integer
    
    frmFields.Visible = False
    
    'Include checked fields
    frmTables.Visible = True
    
    If CurrentTableIndex <= 0 Then Exit Sub
    
    With LvwFields
    
        Set oFields = myTables(CurrentTableIndex).Fields
        
        For i = 1 To .ListItems.Count
            oFields(i).Isincluded = .ListItems(i).Checked
        Next
        
    End With
    
    
    
    
    
    
End Sub

Private Sub cmdCancel_Click()
    
    Dim vbAns As VbMsgBoxResult
    
    vbAns = MsgBox("Are you sure you want to cancel the creation of business objects ? ", vbYesNo + vbQuestion, "Object Builder")
    
    If vbAns = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdCancelFields_Click()
    frmFields.Visible = False
    frmTables.Visible = True
    
End Sub

Private Sub cmdChangeDir_Click()
    
    frmSelectDirectory.Visible = True
    txtDirectory.Enabled = False
    cmdChangeDir.Enabled = False
    
End Sub

Private Sub cmdFields_Click()

    CurrentTablename = txtPropertyName.Text
    
    frmFields.Visible = True
    frmFields.ZOrder 0
    frmFields.Caption = "Fields in " & CurrentTablename
    
    
    PopulateListview
    
End Sub

Private Sub PopulateListview()
    
    Dim i As Integer
    Dim j As Integer
    
    With LvwFields
        .ListItems.Clear
        For i = 1 To myTables.Count
            If myTables(i).tableName = CurrentTablename Then
                'CurrentTableIndex = myTables(i).TableId
                For j = 1 To myTables(i).Fields.Count
                    
                    .ListItems.Add , , myTables(i).Fields(j).Fieldname
                    
                    If myTables(i).Fields(j).Isincluded Then
                        .ListItems(j).Checked = True
                    End If
                    
                    .ListItems(j).ListSubItems.Add , , Convert(myTables(i).Fields(j).FieldType)
                    
                Next
                
                Exit For
            End If
    Next
        
        
    End With
End Sub

Private Function Convert(intType As Variant) As String
    
    Select Case intType
        Case 1
            Convert = intType
        Case 2
            Convert = "Integer"
        Case 3
            Convert = "Long"
        Case 4
            Convert = "Single"
        Case 5
            Convert = "Double"
        Case 6
            Convert = "Currency"
        Case 7
            Convert = "Date"
        Case 8
            Convert = intType
        Case 9
            Convert = intType
        Case 10
            Convert = intType
        Case 11
            Convert = "Boolean"
        Case 17
            Convert = "Byte"
        Case 200
            Convert = "String"
        Case 201
            Convert = "String"
        Case 205
            Convert = "Object"
        
        Case 135 'SQL data type
            Convert = "Date"
            
        Case Else
        
            Convert = intType
            
    End Select
    


End Function

Private Sub cmdFinish_Click()
    
    Dim lngRetVal As Long
    Dim Scr_hDC As Long
    
    Scr_hDC = GetDesktopWindow()
    
    cmdPrev.Enabled = False
    cmdCancel.Enabled = False
    
   'Create classes and then quit
   
   CreateClasses
   
   MsgBox "The classes have been successfully Created !. " & vbCrLf & "Your project will now be opened for editing. ", vbInformation + vbOKOnly, "Classes"
   
   Unload Me
   
   lngRetVal = ShellExecute(Scr_hDC, "Open", myTables.UserDirectory & "\" & strDirectory & ".vbp", 0&, 0&, 1)
   
   
End Sub

Private Sub cmdFolderCancel_Click()
    
    frmSelectDirectory.Visible = False
    txtDirectory.Enabled = True
    cmdChangeDir.Enabled = True
    
End Sub

Private Sub cmdNewFolder_Click()
    
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Dim strFolder  As String
    
    strFolder = InputBox$("Enter the name of the new folder.The folder will be created within the selected folder.", "New Folder")
    
    If Len(strFolder) > 0 Then
        ChDir DirFolders.List(DirFolders.ListIndex)
        Set fld = fso.CreateFolder(strFolder)
        
        DirFolders.Refresh
    End If
    
End Sub

Private Sub cmdNext_Click()

    Select Case intState
    
        Case 1
            
            frmTables.Visible = True
            ShowTables
            frmDatabase.Visible = False
            frmInterfaces.Visible = False
            frmDirectory.Visible = False
            frmFinished.Visible = False
            frmFields.Visible = False
            frmProperties.Visible = False
            
            intState = 2
            lblHelpText = state2
            cmdPrev.Enabled = True
            cmdFinish.Enabled = False
        Case 2
            
            frmTables.Visible = False
            frmDatabase.Visible = False
            frmInterfaces.Visible = True
            ShowInterfaces
            frmFields.Visible = False
            frmDirectory.Visible = False
            frmFinished.Visible = False
            frmProperties.Visible = False
            
            intState = 3
            lblHelpText = state3
            cmdPrev.Enabled = True
            cmdFinish.Enabled = False
        
        Case 3
            frmTables.Visible = False
            frmDatabase.Visible = False
            frmInterfaces.Visible = False
            frmDirectory.Visible = True
            ShowDirectory
            frmFinished.Visible = False
            frmFields.Visible = False
            frmProperties.Visible = False
            
            intState = 4
            lblHelpText = state4
            cmdPrev.Enabled = True
            cmdFinish.Enabled = False
            
        Case 4
        
            frmTables.Visible = False
            frmDatabase.Visible = False
            frmInterfaces.Visible = False
            
            If CheckDirectory Then
                frmDirectory.Visible = False
                frmFinished.Visible = True
                frmFields.Visible = False
                frmProperties.Visible = False
                intState = 5
                lblHelpText = state5
                lblProgressText = ""
                cmdPrev.Enabled = True
                cmdNext.Enabled = False
                cmdFinish.Enabled = True
                myTables.UserDirectory = txtDirectory.Text
            Else
                Dim vbAns As VbMsgBoxResult
                Dim fso As New FileSystemObject
                Dim fld As Folder
                
                vbAns = MsgBox("The folder specified does not exist. Do you want to create this folder ?", vbQuestion + vbYesNo, "Folder does not exist")
                If vbAns = vbYes Then
                    
                    On Error Resume Next
                    
                    Set fld = fso.CreateFolder(txtDirectory)
                    
                    If Err.Number = 76 Then
                        MsgBox "The folder specified could not be created, save your classes by changing to another directory", vbCritical + vbOKCancel, "Folder Error"
                        
                        frmDirectory.Visible = True
                        frmFinished.Visible = False
                        frmFields.Visible = False
                        frmProperties.Visible = False
                        
                        intState = 4
                        lblHelpText = state4
                        cmdPrev.Enabled = True
                        cmdNext.Enabled = True
                        cmdFinish.Enabled = False
                
                    Else
                    myTables.UserDirectory = fld.Path
                    
                        frmDirectory.Visible = False
                        frmFinished.Visible = True
                        frmFields.Visible = False
                        frmProperties.Visible = False
                        intState = 5
                        lblHelpText = state5
                        lblProgressText = ""
                        cmdPrev.Enabled = True
                        cmdNext.Enabled = False
                        cmdFinish.Enabled = True
                    End If
                    
                Else
                    frmDirectory.Visible = True
                    frmFinished.Visible = False
                    frmFields.Visible = False
                    frmProperties.Visible = False
                    
                    intState = 4
                    lblHelpText = state4
                    cmdPrev.Enabled = True
                    cmdNext.Enabled = True
                    cmdFinish.Enabled = False
                
                End If
                
            End If
            
                        
        Case 5
            cmdPrev.Enabled = False
            cmdNext.Enabled = False
            cmdFinish.Enabled = True
    End Select
End Sub

Private Sub ShowInterfaces()

    Dim i As Integer
    Dim j As Integer
    Dim nd As Node
    
    With tvwTables
        .Nodes.Clear
        
      Set nd = .Nodes.Add(, tvwFirst, "Classes", "Classes")
      nd.Expanded = True
      nd.Bold = True
      
        For i = 1 To myTables.Count
            
            If myTables(i).Isincluded Then
                
                Set nd = .Nodes.Add("Classes", tvwChild, myTables(i).Key, myTables(i).tableName)
                nd.ForeColor = vbBlue
                
            End If
        Next
        
    End With
    
    cmdAddEvent.Enabled = False
    cmdAddMethod.Enabled = False
    cmdAddProperty.Enabled = False
    
End Sub


Private Sub ShowTables()
    Dim i As Integer
    Dim j As Integer
    
    lstTables.Clear
    
    txtPropertyName.Text = ""
        
    
    With myTables
        For i = 1 To .Count
            lstTables.AddItem .Item(i).tableName
            If .Item(i).Isincluded = True Then
                If Not ItemExist(.Item(i).tableName) Then
                    lstSelectedFields.AddItem .Item(i).tableName
                End If
            End If
        Next
        
        lstTables.ListIndex = 0
        
                   
    End With
    
    
    
    
End Sub

Private Function ItemExist(strItem) As Boolean
    
    Dim i As Integer
    
    With lstSelectedFields
        For i = 0 To .ListCount - 1
            If strItem = .List(i) Then
                ItemExist = True
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub ShowDirectory()
    If Len(myTables.UserDirectory) = 0 Then
        txtDirectory.Text = myTables.DefaultDirectory
    Else
        txtDirectory.Text = myTables.UserDirectory
    End If
    
    frmSelectDirectory.Visible = False
    
End Sub

Private Function GetDirectory(strPath As String) As String
    Dim lngSeparatorPosition As Long
    
    Do
       lngSeparatorPosition = InStr(1, strPath, "\", vbTextCompare)
       If lngSeparatorPosition = 0 Then Exit Do
       strPath = Right$(strPath, Len(strPath) - lngSeparatorPosition)
       
    Loop
     
    GetDirectory = Replace(strPath, ".mdb", "", 1, vbTextCompare)
    
End Function

Private Sub IncludeAllFields(oTable As Table)
    
    Dim i As Integer
    
    For i = 1 To oTable.Fields.Count
        oTable.Fields(i).Isincluded = True
    Next
    
End Sub
Private Sub cmdPrev_Click()
    
    Select Case intState
    
        Case 1
            
            cmdPrev.Enabled = False
            cmdNext.Enabled = True
            cmdFinish.Enabled = False
            intState = 1
            lblHelpText = state1
            
        Case 2
            
            frmTables.Visible = False
            frmDatabase.Visible = True
            frmInterfaces.Visible = False
            frmDirectory.Visible = False
            frmFinished.Visible = False
            intState = 1
            lblHelpText = state1
            cmdPrev.Enabled = False
            cmdNext.Enabled = True
            cmdFinish.Enabled = False
        Case 3
            
            frmTables.Visible = True
            frmDatabase.Visible = False
            frmInterfaces.Visible = False
            frmDirectory.Visible = False
            frmFinished.Visible = False
            intState = 2
            lblHelpText = state2
            cmdPrev.Enabled = True
            cmdNext.Enabled = True
            cmdFinish.Enabled = False
        Case 4
            frmTables.Visible = False
            frmDatabase.Visible = False
            frmInterfaces.Visible = True
            frmDirectory.Visible = False
            frmFinished.Visible = False
            intState = 3
            lblHelpText = state3
            cmdPrev.Enabled = True
            cmdNext.Enabled = True
            cmdFinish.Enabled = False
         Case 5
            frmTables.Visible = False
            frmDatabase.Visible = False
            frmInterfaces.Visible = False
            frmDirectory.Visible = True
            frmFinished.Visible = False
            intState = 4
            lblHelpText = state4
            cmdPrev.Enabled = True
            cmdNext.Enabled = True
            cmdFinish.Enabled = False
            
    End Select
End Sub

Private Sub cmdPropApply_Click()

    'Add property/method/event to classes collection
    
    Dim i As Integer
    Dim nd As Node
    
    If enumInterfaceAdded = interProperty Then
    For i = 1 To myTables.Count
        If myTables(i).Key = CurrentKey Then
            myTables(i).Fields.Add txtNewPropertyName, txtNewPropertyName, cmbDataType.Text, myTables(i).Fields.Count + 1, True
           Set nd = tvwTables.Nodes.Add(myTables(i).Key, tvwChild, , myTables(i).Fields(myTables(i).Fields.Count).Fieldname)
           nd.ForeColor = vbRed
           nd.Bold = False
            Exit For
        End If
    Next
    
    Else
        
        If enumInterfaceAdded = interMethod Then
            
            For i = 1 To myTables.Count
                If myTables(i).Key = CurrentKey Then
                    myTables(i).Methods.Add txtNewPropertyName, myTables(i).Methods.Count + 1, txtNewPropertyName, cmbDataType.Text, txtArguments.Text, True
                    
                    Set nd = tvwTables.Nodes.Add(myTables(i).Key, tvwChild, , myTables(i).Methods(myTables(i).Methods.Count).MethodName)
                    nd.ForeColor = vbGreen
                    nd.Bold = False
                    Exit For
                End If
            Next
        
        Else 'event !
            For i = 1 To myTables.Count
                If myTables(i).Key = CurrentKey Then
                     myTables(i).Signals.Add txtNewPropertyName, myTables(i).Signals.Count + 1, txtNewPropertyName, txtArguments.Text, True
                    Set nd = tvwTables.Nodes.Add(myTables(i).Key, tvwChild, , myTables(i).Signals(myTables(i).Signals.Count).EventName)
                    
                    nd.BackColor = vbGreen
                    nd.Bold = False
                    Exit For
                End If
            Next
        
        End If
    
    End If
   
    frmInterfaces.Visible = True
    frmProperties.Visible = False
    
End Sub

Private Sub cmdPropCancel_Click()
    frmInterfaces.Visible = True
    frmProperties.Visible = False
End Sub

Private Sub cmdRemove_Click()

    Dim i As Integer
    If (lstSelectedFields.ListIndex <> -1) Then
       
       For i = 1 To myTables.Count
        If lstSelectedFields.List(lstSelectedFields.ListIndex) = myTables(i).tableName Then
            myTables(i).Isincluded = False
            Exit For
        End If
     Next
     
     lstSelectedFields.RemoveItem (lstSelectedFields.ListIndex)
     
    End If
End Sub

Private Sub cmdRemoveAll_Click()
    
    Dim i As Integer
    
    lstSelectedFields.Clear
    txtPropertyName.Text = ""
    
    For i = 1 To myTables.Count
        myTables(i).Isincluded = True
    Next
End Sub




Private Sub cmdSelect_Click()
    
    myTables.UserDirectory = DirFolders.List(DirFolders.ListIndex)
    
    frmSelectDirectory.Visible = False
    txtDirectory.Enabled = True
    cmdChangeDir.Enabled = True
    
    txtDirectory.Text = myTables.UserDirectory
    
End Sub

Private Sub cmdSQLApply_Click()
    
    Dim SQLConn As String
    Dim adoConn As ADODB.Connection
    Dim oTables As Tables
    Dim i As Integer
  
    If Len(txtServer.Text) > 0 Then
        
        strServer = txtServer.Text
        strInitCatalog = txtIntialCatalog
        
        On Error Resume Next
        
        txtDatabaseName.Text = strServer
        frmSQLServer.Visible = False
        frmDatabase.Visible = True
        SQLConn = "Persist Security Info=False;User ID=sa;Initial Catalog= " & strInitCatalog & " ; Data Source= " & strServer
        Me.Refresh
    
        Set adoConn = New ADODB.Connection
          
          With adoConn
              .Provider = SQLPROVIDER
              .ConnectionString = SQLConn
              .Open
              
              If adoConn.State = adStateOpen Then
                  cmdNext.Enabled = True
                  txtDatabaseName.Text = strServer & "\" & strInitCatalog
                  Me.Refresh
                  Me.MousePointer = 11
                  lblEstablishingConnection.Visible = True
                  lblEstablishingConnection = "Testing validity of database ......"
                  strConnectionString = "Provider = " & .Provider & " ;" & .ConnectionString
                  
                  strDirectory = strInitCatalog
                  
                  Set oTables = New Tables
                  
                  Set myTables = oTables.GetTables(strConnectionString)
                  myTables.DefaultDirectory = "C:\" & strDirectory
                  
                  For i = 1 To myTables.Count
                      IncludeAllFields myTables(i)
                  Next
                  
                  lblEstablishingConnection.Visible = False
                  Me.MousePointer = 0
                  
              Else
                  MsgBox "The database you selected cannot be accessed !", vbCritical + vbOKOnly, "Object Builder"
              End If
              
          End With
    
  Else
    If Len(txtIntialCatalog.Text) > 0 Then
        MsgBox "Server name is required to connect to the database !", vbCritical + vbOKOnly, "Server name Empty"
    Else
        MsgBox "Both Server and Default database  names are required !", vbCritical + vbOKOnly, "Server name Empty"
    End If
    
  End If
    
    
    
    
End Sub

Private Sub cmdSQLCancel_Click()
    
    frmSQLServer.Visible = False
    frmDatabase.Visible = True
    
End Sub

Private Sub DrvDrives_Change()
    On Error Resume Next
    DirFolders.Path = DrvDrives.List(DrvDrives.ListIndex)
End Sub

Private Sub Form_Load()
    
    cmdFinish.Enabled = False
    
    frmDatabase.ZOrder 0
    frmDatabase.Visible = True
    frmDirectory.Visible = False
    frmTables.Visible = False
    frmInterfaces.Visible = False
    frmFinished.Visible = False
    frmFields.Visible = False
    frmProperties.Visible = False
    frmSQLServer.Visible = False
    
    cmdOPenDatabase.Visible = False
    
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    
    intState = 1
    lblHelpText = state1
    lblEstablishingConnection.Visible = False
    optAccess.Value = False
    optSQL.Value = False
    
End Sub

Private Sub lstSelectedFields_Click()
    
    Dim strItem As String
    Dim i As Integer
    
    strItem = lstSelectedFields.List(lstSelectedFields.ListIndex)
    
    For i = 1 To myTables.Count
     If strItem = myTables(i).tableName Then
        CurrentTableIndex = myTables(i).TableID
        Exit For
    End If
    Next
    ' Get property name from objects if they exist
    
    txtPropertyName = strItem
    
End Sub


Private Sub optAccess_Click()

    Dim strFile As String
    Dim adoConn As ADODB.Connection
    Dim SQLConn As String
    Dim AccessConn As String
    
    Dim oTables As New Tables
    Dim i As Integer
    
    cdlgDatabase.Filter = "Databases(*.mdb)|*.mdb"
    
    cdlgDatabase.DialogTitle = "Select database to use"
    cdlgDatabase.CancelError = False
    cdlgDatabase.ShowOpen
    
    strFile = cdlgDatabase.FileName
    
    If Len(strFile) > 0 Then
        AccessConn = "Data Source=" & strFile
        strDirectory = GetDirectory(strFile)
    End If
    
    On Error Resume Next
    
    Set adoConn = New ADODB.Connection
        
        With adoConn
            .Provider = IIf(optAccess.Value = True, ACCESSPROVIDER, SQLPROVIDER)
            .ConnectionString = IIf(optAccess.Value = True, AccessConn, SQLConn)
            .Open
            
            If adoConn.State = adStateOpen Then
                cmdNext.Enabled = True
                txtDatabaseName.Text = strFile
                Me.Refresh
                Me.MousePointer = 11
                lblEstablishingConnection.Visible = True
                lblEstablishingConnection = "Testing validity of database ......"
                strConnectionString = "Provider = " & .Provider & " ;" & .ConnectionString
                
                Set myTables = oTables.GetTables(strConnectionString)
                myTables.DefaultDirectory = "C:\" & strDirectory
                
                For i = 1 To myTables.Count
                    IncludeAllFields myTables(i)
                Next
                
                lblEstablishingConnection.Visible = False
                Me.MousePointer = 0
                
            Else
                MsgBox "The database you selected cannot be accessed !", vbCritical + vbOKOnly, "Object Builder"
            End If
            
        End With
    
End Sub

Private Sub optSQL_Click()

    frmSQLServer.Visible = True
    frmDatabase.Visible = False
    frmDatabase.ZOrder 1
    frmSQLServer.ZOrder 0
    
End Sub

Private Sub tvwTables_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub tvwTables_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim i As Long
    
    If Node.Index = 1 Then
        Node.Expanded = True
        
        cmdAddEvent.Enabled = False
        cmdAddMethod.Enabled = False
        cmdAddProperty.Enabled = False
        
        Exit Sub
        
    Else
        CurrentKey = Node.Key
        cmdAddEvent.Enabled = True
        cmdAddMethod.Enabled = True
        cmdAddProperty.Enabled = True
        
        For i = 1 To myTables.Count
            If CurrentKey = myTables(i).Key Then
                Check1.Value = IIf(myTables(i).HasCollection, 1, 0)
            End If
        Next
        
    End If
    
End Sub
Private Function CheckDirectory() As Boolean
    Dim fso As New FileSystemObject
    
    CheckDirectory = fso.FolderExists(txtDirectory.Text)
    Set fso = Nothing
    
End Function

Private Sub CreateClasses()
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strArgs As String
    
    
    With myTables
        
        animProgress.Open "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Videos\Filemove.avi"
        animProgress.Play
        ProgressBar1.Max = countClasses + 1
        
        Open myTables.UserDirectory & "\" & strDirectory & ".vbp" For Output As #1
        
        lblProgressText = "Creating Project " & strDirectory & " ......."
        
        Print #1, "Type=OleDll"
        Print #1, "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\WINDOWS\SYSTEM\stdole2.tlb#OLE Automation"
        
        For i = 1 To myTables.Count
            If myTables(i).Isincluded Then
                Print #1, "Class=" & RemoveSpace(myTables(i).tableName) & "; " & RemoveSpace(myTables(i).tableName) & ".cls"
                If myTables(i).HasCollection Then
                    Print #1, "Class=" & RemoveSpace(myTables(i).tableName) & "Collection" & "; " & RemoveSpace(myTables(i).tableName) & "Collection" & ".cls"
                End If
            End If
        Next
        
        Print #1, "Startup = " & """" & "(" & "None" & ")" & """"
        Print #1, "Command32 = "" """
        Print #1, "Name = " & """" & strDirectory & """"
        Print #1, "HelpContextID = " & """" & "0" & """"
        Print #1, "CompatibleMode = " & """" & "1" & """"
        Print #1, "MajorVer = 1"
        Print #1, "MinorVer = 0"
        Print #1, "RevisionVer = 0"
        Print #1, "AutoIncrementVer = 0"
        Print #1, "ServerSupportFiles = 0"
        Print #1, "VersionCompanyName = " & """" & "Nathan Hassan Omukwenyi" & """"
        Print #1, "CompilationType = 0"
        Print #1, "OptimizationType = 0"
        Print #1, "FavorPentiumPro(tm) = 0"
        Print #1, "CodeViewDebugInfo = 0"
        Print #1, "NoAliasing = 0"
        Print #1, "BoundsCheck = 0"
        Print #1, "OverflowCheck = 0"
        Print #1, "FlPointCheck = 0"
        Print #1, "FDIVCheck = 0"
        Print #1, "UnroundedFP = 0"
        Print #1, "StartMode = 1"
        Print #1, "Unattended = 0"
        Print #1, "Retained = 0"
        Print #1, "ThreadPerObject = 0"
        Print #1, "MaxNumberOfThreads = 1"
        Print #1, "ThreadingModel = 1"
        Print #1, "        "
        Print #1, "[MS Transaction Server]"
        Print #1, "AutoRefresh = 1"
        
        Close #1
        
        ProgressBar1.Value = 1

    'create classes
        For i = 1 To myTables.Count
            If myTables(i).Isincluded Then
                    
                If myTables(i).HasCollection Then
                 'create collection of class
                    
                    Open myTables.UserDirectory & "\" & RemoveSpace(myTables(i).tableName) & "Collection" & ".cls" For Output As #1
                    
                    Print #1, "VERSION 1.0 CLASS"
                    
                    Print #1, "BEGIN"
                    Print #1, "  MultiUse = -1" & "  '" & "True"
                    Print #1, "  Persistable = 0" & "  '" & "NotPersistable"
                    Print #1, "  DataBindingBehavior = 0  ' vbNone"
                    Print #1, "  DataSourceBehavior = 0   ' vbNone"
                    Print #1, "  MTSTransactionMode = 0   ' NotAnMTSObject"
                    Print #1, "END"
                    
                    Print #1, "Attribute VB_Name = " & """" & RemoveSpace(myTables(i).tableName) & "Collection" & """"
                    Print #1, "Attribute VB_GlobalNameSpace = False"
                    Print #1, "Attribute VB_Creatable = False"
                    Print #1, "Attribute VB_PredeclaredId = False"
                    Print #1, "Attribute VB_Exposed = True"
                    Print #1, "Attribute VB_Ext_KEY = " & """" & "SavedWithClassBuilder6" & """" & "," & """" & "Yes" & """"
                    Print #1, "Attribute VB_Ext_KEY = " & """" & "Top_Level" & """" & "," & """" & "No" & """"
                    Print #1, "Attribute VB_Ext_KEY = " & """" & "Collection " & """" & "," & """" & " Field " & """"
                    Print #1, "Attribute VB_Ext_KEY = " & """" & "Member0 " & """" & "," & """" & "Field" & """"
                     
                    Print #1, "Option Explicit"
                    Print #1, ""
                    
                    Print #1, "Private mCol As Collection"
                    Print #1, ""
                    Print #1, ""
                        
                    strArgs = myTables(i).Fields(1).Fieldname & "  as " & Convert(myTables(i).Fields(1).FieldType) & ","
                    
                    For j = 2 To myTables(i).Fields.Count
                        strArgs = strArgs & myTables(i).Fields(j).Fieldname & "  as " & Convert(myTables(i).Fields(j).FieldType) & ","
                    Next
                    
                    strArgs = Left$(strArgs, Len(strArgs) - 1)
                    
                    Print #1, "Public Function Add(Key As String," & strArgs & ") As " & myTables(i).tableName
                    Print #1, ""
                    Print #1, "    Dim objNewMember As " & RemoveSpace(myTables(i).tableName)
                    Print #1, "    Set objNewMember = New  " & RemoveSpace(myTables(i).tableName)
                    Print #1, ""
                    Print #1, ""
                    Print #1, "   'set the properties passed into the method"
                    Print #1, ""
                    Print #1, "    objNewMember.Key = Key"
                    
                    For j = 1 To myTables(i).Fields.Count
                        Print #1, "    objNewMember." & myTables(i).Fields(j).Fieldname & " =  " & myTables(i).Fields(j).Fieldname
                    Next
                    
                    Print #1, "    mCol.Add objNewMember"
                    
                    Print #1, ""
                    Print #1, ""
                    Print #1, " 'return the object created"
                    Print #1, "    Set Add = objNewMember"
                    Print #1, "    Set objNewMember = Nothing"
                    Print #1, ""
                    Print #1, ""
                    Print #1, "End Function"
                    Print #1, ""
                    Print #1, "Public Property Get Item(vntIndexKey As Variant) As Field"
                    Print #1, ""
                    Print #1, "  Set Item = mCol(vntIndexKey)"
                    Print #1, "End Property"
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, "Public Property Get Count() As Long"
                    Print #1, ""
                    Print #1, "    Count = mCol.Count"
                    Print #1, "End Property"
                    Print #1, ""
                    Print #1, ""
                    Print #1, "Public Sub Remove(vntIndexKey As Variant)"
                    Print #1, ""
                    Print #1, ""
                    Print #1, "    mCol.Remove vntIndexKey"
                    Print #1, "End Sub"
                    Print #1, ""
                    Print #1, ""
                    Print #1, "Public Property Get NewEnum() As IUnknown"
                    Print #1, ""
                    Print #1, "    Set NewEnum = mCol.[_NewEnum]"
                    Print #1, "End Property"
                    Print #1, ""
                    Print #1, ""
                    Print #1, "Private Sub Class_Initialize()"
                    Print #1, "    Set mCol = New Collection"
                    Print #1, "End Sub"
                    Print #1, ""
                    Print #1, ""
                    Print #1, "Private Sub Class_Terminate()"
                    Print #1, "    Set mCol = Nothing"
                    Print #1, "End Sub"
    
                    Close #1
                
                End If
                    
                Open myTables.UserDirectory & "\" & RemoveSpace(myTables(i).tableName) & ".cls" For Output As #1
                lblProgressText = " Creating Class " & RemoveSpace(myTables(i).tableName) & " ........."
                
                
                Print #1, "VERSION 1.0 CLASS"
                Print #1, "BEGIN"
                Print #1, "  MultiUse = -1" & "  '" & "True"
                Print #1, "  Persistable = 0" & "  '" & "NotPersistable"
                Print #1, "  DataBindingBehavior = 0  ' vbNone"
                Print #1, "  DataSourceBehavior = 0   ' vbNone"
                Print #1, "  MTSTransactionMode = 0   ' NotAnMTSObject"
                Print #1, "END"
                Print #1, "Attribute VB_Name = " & """" & RemoveSpace(myTables(i).tableName) & """"
                Print #1, "Attribute VB_GlobalNameSpace = False"
                Print #1, "Attribute VB_Creatable = False"
                Print #1, "Attribute VB_PredeclaredId = False"
                Print #1, "Attribute VB_Exposed = True"
                Print #1, "Attribute VB_Ext_KEY = " & """" & "SavedWithClassBuilder6" & """" & "," & """" & "Yes" & """"
                Print #1, "Attribute VB_Ext_KEY = " & """" & "Top_Level" & """" & "," & """" & "No" & """"
                Print #1, "Option Explicit"
                Print #1, "            "
                Print #1, "Public Key As String"
                Print #1, "  "
                Print #1, "  "
                
                For j = 1 To myTables(i).Fields.Count
                    If myTables(i).Fields(j).Isincluded Then
                        Print #1, "Private mvar" & RemoveSpace(myTables(i).Fields(j).Fieldname) & " as " & Convert(myTables(i).Fields(j).FieldType)
                    End If
                Next
                
                Print #1, "      "
                
                'Common Methods and procedures
                
                Print #1, "Private m_bflgIsNew As Boolean ' indicates that an object is new"
                Print #1, "Private m_bflgIsDirty As Boolean ' indicates that one of the object's properties has changed"
                Print #1, "Private m_bflgDelete As Boolean ' marks the object for deletion"
                Print #1, "Private m_bflgEditing As Boolean ' marks the object ready for editing"
                
                Print #1, "  "
                Print #1, "  "
                
                For k = 1 To myTables(i).Signals.Count
                    If myTables(i).Signals(k).Isincluded Then
                        Print #1, "Public Event " & RemoveSpace(myTables(i).Signals(k).EventName) & "(" & myTables(i).Signals(k).EventArguments & ")"
                    End If
                Next
                
                Print #1, "     "
                
                
                Print #1, "Public Property Get IsNew() As Boolean"
                Print #1, "'// Note: this property is read-only."
                Print #1, ""
                Print #1, "    IsNew = m_bflgIsNew"
                Print #1, ""
                Print #1, "End Property"
                Print #1, ""
                Print #1, "Public Property Get IsDirty() As Boolean"
                Print #1, "'// Note: this prpoerty is read-only."
                Print #1, ""
                Print #1, "    IsDirty = m_bflgIsDirty"
                Print #1, ""
                Print #1, "End Property"
                Print #1, ""
                Print #1, "Private Sub Class_Initialize()"
                Print #1, ""
                Print #1, "    m_bflgIsNew = True"
                Print #1, "    m_bflgIsDirty = False"
                Print #1, "    m_bflgDelete = False"
                Print #1, "    m_bflgEditing = False"
                Print #1, ""
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Public Sub Load(vID As Variant)"
                Print #1, "'// This method is used to retrieve an object's state e.g."
                Print #1, "'// from the database. The argument can be a string, integer etc"
                Print #1, "'// If successful, then the object is no longer marked as new"
                Print #1, ""
                Print #1, "    '// call method/s to load the objects data"
                Print #1, ""
                Print #1, "    '// If (object's data is found) Then"
                Print #1, "    '//     m_bflgIsNew = False"
                Print #1, "    '// Else"
                Print #1, "    '//     m_bflgIsNew = True"
                Print #1, "    '// End if"
                Print #1, ""
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Public Sub BeginEdit()"
                Print #1, "'// This method marks the object ready for editing. Any property Lets/Sets used"
                Print #1, "'// before this method is called will raise an error"
                Print #1, ""
                Print #1, "    m_bflgEditing = True"
                Print #1, ""
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Public Sub CancelEdit()"
                Print #1, "'// This method cancels any edits made to the object"
                Print #1, ""
                Print #1, "    m_bflgEditing = False"
                Print #1, ""
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Public Sub Delete()"
                Print #1, "'// This method marks the object for deletion when the save method is called."
                Print #1, ""
                Print #1, "    ' First; check; whether; BeginEdit; has; been; called"
                Print #1, "    If Not m_bflgEditing Then Err.Raise 1000000"
                Print #1, ""
                Print #1, "    m_bflgDelete = True"
                Print #1, ""
                Print #1, "End Sub"
                Print #1, ""
                Print #1, "Public Sub Save()"
                Print #1, "'// This method is used to save the object's data"
                Print #1, ""
                Print #1, "    ' first check whether BeginEdit has been called"
                Print #1, "    If Not m_bflgEditing Then Err.Raise 1000000"
                Print #1, ""
                Print #1, "    If m_bflgDelete Then"
                Print #1, "        ' delete the object"
                Print #1, "        m_bflgDelete = False"
                Print #1, "        m_bflgEditing = False"
                Print #1, "        '// clear all properties"
                Print #1, "        m_bflgIsNew = True"
                Print #1, "    Else"
                Print #1, "        ' save the object"
                Print #1, "        m_bflgEditing = False"
                Print #1, "        m_bflgIsDirty = False"
                Print #1, "    End If"
                Print #1, ""
                Print #1, "End Sub"

                'Add specific properties and methods
                
                For j = 1 To myTables(i).Fields.Count
                    If myTables(i).Fields(j).Isincluded Then
                    
                        Print #1, "Public Property Get  " & RemoveSpace(myTables(i).Fields(j).Fieldname) & "() as " & Convert(myTables(i).Fields(j).FieldType)
                        Print #1, "     " & RemoveSpace(myTables(i).Fields(j).Fieldname) & " = " & "mvar" & RemoveSpace(myTables(i).Fields(j).Fieldname)
                        Print #1, "End Property"
                        
                        Print #1, "  "
                        Print #1, "  "
                        
                        If InStr(1, myTables(i).Fields(j).FieldType, "Obj", vbTextCompare) > 0 Then
                            Print #1, "Public Property set  " & RemoveSpace(myTables(i).Fields(j).Fieldname) & "(Byval oNewValue as " & Convert(myTables(i).Fields(j).FieldType) & ")"
                            Print #1, "     set mvar" & RemoveSpace(myTables(i).Fields(j).Fieldname) & " = " & "oNewValue"
                            Print #1, "End Property"
                            Print #1, "  "
                            Print #1, "  "
                            
                        Else
                        
                            Print #1, "Public Property   Let " & RemoveSpace(myTables(i).Fields(j).Fieldname) & "(Byval vNewValue as " & Convert(myTables(i).Fields(j).FieldType) & ")"
                            Print #1, "     mvar" & RemoveSpace(myTables(i).Fields(j).Fieldname) & " = " & "vNewValue"
                            Print #1, "End Property"
                            Print #1, "  "
                            Print #1, "  "
                        
                        End If
                        
                        
                        
                    End If
                Next
                
                Print #1, "  "
                Print #1, " "
                
                For j = 1 To myTables(i).Methods.Count
                    If myTables(i).Methods(j).Isincluded Then
                        If myTables(i).Methods(j).ReturnType = "(None)" Then
                            
                            Print #1, "   "
                            Print #1, "Public sub " & RemoveSpace(myTables(i).Methods(j).MethodName) & "(" & myTables(i).Methods(j).MethodArguments & ")"
                            Print #1, "      "
                            Print #1, "End sub"
                            Print #1, "      "
                            
                        Else
                        
                            Print #1, "   "
                            Print #1, "Public Function " & RemoveSpace(myTables(i).Methods(j).MethodName) & "(" & myTables(i).Methods(j).MethodArguments & ")" & " as " & Convert(myTables(i).Methods(j).ReturnType)
                            Print #1, "      "
                            Print #1, "End Function"
                            Print #1, "      "
                            
                        End If
                        
                
                    End If
                    
                    
                    
                Next
                
                
                Close #1
                
                ProgressBar1.Value = ProgressBar1.Value + 1
            End If
            
        Next
        
    End With
    
    lblProgressText = "Finished creating classes"
    
    animProgress.Stop
    animProgress.Visible = False
    
End Sub


Private Function RemoveSpace(strString As String) As String
    Dim i As Long
    Do
        i = 0
        i = InStr(1, strString, " ", vbTextCompare)
        If i = 0 Then Exit Do
        strString = Replace(strString, " ", "", , , vbTextCompare)
        
    Loop
    
    RemoveSpace = strString
    
End Function


Private Function countClasses() As Integer
    Dim i As Integer
    
    For i = 1 To myTables.Count
        If myTables(i).Isincluded Then
            countClasses = countClasses + 1
        End If
    Next
End Function
