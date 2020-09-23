VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Fields"
Attribute VB_Ext_KEY = "Member1" ,"Methods"
Attribute VB_Ext_KEY = "Member2" ,"Signals"
Option Explicit

Public Key As String

Private mvartableName As String
Private mvarFields As Fields
Private mvarTableID As Integer
Private mvarIsIncluded As Boolean

Private mvarMethods As Methods
Private mvarSignals As Signals
Private mvarHasCollection As Boolean
Public Property Let HasCollection(ByVal vData As Boolean)

    mvarHasCollection = vData
End Property


Public Property Get HasCollection() As Boolean

    HasCollection = mvarHasCollection
End Property

Public Property Get Signals() As Signals
    If mvarSignals Is Nothing Then
        Set mvarSignals = New Signals
    End If

    Set Signals = mvarSignals
End Property


Public Property Set Signals(vData As Signals)
    Set mvarSignals = vData
End Property



Public Property Get Methods() As Methods
    If mvarMethods Is Nothing Then
        Set mvarMethods = New Methods
    End If


    Set Methods = mvarMethods
End Property


Public Property Set Methods(vData As Methods)
    Set mvarMethods = vData
End Property



Public Property Let TableID(ByVal vData As Integer)
    mvarTableID = vData
End Property


Public Property Get TableID() As Integer
    TableID = mvarTableID
End Property


Public Property Get Fields() As Fields
    If mvarFields Is Nothing Then
        Set mvarFields = New Fields
    End If


    Set Fields = mvarFields
End Property


Public Property Set Fields(vData As Fields)
    Set mvarFields = vData
End Property
Private Sub Class_Terminate()
  Set mvarSignals = Nothing
  Set mvarMethods = Nothing
  Set mvarFields = Nothing
End Sub


Public Property Let tableName(ByVal vData As String)
    mvartableName = vData
End Property


Public Property Get tableName() As String

    tableName = mvartableName
End Property


Public Property Get Isincluded() As Boolean
    Isincluded = mvarIsIncluded
End Property

Public Property Let Isincluded(ByVal vNewValue As Boolean)
    mvarIsIncluded = vNewValue
End Property
