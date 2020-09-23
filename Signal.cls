VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Signal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"EventArguments"
Option Explicit

Public Key As String

Private mvarEventName As String
Private mvarEventID As Integer
Private mvarEventArguments As String
Private mvarIsIncluded As Boolean

Public Property Let Isincluded(ByVal vData As Boolean)
    mvarIsIncluded = vData
End Property

Public Property Get Isincluded() As Boolean
    Isincluded = mvarIsIncluded
End Property

Public Property Let EventArguments(ByVal vData As String)

    mvarEventArguments = vData
End Property


Public Property Get EventArguments() As String

    EventArguments = mvarEventArguments
End Property

Public Property Let EventID(ByVal vData As Integer)
    mvarEventID = vData
End Property


Public Property Get EventID() As Integer

    EventID = mvarEventID
End Property


Public Property Let EventName(ByVal vData As String)

    mvarEventName = vData
End Property


Public Property Get EventName() As String

    EventName = mvarEventName
End Property



