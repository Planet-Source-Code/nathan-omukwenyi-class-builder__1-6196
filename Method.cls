VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Method"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Arguments"
Option Explicit

Public Key As String

Private mvarMethodName As String 'local copy
Private mvarReturnType As String 'local copy
Private mvarMethodID As Integer 'local copy

Private mvarMethodArguments As String 'local copy
Private mvarIsIncluded As Boolean 'local copy
Public Property Let Isincluded(ByVal vData As Boolean)

    mvarIsIncluded = vData
End Property


Public Property Get Isincluded() As Boolean

    Isincluded = mvarIsIncluded
End Property



Public Property Let MethodArguments(ByVal vData As String)
    mvarMethodArguments = vData
End Property


Public Property Get MethodArguments() As String

    MethodArguments = mvarMethodArguments
End Property


Public Property Let MethodID(ByVal vData As Integer)
    mvarMethodID = vData
End Property


Public Property Get MethodID() As Integer
    MethodID = mvarMethodID
End Property

Public Property Let ReturnType(ByVal vData As String)

    mvarReturnType = vData
End Property


Public Property Get ReturnType() As String
    ReturnType = mvarReturnType
End Property


Public Property Let MethodName(ByVal vData As String)
    mvarMethodName = vData
End Property


Public Property Get MethodName() As String
    MethodName = mvarMethodName
End Property



