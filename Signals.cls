VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Signals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Signal"
Attribute VB_Ext_KEY = "Member0" ,"Signal"
Option Explicit

Private mCol As Collection

Public Function Add(Key As String, EventID As Integer, EventName As String, EventArguments As String, Isincluded As Boolean) As Signal
    
    Dim objNewMember As Signal
    Set objNewMember = New Signal


    objNewMember.Key = Key
    objNewMember.EventName = EventName
    objNewMember.EventArguments = EventArguments
    objNewMember.EventID = EventID
    objNewMember.Isincluded = Isincluded
    
    mCol.Add objNewMember
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Signal
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property


Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
 
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

