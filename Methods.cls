VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Methods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Method"
Attribute VB_Ext_KEY = "Member0" ,"Method"
Option Explicit

Private mCol As Collection

Public Function Add(Key As String, MethodID As Integer, MethodName As String, ReturnType As String, MethodArguments As String, Isincluded As Boolean) As Method
       
    Dim objNewMember As Method
    Set objNewMember = New Method
    
    objNewMember.Key = Key
    objNewMember.MethodName = MethodName
    objNewMember.ReturnType = ReturnType
    objNewMember.MethodID = MethodID
    objNewMember.MethodArguments = MethodArguments
    objNewMember.Isincluded = Isincluded
    
    mCol.Add objNewMember
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Method
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

