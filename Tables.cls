VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Table"
Attribute VB_Ext_KEY = "Member0" ,"Table"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit



Private mvarUserDirectory As String
Private mvarDefaultDirectory As String

'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, tableName As String, Fields As Fields, TableID As Integer, Isincluded As Boolean, Optional sKey As String) As Table
    'create a new object
    Dim objNewMember As Table
    Set objNewMember = New Table


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.tableName = tableName
    Set objNewMember.Fields = Fields
    objNewMember.TableID = TableID
    objNewMember.Isincluded = Isincluded
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Table
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

Public Function GetTables(strConnectionString As String) As Tables
    
    Dim adoConn As ADODB.Connection
    Dim rsTables As ADODB.Recordset
    Dim oTables As Tables
    Dim rsFields As ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    
    Set adoConn = New ADODB.Connection
    
    With adoConn
        .ConnectionString = strConnectionString
        .Open
        Set rsTables = .OpenSchema(adSchemaTables)
        
    End With
    
    With rsTables
        .MoveFirst
        
        Set oTables = New Tables
        
        Do While Not .EOF
            
            If (InStr(1, .Fields("Table_Name").Value, "sys", vbTextCompare) = 0) Then
                 j = j + 1
                oTables.Add .Fields("Table_Name").Value, .Fields("Table_Name").Value, Nothing, j, False
                               
            End If
            .MoveNext
        Loop
    
    End With
    
    On Error Resume Next
    
    Dim intTablesCount As Integer
    
    intTablesCount = oTables.Count
    
    For j = 1 To intTablesCount
        
        Set rsFields = New ADODB.Recordset
        
        With rsFields
                  
             .Open "[" & oTables(j).tableName & "]", adoConn
             
            
            For i = 0 To .Fields.Count - 1
                oTables(j).Fields.Add " ", RemoveUnderScore(.Fields(i).Name), .Fields(i).Type, i, True
            Next
            
        
        End With
        
        Set rsFields = Nothing
        
    Next
    
    Dim oTab As Tables
    
    Set oTab = New Tables
    
    i = 0: j = 0
    
    'Remove database objects without fields
    
    For i = 1 To oTables.Count
        If oTables(i).Fields.Count > 0 Then
            j = j + 1
            oTab.Add oTables(i).Key, oTables(i).tableName, oTables(i).Fields, j, False
        End If
    Next i
        

    
    Set GetTables = oTab
    
    Set rsTables = Nothing
    Set oTab = Nothing
    Set oTables = Nothing
    
End Function

Private Function RemoveUnderScore(strFieldName As String) As String
    
    Dim i As Long
    
    i = InStr(1, strFieldName, "_", vbTextCompare)
    
    RemoveUnderScore = StrConv(Right$(strFieldName, Len(strFieldName) - i), vbProperCase)
    
End Function


Public Property Get UserDirectory() As String
    UserDirectory = mvarUserDirectory
End Property

Public Property Let UserDirectory(ByVal vNewValue As String)
    mvarUserDirectory = vNewValue
End Property

Public Property Get DefaultDirectory() As String
    DefaultDirectory = mvarDefaultDirectory
End Property

Public Property Let DefaultDirectory(ByVal vNewValue As String)
    mvarDefaultDirectory = vNewValue
End Property
