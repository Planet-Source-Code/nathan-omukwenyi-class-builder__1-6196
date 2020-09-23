VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarFieldname As String 'local copy
Private mvarFieldType As String 'local copy
Private mvarFieldID As Integer 'local copy
Private mvarIsIncluded As Boolean

Public Property Get Isincluded() As Boolean
    Isincluded = mvarIsIncluded
End Property

Public Property Let Isincluded(ByVal vNewValue As Boolean)
    mvarIsIncluded = vNewValue
End Property

Public Property Let FieldID(ByVal vData As Integer)

    mvarFieldID = vData
End Property


Public Property Get FieldID() As Integer
    FieldID = mvarFieldID
End Property



Public Property Let FieldType(ByVal vData As String)
    mvarFieldType = vData
End Property


Public Property Get FieldType() As String
    FieldType = mvarFieldType
End Property



Public Property Let Fieldname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Fieldname = 5
    mvarFieldname = vData
End Property


Public Property Get Fieldname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Fieldname
    Fieldname = mvarFieldname
End Property



