VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParent1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'THE NECESSARY STEPS TO ALLOW CLASSES TO CONTAIN OTHER CLASSES!!!!
'1) In the VBA IDE export this module
'2) Open the file in your text editor of choice
'3) You will see the normally hidden text below:

    'VERSION 1.0 CLASS
    'BEGIN
    '  MultiUse = -1  'True
    'End
    'Attribute VB_Name = "clsProjects"
    'Attribute VB_GlobalNameSpace = False
    'Attribute VB_Creatable = False
    'Attribute VB_PredeclaredId = False
    'Attribute VB_Exposed = False

'4) Set Attribute VB_PredeclaredId = True
'5) Save your file
'6) Import your class module to the VBA IDE

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'clsParent1
'A Collection Class that is the Parent of all clsChild1Parent2 instances

'Parent Class Variable
Private mcollChildren1 As Collection

'Other Variables
Private pName As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GENERAL

Private Sub Class_Initialize()
    Set mcollChildren1 = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcollChildren1 = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MANDATORY PARENT CLASS METHODS AND PROPERTIES

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mcollChildren1.[_NewEnum]
End Property

Public Sub Add(cChild As clsChild1Parent2)
    If cChild.prID = 0 Then
        cChild.prID = Me.ChildCount + 1
    End If
    Set cChild.Parent = Me
    mcollChildren1.Add cChild, CStr(cChild.prID)
End Sub

Public Property Get Child(vItem As Variant) As clsChild1Parent2
    Set Child = mcollChildren1.Item(vItem)
End Property

Public Property Get ChildCount() As Long
    ChildCount = mcollChildren1.Count
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OTHER PROPERTIES

'1
Public Property Let Name(Value As String)
    pName = Value
End Property
Public Property Get Name() As String
    Name = pName
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OTHER METHODS

'1
Public Sub CalcSomething(Value As String)
    'Do something
End Sub

'2
Public Sub CalcSomethingElse(Value As Double)
    'Do something
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
