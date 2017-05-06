VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChild1Parent2"
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

'clsChild1Parent2
'Child Class to clsParent1
'A Collection Class that is the Parent of all clsChild2 instances
    
'Parent Class Variables
Private mcollChildren2 As Collection

'Child Class Variables
Private mlID As Long
Private mlParentPtr As LongLong 'Long or Long Ptr in 32-Bit systems
Private mclsParent As clsParent1

'Other Variables
Private pName As String

'This sub must be the first sub in your class
'****PtrSafe ensures that this API call can interact with a 64-bit system***************
'****LongPtr must be used to allow 64-bit systems to utilize this***********************
'****Otherwise, in a 32-Bit system PtrSafe must be omitted and Long or LongPtr is used**
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, Source As Any, ByVal bytes As LongLong) 'Accesses the windows API


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GENERAL

Private Sub Class_Initialize()
    Set mcollChildren2 = New Collection
    'Add reference to parent
    Set mclsParent = cP1
End Sub

Private Sub Class_Terminate()
    Set mcollChildren2 = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MANDATORY CHILD CLASS METHODS AND PROPERTIES

Private Function ObjFromPtr(ByVal pObj As LongLong) As Object 'Long or Long Ptr in 32-Bit systems
    Dim obj As Object
    CopyMemory obj, pObj, 4
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
End Function

Public Property Set ParentalRights(ByVal cParent As clsParent1): Set mclsParent = cParent: End Property
Public Property Get ParentalRights() As clsParent1: Set ParentalRights = mclsParent: End Property

Public Property Let prID(ByVal ID As Long): mlID = ID: End Property
Public Property Get prID() As Long: prID = mlID: End Property

Public Property Get Parent() As clsParent1: Set Parent = ObjFromPtr(mlParentPtr): End Property
Public Property Set Parent(obj As clsParent1): mlParentPtr = ObjPtr(obj): End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MANDATORY PARENT CLASS METHODS AND PROPERTIES

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mcollChildren2.[_NewEnum]
End Property

Public Sub Add(cChild As clsChild2)
    If cChild.prID = 0 Then
        cChild.prID = Me.ChildCount + 1
    End If
    Set cChild.Parent = Me
    mcollChildren2.Add cChild, CStr(cChild.prID)
End Sub

Public Property Get Child(vItem As Variant) As clsChild2
    Set Child = mcollChildren2.Item(vItem)
End Property

Public Property Get ChildCount() As Long
    ChildCount = mcollChildren2.Count
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


