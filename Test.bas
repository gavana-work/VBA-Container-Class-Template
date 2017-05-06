Attribute VB_Name = "Test"
Public cP1 As clsParent1
Public cC1P2 As clsChild1Parent2
Public cC2 As clsChild2

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MAIN

Public Sub Run()
    SetContents
    DisplayContents
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SECONDARY SUBS

Private Sub SetContents()
    
    Dim i As Long
    Dim j As Long
    i = 1
    j = 1
    
    Set cP1 = New clsParent1
    cP1.Name = "All Father Lloyd"
    
    For i = 1 To 10
    
        Set cC1P2 = New clsChild1Parent2
        cC1P2.Name = "Way of White Follower " & CStr(i)
        
        For j = 1 To 5
        
            Set cC2 = New clsChild2
            cC2.Name = "Miracle " & CStr(j)
            cC1P2.Add cC2
        
        Next j
        
        cP1.Add cC1P2
    
    Next i
    
End Sub

Private Sub DisplayContents()

    Dim i As Long
    Dim j As Long
    i = 1
    j = 1
    
    'cP1 Info
    Debug.Print vbNewLine & "Instance Name:  " & cP1.Name & "                             " & " Child Count:  " & cP1.ChildCount
    
    For i = 1 To cP1.ChildCount
        'cC1P2 Info
        Debug.Print vbNewLine & "                " & "Instance Name:  " & cP1.Child(i).Name & "      " & " Child Count:  " & cP1.Child(i).ChildCount
        
        For j = 1 To cP1.Child(i).ChildCount
        
            'cC2 Info
            Debug.Print vbNewLine & "                                " & "Instance Name:  " & cP1.Child(i).Child(j).Name
            
        Next j
        
    Next i
    
End Sub
