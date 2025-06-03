sub MakePerms()
    Dim plan As String
    Dim cpText As String
    Dim perms(1 To 5) As String
    Dim descriptions(1 to 5) As String
    Dim objectText As String
    Dim useText As String
    Dim output As String
    Dim i As Integer
    ' Check if user has selected any text
    If Selection.Type = wdSelectionNormal Then
        cpText = Selection.Text
    Else
        MsgBox "Please select some text first."
        Exit Sub
    End If

    plan = "The United States ought to become party to the United Nations Convention on the Law of the Sea."
    objectText = "the United Nations Convention on the Law of the Sea"
    
    useText = Replace(cpText, "[xxx]", objectText)
    useText = Replace(useText, "[xxxx]", objectText)
    
    ' Generate permutations
    perms(1) = Replace(useText, objectText, Striked(objectText))
    perms(2) = Replace(useText, objectText, "the " & Striked("United Nations") & " Convention " & Striked("on the Law of the Sea"))
    perms(3) = Replace(useText, objectText, "the United Nations " & Striked("Convention on the Law of the Sea"))
    perms(4) = Replace(useText, objectText, "the United Nations " & Striked("Convention on the ") & "Law " & Striked("of the Sea"))
    perms(5) = Replace(useText, objectText, "the " & Striked("United Nations Convention on the ") & "Law " & Striked("of the Sea"))
    
    descriptions(1) = "1---Other Issues"
    descriptions(2) = "2---The Convention"
    descriptions(3) = "3---United Nations"
    descriptions(4) = "4---United Nations Law"
    descriptions(5) = "5---The Law"
    
    output = ""
    
    For i = 1 To 5
        output = output & descriptions(i) & vbCrLf & plan & " " & perms(i) & vbCrLf & vbCrLf
    Next i

    ' Output to clipboard
    Dim DataObj As Object
    Set DataObj = CreateObject("MSForms.DataObject")
    DataObj.SetText output
    DataObj.PutInClipboard
    
    MsgBox "Permutations copied to clipboard!" & vbCrLf & vbCrLf & output, vbOKOnly, "Perm Generator"
End Sub

Function Striked(text As String) As String
    Dim i As Integer
    Dim result As String
    result = ""
    For i = 1 To Len(text)
        result = result & Mid(text, i, 1) & ChrW(&H336)
    Next i
    Striked = result
End Function
