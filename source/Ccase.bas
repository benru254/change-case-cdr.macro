Attribute VB_Name = "Ccase"
Option Explicit
#If VBA7 Then  'this declare allows it to run in 64bit or 32bit CorelDRAW
    Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                            (ByVal hWnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) As Long
#Else
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                            (ByVal hWnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) As Long
#End If
                            


Sub Menu()
    frmChangeCase.Show vbModeless
End Sub
Sub TitleCase()
    Dim s As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    ActiveDocument.BeginCommandGroup ("Change Case to Title case")
    For Each s In sr
        If s.Type = cdrTextShape Then
            On Error Resume Next
            s.Text.Story.ChangeCase cdrTextTitleCase
        End If 'textleCase
    Next s
    ActiveDocument.EndCommandGroup '("Change Case to Title case")
    
End Sub 'Title case
Sub UpperCase()
    Dim s As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    ActiveDocument.BeginCommandGroup ("Change Case to Uppercase")
    For Each s In sr
        If s.Type = cdrTextShape Then
            On Error Resume Next
            s.Text.Story.ChangeCase cdrTextUpperCase
        End If 'text
    Next s
    ActiveDocument.EndCommandGroup ' ("Change Case to Uppercase")
    
End Sub ' UpperCase
Sub LowerCase()
    Dim s As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    ActiveDocument.BeginCommandGroup ("Change Case to Lower")
    For Each s In sr
        If s.Type = cdrTextShape Then
            On Error Resume Next
            s.Text.Story.ChangeCase cdrTextLowerCase
        End If 'text
    Next s
    ActiveDocument.EndCommandGroup '("Change Case to Lower")
    
End Sub ' LowerCase
Sub ToggleCase()
    Dim s As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    ActiveDocument.BeginCommandGroup ("Toggle Case")
    For Each s In sr
        If s.Type = cdrTextShape Then
            On Error Resume Next
            s.Text.Story.ChangeCase cdrTextToggleCase
        End If 'text
    Next s
    ActiveDocument.EndCommandGroup ' ("Toggle Case")

End Sub ' ToggleCase
Sub SentenceCase()
    Dim s As Shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    ActiveDocument.BeginCommandGroup ("Change to Sentence Case")
    For Each s In sr
        If s.Type = cdrTextShape Then
            On Error Resume Next
            s.Text.Story.ChangeCase cdrTextSentenceCase
        End If 'text
    Next s
    ActiveDocument.EndCommandGroup '("Change to Sentence Case")

End Sub 'Sentence Case



