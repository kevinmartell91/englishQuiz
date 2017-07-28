Option Explicit

Public qstCllSortedQuestions As Collection
Private qstCllUnsortedQuestion As Collection
Private qstQuestion As CQuestion
Dim blnValidData As Boolean




Private Sub Class_Initialize()
    Set qstCllSortedQuestions = New Collection
    Set qstQuestion = New CQuestion
    blnValidData = False
    
    Call LoadDataFromTXT
    Call CreateDisortedQuizCollection
  
End Sub

Public Property Get SortedQuestions() As Collection
    Set SortedQuestions = qstCllSortedQuestions
End Property

Public Property Get UnsortedQuestions() As Collection
    Set UnsortedQuestions = qstCllUnsortedQuestion
End Property

Public Property Let AddQuestion(Question As CQuestion)
    ' add question object in list
    qstCllQuestions.Add Question
End Property

Public Function ValidateDataFromTXT() As Boolean
      
    Dim fileName As String, textData As String, textRow As String, fileNo As Integer
    Dim strKey As String
    Dim intContQuestions As Integer, intContAnswers As Integer
    
    ' File directory where the .txt is
    fileName = "C:\Users\Kevin\Documents\Projects\EnglishPowerPointApp\testText.txt"
    fileNo = FreeFile 'Get first free file number
        
    Open fileName For Input As #fileNo
    Do While Not EOF(fileNo)
       Line Input #fileNo, textRow
       textData = textData & textRow

       If StrComp(strKey, "q:") = 0 Then
            intContQuestions = intContQuestions + 1
       End If
       If StrComp(strKey, "a:") = 0 Then
            intContAnswers = intContAnswers + 1
       End If
       
       
    Loop
    Close #fileNo
    
    If intContAnswers = intContQuestions Then
        Set ValidateDataFromTXT = True
        Debug.Print "iquals"
    Else
        Set ValidateDataFromTXT = False
    End If

End Function

Public Function LoadDataFromTXT()

    Dim fileName As String, textData As String, textRow As String, fileNo As Integer
    Dim strKey As String
    Dim intContRow As Integer, intEmpNum As Integer
    Dim Id As String
    Dim varAnswer As Variant, varQuestion As Variant
    
    fileName = "C:\Users\Kevin\Documents\Projects\EnglishPowerPointApp\testText.txt"
    fileNo = FreeFile 'Get first free file number
        
    intContRow = 0
    Open fileName For Input As #fileNo
    Do While Not EOF(fileNo)
       Line Input #fileNo, textRow
       textData = textData & textRow
       
       strKey = Left(textRow, 2)
       If StrComp(strKey, "q:") = 0 Then
            ' Mid
            ' To extract a substring, starting in the middle of a string, use Mid.
            ' Mid("example text", 9, 2) =  "te"
            ' started at position 9 (t) with length 2. You can omit the third argument if you want to extract a substring starting in the middle of a string, until the end of the string.
            
            varQuestion = Mid(textRow, 4)
            intContRow = intContRow + 1
       End If
       
       If StrComp(strKey, "a:") = 0 Then
            varAnswer = Mid(textRow, 4)
            intContRow = intContRow + 1
       End If
       
       If 5 > Len(Trim(textRow)) Then
            intContRow = intContRow + 1
           
       End If
    
       
       ' Save data into collection when the answer and the question are captured
       If (intContRow Mod 3 = 0) Then  '>
            
            ' Create an Object of CQuestion class type
            Set qstQuestion = New CQuestion
            ' set data captured from .txt file
            qstQuestion.Question = varQuestion
            qstQuestion.Answer = varAnswer
            
            ' create a dynamic numeration as Id
            intEmpNum = intEmpNum + 1
            Id = "E" & Format$(intEmpNum, "00000")
            ' Add the qstQuestion Object to the Collection plus its Id
            qstCllSortedQuestions.Add qstQuestion, Id
            
       End If
            
    Loop
    Close #fileNo
    
    'Dim obj As CQuestion
    'Set obj = qstCllSortedQuestions.Item(1)
    'Debug.Print qstCllSortedQuestions.Item(2).Answer
    
    'Dim MyObject As CQuestion
    'For Each MyObject In qstCllSortedQuestions
        
        'Debug.Print MyObject.Answer
        'Debug.Print MyObject.Question
        ' Exit For
    'Next

End Function

Public Function CreateHardCodeData() As Collection
    Dim qstCllSortedQuestions As New Collection
    qstQuestion.Question = "pregunta 1"
    qstQuestion.Answer = " respuesta 1"
    qstCllSortedQuestions.Add qstQuestion
    
    qstQuestion.Question = "pregunta 2"
    qstQuestion.Answer = "respuesta 2"
    'qstQuestion.Answer = qstCllQuestions(0).Answer
    qstCllSortedQuestions.Add qstQuestion
    
    Set CreateHardCodeData = qstCllSortedQuestions
    Debug.Print "Created from class " + LTrim(Str(qstCllSortedQuestions.Count))
End Function


Public Function dummy()
    Dim fileName As String, textData As String, textRow As String, fileNo As Integer
    Dim strKey As String
    Dim intContRow As Integer
    fileName = "C:\Users\Kevin\Documents\Projects\EnglishPowerPointApp\testText.txt"
    fileNo = FreeFile 'Get first free file number
        
    intContRow = 0
    Open fileName For Input As #fileNo
    Do While Not EOF(fileNo)
       Line Input #fileNo, textRow
       textData = textData & textRow
         
         strKey = Left(textRow, 2)
       If StrComp(strKey, "q:") = 0 Then
            ' Mid
            ' To extract a substring, starting in the middle of a string, use Mid.
            ' Mid("example text", 9, 2) =  "te"
            ' started at position 9 (t) with length 2. You can omit the third argument if you want to extract a substring starting in the middle of a string, until the end of the string.
             'Debug.Print "Q:"
              qstQuestion.Question = Mid(textRow, 4)
              intContRow = intContRow + 1
       End If
       If StrComp(strKey, "a:") = 0 Then
            'Debug.Print "a:"
            qstQuestion.Answer = Mid(textRow, 4)
            intContRow = intContRow + 1
           
       End If
       
       If 5 > Len(Trim(textRow)) Then
            
            intContRow = intContRow + 1
           
       End If
    
       
       'Debug.Print "MOD: " + LTrim(Str(intContRow Mod 2))
       
       If (intContRow Mod 3 = 0) Then  '>
            'Debug.Print "saved in count " + LTrim(Str(intContRow))
            qstCllSortedQuestions.Add qstQuestion
       End If
       
     
     
       
    Loop
    Close #fileNo

End Function

Public Function CreateDisortedQuizCollection()
    
    Dim qstCllSortedQuestionsCopy As Collection
    Set qstCllSortedQuestionsCopy = qstCllSortedQuestions
    
    Set qstCllUnsortedQuestion = New Collection
  
    Dim qstQuestion As CQuestion
    Dim intRandomValue As Integer
    Dim intUpperBound As Integer, intLowerBound As Integer
  
    Do While Not (qstCllSortedQuestionsCopy.Count = 0)
        
        intUpperBound = qstCllSortedQuestionsCopy.Count
        intLowerBound = 1
        
        intRandomValue = CInt(Int((intUpperBound * Rnd()) + intLowerBound))
        
        qstCllUnsortedQuestion.Add qstCllSortedQuestionsCopy.Item(intRandomValue)
        qstCllSortedQuestionsCopy.Remove (intRandomValue)
        
    Loop
    
    Dim MyObject As CQuestion
    For Each MyObject In qstCllUnsortedQuestion
        
        Debug.Print MyObject.Answer
        Debug.Print MyObject.Question
        ' Exit For
    Next

End Function


Public Function Item(ByVal Index As Variant) _
As CQuestion
    Debug.Print qstCllSortedQuestions.Item(Index).Answer
   Set Item = qstCllSortedQuestions.Item(Index)
End Function









