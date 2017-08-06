Private qstCllSortedQuestions As Collection
Private qstCllUnsortedQuestion As Collection
Private qstCllSortedQuestionsCopy As Collection
Private qstQuestion As CQuestion
Dim blnValidData As Boolean
Public strPath As String




Private Sub Class_Initialize()
    Set qstCllSortedQuestions = New Collection
    Set qstCllUnsortedQuestion = New Collection
    Set qstCllSortedQuestionsCopy = New Collection
    Set qstQuestion = New CQuestion
    blnValidData = False
    
    
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
''''''''''''''''''''''
' Path property
''''''''''''''''''''''
Public Property Let Path(value As String)
    strPath = value
End Property



Public Function LoadDataFromTXT()

    Dim fileName As String, textData As String, textRow As String, fileNo As Integer
    Dim strKey As String
    Dim intContRow As Integer, intEmpNum As Integer
    Dim Id As String
    Dim varAnswer As Variant, varQuestion As Variant
    
    'fileName = "C:\Users\Kevin\Documents\Projects\EnglishPowerPointApp\testText.txt"
    fileName = strPath
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
            ' Add to copy collection
            qstCllSortedQuestionsCopy.Add qstQuestion, Id
            
       End If
            
    Loop
    Close #fileNo
    
    
    
    'Dim obj As CQuestion
    'Set obj = qstCllSortedQuestions.Item(1)
    'Debug.Print qstCllSortedQuestions.Item(2).Answer
    
    Dim MyObject As CQuestion
    For Each MyObject In qstCllSortedQuestions
        
        Debug.Print MyObject.Answer
        Debug.Print MyObject.Question
        ' Exit For
    Next
    
    Call CreateDisortedQuizCollection

End Function


Public Function CreateDisortedQuizCollection()
    
    Dim qstQuestion As CQuestion
    Dim intRandomValue As Integer
    Dim intUpperBound As Integer, intLowerBound As Integer
  
    Do While Not (qstCllSortedQuestionsCopy.Count = 0)
        Debug.Print qstCllSortedQuestions.Count & "from till in DB"
        
        intUpperBound = qstCllSortedQuestionsCopy.Count
        intLowerBound = 1
        
        intRandomValue = CInt(Int((intUpperBound * Rnd()) + intLowerBound))
        
        qstCllUnsortedQuestion.Add qstCllSortedQuestionsCopy.Item(intRandomValue)
        qstCllSortedQuestionsCopy.Remove (intRandomValue)
        
    Loop
    
    'Dim MyObject As CQuestion
    'For Each MyObject In qstCllUnsortedQuestion
        
        'Debug.Print MyObject.Answer
        'Debug.Print MyObject.Question
    'Next
    
    Debug.Print qstCllSortedQuestions.Count & "from till in DB"

End Function


Public Function Item(ByVal Index As Variant) _
As CQuestion
    Debug.Print qstCllSortedQuestions.Item(Index).Answer
   Set Item = qstCllSortedQuestions.Item(Index)
End Function







