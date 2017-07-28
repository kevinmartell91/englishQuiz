Option Explicit

' add quizType = 'question' or 'answer'
' add strStatusQuiz = 'started' or 'finished'
Dim blnQuestion As Boolean
Dim blnAnswerInMind As Boolean
Dim blnRetryQuiz As Boolean
Dim blnStatusQuiz As Boolean
Dim intCorrect As Integer
Dim intIncorrect As Integer
Dim intContQuestions As Integer
Dim intTotalQuestions As Integer
Dim strUserName As String
Dim strQuestion As String
Dim strAnswer As String
Public gIntCount As Integer

Dim qzQuiz As CQuiz
Dim qstQuestion As CQuestion
Dim lqstCllQuestions As Collection
Dim blnValidData As Boolean

Dim sCountDown As Shape
Dim sContCorrect As Shape
Dim sConIncorrect As Shape
Dim sContQuestions As Shape
Dim sQuestion As Shape
Dim sAnswer As Shape

Dim sTimeUp As Shape
Dim sOptNext As Shape
Dim sOptCorrect As Shape
Dim sOptIncorrect As Shape
Dim sOptCheckMyAnswer As Shape
Dim sOptRetryQuiz As Shape

Dim lCorrectColor As Long
Dim lIncorrectColor As Long
Dim lTransparentYellowColor As Long


Sub YourName()
    strUserName = InputBox(Prompt:="Type you name")
    MsgBox "Welcome " + strUserName, vbApplicationModal, "France trivia quiz"
End Sub

' ==============================================================================================================
' Slider 01
' ==============================================================================================================

' Methods that initialize the quiz

Sub ShortVariablesName()

    Set sCountDown = ActivePresentation.Slides(3).Shapes("tltCountdown")
    Set sContCorrect = ActivePresentation.Slides(3).Shapes("tltContCorrect")
    Set sConIncorrect = ActivePresentation.Slides(3).Shapes("tltContIncorrect")
    Set sContQuestions = ActivePresentation.Slides(3).Shapes("tltContQuestions")
    Set sQuestion = ActivePresentation.Slides(3).Shapes("tltQuestion")
    Set sAnswer = ActivePresentation.Slides(3).Shapes("tltAnswer")
    
    Set sTimeUp = ActivePresentation.Slides(3).Shapes("tltTimeUp")
    Set sOptNext = ActivePresentation.Slides(3).Shapes("optNext")
    Set sOptCorrect = ActivePresentation.Slides(3).Shapes("optCorrect")
    Set sOptIncorrect = ActivePresentation.Slides(3).Shapes("optIncorrect")
    Set sOptCheckMyAnswer = ActivePresentation.Slides(3).Shapes("optCheckMyAnswer")
    Set sOptRetryQuiz = ActivePresentation.Slides(3).Shapes("optRetryQuiz")
    
    lCorrectColor = RGB(0, 176, 80)
    lIncorrectColor = RGB(255, 0, 0)
    lTransparentYellowColor = RGB(254, 195, 6)

End Sub

Sub InitializeVarables()

    intCorrect = 0
    intIncorrect = 0
    blnQuestion = False
    strUserName = "Student"
    intContQuestions = 0
    intTotalQuestions = 0
    gIntCount = 0
    blnAnswerInMind = False
    blnRetryQuiz = False
    blnStatusQuiz = True
    strQuestion = ""
    strAnswer = ""
      
    sCountDown.TextFrame.TextRange.Text = "-"
    sContCorrect.TextFrame.TextRange.Text = "0"
    sConIncorrect.TextFrame.TextRange.Text = "0"
    sContQuestions.TextFrame.TextRange.Text = "-"
    sQuestion.TextFrame.TextRange.Text = ""
    sAnswer.TextFrame.TextRange.Text = ""
    
    Call UpdateColorAnswerBox(0)
      
    Set qzQuiz = New CQuiz
    Set lqstCllQuestions = New Collection
  
    Set lqstCllQuestions = qzQuiz.UnsortedQuestions
    'Set lqstCllQuestions = qzQuiz.SortedQuestions
    
    Call SetQuizNumberQuestions
         
        
End Sub

Sub StartQuiz()
    
    ' initialize values
    Call ShortVariablesName

    Call InitializeVarables
         
    ActivePresentation.SlideShowWindow.View.Next
    
End Sub


' ==============================================================================================================
' Slider 02
' ==============================================================================================================

' User choose type of the quiz to begin: quetions or answers

Sub setQuestion()
    blnQuestion = True
    
    ' got to the next silder
    ActivePresentation.SlideShowWindow.View.Next
    
    Call StartNewQuestion
End Sub

Sub setAnswer()
    blnQuestion = False
    
    ' got to the next silder
    ActivePresentation.SlideShowWindow.View.Next
    
    Call StartNewQuestion
End Sub

' ==============================================================================================================
' Slider 03'
' ==============================================================================================================

' Options to answer each question

Sub SetCorrect()
    intCorrect = intCorrect + 1
    sContCorrect.TextFrame.TextRange.Text = LTrim(Str(intCorrect))
    Call StartNewQuestion
End Sub

Sub SetIncorrect()
    intIncorrect = intIncorrect + 1
    sConIncorrect.TextFrame.TextRange.Text = LTrim(Str(intIncorrect))
    Call StartNewQuestion
End Sub


' Dynamic methods used to show a new question in the same slider

Sub SetQuizNumberQuestions()
    '   TODO REad from List
    intTotalQuestions = lqstCllQuestions.Count
    
End Sub

Sub StartNewQuestion()
    If blnStatusQuiz Then
        Call InizializeControlToBeging
        Call UpdateColorAnswerBox(0)
        Call SetUpNewQuestion
        Call StartCountDown
    End If
End Sub

Sub InizializeControlToBeging()
    
    ' first, ask if the user did not choose Retry Quiz option
    ' if he did , set to question number 0 to start the retry quiz
    If blnRetryQuiz Then
        intContQuestions = 0
    End If

    blnAnswerInMind = False
    sTimeUp.Visible = False
    sOptNext.Visible = False
    sOptCorrect.Visible = False
    sOptIncorrect.Visible = False
    sOptCheckMyAnswer.Visible = True
    sOptRetryQuiz.Visible = False

End Sub

Sub SetUpNewQuestion()
    ' if the game is in progress
    If intContQuestions < intTotalQuestions Then
        ' recall UpdateQuestion to show the next question
        Call UpdateQuestion
    Else
        ' to stop the counter during the RetryQuiz option
        blnStatusQuiz = False
        Call SetupFinishControls

    End If
End Sub

Sub IsTheLastQuestion()
    If intContQuestions = intTotalQuestions Then
        blnStatusQuiz = False
    Else
        blnStatusQuiz = True
    End If
End Sub



Sub UpdateQuestion()
  
    ' update the number of question
    intContQuestions = intContQuestions + 1
    sContQuestions.TextFrame.TextRange.Text = LTrim(Str(intContQuestions)) + "/" + LTrim(Str(intTotalQuestions))


    ' from class
    strQuestion = lqstCllQuestions.Item(intContQuestions).Question
    strAnswer = lqstCllQuestions.Item(intContQuestions).Answer
    
    ' if the type of quiz is Question then it fills the respective blank spaces
    If blnQuestion = True Then
        sQuestion.TextFrame.TextRange.Text = strQuestion
        sAnswer.TextFrame.TextRange.Text = ""
    Else
        sQuestion.TextFrame.TextRange.Text = ""
        sAnswer.TextFrame.TextRange.Text = strAnswer
    End If
End Sub

' Methods which include option actions

Sub CheckMyAnswer()
    ' It makes invisible control
    sOptCheckMyAnswer.Visible = False
    
    ' It makes visible CORRECT and INCORRECT options
    sOptCorrect.Visible = True
    sOptIncorrect.Visible = True

    blnAnswerInMind = True
    Call ShowAnswer(1)
End Sub

Sub NextQuestion()
    Call StartNewQuestion
End Sub

Sub RetryQuiz()
    blnRetryQuiz = True

    ' clear all values from previous quiz'
    Call InitializeVarables

    ' Goto slide 2 to repeat all the quiz
    SlideShowWindows(1).View.GotoSlide 2
End Sub


Sub ShowAnswer(intAnswerBoxState As Integer)
'Sub ShowAnswer()

    UpdateColorAnswerBox (intAnswerBoxState)
    ' if the type of quiz is Question, then it fills the respective blank spaces
    If blnQuestion = True Then
        sAnswer.TextFrame.TextRange.Text = strAnswer
    Else
        sQuestion.TextFrame.TextRange.Text = strQuestion
    End If
End Sub

' ==============================================================================================================
' CountDown method

Sub StartCountDown()
    Dim dteTime As Date
    Dim intCountDown As Integer

    dteTime = Now()
    intCountDown = 5

    dteTime = DateAdd("s", intCountDown, dteTime)

    Do Until dteTime <= Now()
    DoEvents

        If Not (blnAnswerInMind) Then
            ' run counter'
            sCountDown.TextFrame.TextRange.Text = Format((dteTime - Now()), "s")
        End If
    Loop

    ' When time is up and there is no answer in student's mind
    If Not (blnAnswerInMind) Then
        ' It increase in 1 the number of incorrect questions
        intIncorrect = intIncorrect + 1
        ' It update the number of question in the view by its name'
        sConIncorrect.TextFrame.TextRange.Text = LTrim(Str(intIncorrect))

        ' Setup fail controls
        Call SetUpFailControls
        ' Show answer
        Call ShowAnswer(2)
    End If

End Sub

Sub SetUpFailControls()
    ' Show time up message
    sTimeUp.Visible = True
    
    ' Show Next button if the game is still in progress, otherwise Retry Quiz button will be shown
    If intContQuestions < intTotalQuestions Then
        sOptNext.Visible = True
    Else
        sOptRetryQuiz.Visible = True
    End If
    
    sOptCheckMyAnswer.Visible = False
End Sub

Sub SetupFinishControls()
    ' It updates the visibility of the buttons in the end of the quiz
    blnAnswerInMind = True ' to stopt the counter in the las question
    sTimeUp.Visible = False
    sOptNext.Visible = False
    sOptCorrect.Visible = False
    sOptCheckMyAnswer.Visible = False
    sOptIncorrect.Visible = False
    sOptCheckMyAnswer.Visible = False
    sOptRetryQuiz.Visible = True
End Sub

' ==============================================================================================================
' help functions to retrive new question and answer data

Function getQuestion(idQuestion As Integer) As String
  
End Function

Function getAnswer(idAnswer As Integer) As String
    
End Function

' ==============================================================================================================
' help methods to show scores

Sub Results()
    MsgBox "You've got" & intCorrect & " out of " & intCorrect + intIncorrect, vbApplicationModal, "France trivia quiz"
End Sub

Sub ResultsGlobal()
    MsgBox "You've got" & gIntCount
End Sub

Sub UpdateColorAnswerBox(intStateAnswerColor As Integer)
        
    Dim x As Long
    
    Select Case intStateAnswerColor
      Case 0 ' clear all
         sQuestion.Fill.ForeColor.RGB = lTransparentYellowColor
         sAnswer.Fill.ForeColor.RGB = lTransparentYellowColor
      Case 1 ' Update to GREEN color as a correct answer or question
         If blnQuestion = True Then
            sAnswer.Fill.ForeColor.RGB = lCorrectColor
            'For x = 10 To 12 'sAnswer.TextFrame.TextRange.Runs.Count
             '   sAnswer.TextFrame.TextRange.Runs(x).Font.Color.RGB = lTransparentYellowColor
            'Next
         Else
            sQuestion.Fill.ForeColor.RGB = lCorrectColor
         End If
      Case 2 ' Update to RED color as an incorrect answer or question
         If blnQuestion = True Then
            sAnswer.Fill.ForeColor.RGB = lIncorrectColor
         Else
            sQuestion.Fill.ForeColor.RGB = lIncorrectColor
         End If
      Case Else
         MsgBox "Unknown Number"
   End Select

   Debug.Print "ReplaceColors"

End Sub

Sub ReplaceColorsSample()

    Dim lFindColor As Long
    Dim lReplaceColor As Long
    Dim oSl As Slide
    Dim oSh As Shape
    Dim x As Long
    
    lFindColor = RGB(255, 128, 128)
    lReplaceColor = RGB(128, 128, 255)
       
        For Each oSh In ActivePresentation.Slides(3).Shapes
            With oSh

                ' Fill
                If .Fill.ForeColor.RGB = lFindColor Then
                    .Fill.ForeColor.RGB = lReplaceColor
                End If

                ' Line
                If .Line.Visible Then
                    If .Line.ForeColor.RGB = lFindColor Then
                        .Line.ForeColor.RGB = lReplaceColor
                    End If
                End If

                ' Text
                If .HasTextFrame Then
                    If .TextFrame.HasText Then
                        For x = 1 To .TextFrame.TextRange.Runs.Count
                            If .TextFrame.TextRange.Runs(x).Font.Color.RGB = lFindColor Then
                                .TextFrame.TextRange.Runs(x).Font.Color.RGB = lReplaceColor
                            End If
                        Next
                    End If
                End If
            End With
        Next
End Sub





