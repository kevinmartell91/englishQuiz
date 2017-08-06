Option Explicit

Private mstrQuestion As String
Private mstrAnswer As String

''''''''''''''''''''''
' Question property
''''''''''''''''''''''
Public Property Get Question() As String
    Question = mstrQuestion
End Property
Public Property Let Question(value As String)
    mstrQuestion = value
End Property

''''''''''''''''''''''
' Answer property
''''''''''''''''''''''
Public Property Get Answer() As String
    Answer = mstrAnswer
End Property
Public Property Let Answer(value As String)
    mstrAnswer = value
End Property
