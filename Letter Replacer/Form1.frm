VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Letter replacer by Peter Hebels"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Open text file"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Randomize letters"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ouput:"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5895
      Begin VB.TextBox Text2 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter some text here:"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5895
      Begin VB.TextBox Text1 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Letter replacer written by Peter Hebels, website: http://www.phsoft.nl
'Thanks to Roger Gilchrist for sending me some nice code optimalizations!
'
'This code only replaces the middle letters of a word, the first and last letter
'will not be replaced. Also spaces, comma's and other special characters will not be
'replaced.
'
'Please note that the author of this code cannot be held responsible for any damages
'may caused by the use of this code, you use it at your own risk.

Option Explicit 'Never code without it :)

Private Sub Command1_Click()
Dim TheWord As String      'Here we put the characters to form a whole word.
Dim TheChar As String      'Single characters, used for detecting spaces and other characters like comma's.
Dim OutputText As String   'Here are all the characters put in, this is the final string and will be put in the textbox.
Dim InputText As String    'The text you enter will be put in here.
Dim MidLetters As String   'The letters in the middle of a word, without the first and the last character.
Dim SpecialChar As Boolean 'Do we have a special character like a comma or a space?
Dim x As Long              'Used for the main loop.

'If an error appears, then just resume:
On Error Resume Next

'Put the text you enterd into our variable:
InputText = Text1.Text
'If we don't add a space to the end of the string, the last word will not be written!
InputText = InputText & " "
                            
'Loop trough the entire text you entered:
For x = 1 To Len(InputText)
    'Get a single character from the text so it can be recognized:
    TheChar = Mid(InputText, x, 1)
    'Look for special characters in the string:
    SpecialChar = InStr(" ,.!?'" & Chr$(34) & vbCr & vbLf & vbTab, TheChar) > 0
    'If we have a special character like a space or comma, we can't just put it somewhere
    'in the middle of a word, it has to be placed at the end:
    If SpecialChar = True Then
        'Only single letter words need special treatment:
        If Len(TheWord) <= 1 Then
            'Look for special characters that have to be put back at the end of the string:
            OutputText = OutputText & TheWord & TheChar
        'If the word is longer, we are going to replace te middle characters:
        Else
            'Get the middle letters from the word:
            MidLetters = Mid$(TheWord, 2, Len(TheWord) - 2)
            'Put the first and last character back and add the randomized letters to the OutputText variable:
            OutputText = OutputText & Left$(TheWord, 1) & RandomizeLetters(MidLetters) & Right$(TheWord, 1) & TheChar
            'Empty the MidLetters variable.
            MidLetters = vbNullString
        End If
    'Reset the word string:
    TheWord = vbNullString
    'Reset the SpecialChar boolean to flase:
    SpecialChar = False
    Else
        'Add the characters back to the word string:
        TheWord = TheWord & TheChar
    End If
Next x
'Put the string back into our textbox:
Text2.Text = OutputText
On Error GoTo 0
End Sub

Function RandomizeLetters(InputWord As String) As String
'This is our randomize function, actually it's not really randomizing the characters,
'we only reverse the middle characters.

Dim AddLetter As String 'This is the variable where all the replaced letters are put into:
Dim SelChar As String 'A single character is put in here:
Dim I As Long 'For our reverse loop:
    
    'Read the characters in reverse:
    For I = Len(InputWord) + 1 To 1 Step -1
         'Get a single character from the string:
        SelChar = Mid(InputWord, I, 1)
        'Put the characters back into the string, in reverse of cource:
        AddLetter = AddLetter & SelChar
    Next I
    'Return the reversed string:
    RandomizeLetters = AddLetter
End Function

Private Sub Command2_Click()
Dim ToTextBox As String
Dim TextLine As String

'Stop on errors:
On Error GoTo ErrHandler
    'CD_File is a replacement module for the CommonDialog control:
    CD_File.FileName = ""
    CD_File.FileTitle = ""
    CD_File.hWndOwner = Form1.hwnd
    CD_File.DialogTitle = "Open test file"
    CD_File.CancelError = True
    CD_File.filter = "Text Files (*.txt*)|*.txt*"
    CD_File.ShowOpen
   
    'Open the file for input:
    Open CD_File.FileName For Input As #1
    'Loop until we are at the end of the file:
    Do While Not EOF(1)
       'Read the lines into the TextLine variable:
       Line Input #1, TextLine
       'Add the line and a newline to te textbox:
       ToTextBox = ToTextBox + vbCrLf + TextLine
    Loop
    'Remove any existing text from the textbox:
    Text1.Text = vbNullString
    'Put the text into the textbox.
    SetWindowText Text1.hwnd, Right(ToTextBox, Len(ToTextBox) - 2)
    'Empty the ToTextBox variable, used to free some memory:
    ToTextBox = vbNullString
    'Close the opened file:
    Close #1
'Our error handler:
ErrHandler:
'Make sure the file is always closed:
Close #1
End Sub

Private Sub Form_Load()
'Add some helping text to the label:
Label1.Caption = "Enter some text into the first textbox and click on the 'Randomize letters' button. You will notice that you are " & _
                 "still able to read the text in the second textbox, even if the letters are a total mess!"
    
    'Put a nice little text into the textbox:
    Text1.Text = "According to research at an english university, it doesn't matter in what order the letters in a word are, " & _
                 "the only important thing is that the first and last letter is at the right place. The rest can be a total " & _
                 "mess and you can still read it without problem. This is because we do not read every letter by it self but the word as a whole."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload the form before closing the program:
    Unload Me
End Sub
