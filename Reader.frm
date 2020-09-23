VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form Reader
BorderStyle     =   3 'Fixed Dialog
Caption         =   "TAS Site Words And Beyond"
ClientHeight    =   1830
ClientLeft      =   45
ClientTop       =   435
ClientWidth     =   10800
Icon            =   "Reader.frx":0000
LinkTopic       =   "Form1"
LockControls    =   -1 'True
MaxButton       =   0 'False
MinButton       =   0 'False
ScaleHeight     =   1830
ScaleWidth      =   10800
StartUpPosition =   2 'CenterScreen
Begin VB.CommandButton Command2
Caption         =   "Load List"
Height          =   270
Left            =   3645
TabIndex        =   13
Top             =   60
Width           =   1050
End
Begin VB.ComboBox ComWordLists
Height          =   315
Left            =   45
TabIndex        =   12
Text            =   "ComWordLists"
Top             =   30
Width           =   3555
End
Begin VB.CheckBox OneChance
Caption         =   "One Chance"
Height          =   225
Left            =   45
TabIndex        =   11
Top             =   405
Width           =   1290
End
Begin VB.Frame Frame2
Caption         =   "Right"
Height          =   645
Left            =   9015
TabIndex        =   8
Top             =   45
Width           =   1095
Begin VB.Label LabRight
Alignment       =   2 'Center
BackStyle       =   0 'Transparent
Caption         =   "0"
Height          =   300
Left            =   45
TabIndex        =   10
Top             =   225
Width           =   960
End
End
Begin VB.Frame Frame1
Caption         =   "Wrong"
Height          =   645
Left            =   7650
TabIndex        =   7
Top             =   45
Width           =   1095
Begin VB.Label LabWrong
Alignment       =   2 'Center
BackStyle       =   0 'Transparent
Caption         =   "0"
Height          =   300
Left            =   60
TabIndex        =   9
Top             =   255
Width           =   960
End
End
Begin VB.ListBox WordList
Height          =   2205
ItemData        =   "Reader.frx":0442
Left            =   2220
List            =   "Reader.frx":0444
TabIndex        =   3
Top             =   2580
Width           =   4320
End
Begin HTTSLibCtl.TextToSpeech TTS1
Height          =   615
Left            =   240
OleObjectBlob   =   "Reader.frx":0446
TabIndex        =   2
Top             =   2340
Visible         =   0 'False
Width           =   945
End
Begin VB.CommandButton Command1
Caption         =   "Speak Word"
Height          =   465
Left            =   4335
TabIndex        =   0
TabStop         =   0 'False
Top             =   645
Width           =   2130
End
Begin VB.Label Words
Alignment       =   2 'Center
BackColor       =   &H00E0E0E0&
BorderStyle     =   1 'Fixed Single
Caption         =   "Windmill"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   13.5
Charset         =   0
Weight          =   700
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
Height          =   420
Index           =   3
Left            =   8108
TabIndex        =   6
Top             =   1305
Width           =   2310
End
Begin VB.Label Words
Alignment       =   2 'Center
BackColor       =   &H00E0E0E0&
BorderStyle     =   1 'Fixed Single
Caption         =   "Windmill"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   13.5
Charset         =   0
Weight          =   700
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
Height          =   420
Index           =   2
Left            =   5533
TabIndex        =   5
Top             =   1305
Width           =   2310
End
Begin VB.Label Words
Alignment       =   2 'Center
BackColor       =   &H00E0E0E0&
BorderStyle     =   1 'Fixed Single
Caption         =   "Windmill"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   13.5
Charset         =   0
Weight          =   700
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
Height          =   420
Index           =   1
Left            =   2958
TabIndex        =   4
Top             =   1305
Width           =   2310
End
Begin VB.Label Words
Alignment       =   2 'Center
BackColor       =   &H00E0E0E0&
BorderStyle     =   1 'Fixed Single
Caption         =   "Windmill"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   13.5
Charset         =   0
Weight          =   700
Underline       =   0 'False
Italic          =   0 'False
Strikethrough   =   0 'False
EndProperty
Height          =   420
Index           =   0
Left            =   390
TabIndex        =   1
Top             =   1305
Width           =   2310
End
End
Attribute VB_Name = "Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private WordKey As Integer
Private GotRight As Boolean
Private LastRnd As Integer
Private Sub SetTest()
    Dim X As Integer
    Dim WordList() As String
    GotRight = True
    WordList() = Split(GenWords, vbCrLf)
    For X = 0 To 3
        Words(X).Caption = StrConv(WordList(X), vbProperCase)
    Next X
    Randomize
    Do Until WordKey <> LastRnd
        WordKey = Left(Int(4 * Rnd + 1), 1) - 1
        DoEvents
    Loop
    LastRnd = WordKey
    TTS1.Speak LCase(Words(WordKey).Caption) & " ."
End Sub
Private Function GenWords() As String
    'Idea here is to generate the 4 words used for each test
    Dim X As Integer
    Dim NewWord As String
    For X = 1 To 4
        NewWord = WordList.List(Shuffle(0, WordList.ListCount - 1))
        If InStr(1, GenWords, NewWord, vbTextCompare) = 0 Then
            GenWords = GenWords & NewWord & vbCrLf
        Else
            X = X - 1
        End If
    Next X
End Function
'**************************************
'Windows API/Global Declarations for
':Extendid Random Number with Range
'**************************************
Static Function Shuffle(Lower As Integer, Upper As Integer) As Integer
Static PrimeFactor(10) As Integer
Static a As Integer, c As Integer, B As Integer
Static s As Long, n As Integer
Dim i As Integer, J As Integer, k As Integer, n1 As Integer
Dim m As Integer
Dim t As Boolean


If (n <> Upper - Lower + 1) Then
    n = Upper - Lower + 1
    i = 0
    n1 = n
    k = 2
    
    
    Do While k <= n1
        
        
        If (n1 Mod k = 0) Then
            
            
            If (i = 0 Or PrimeFactor(i) <> k) Then
                i = i + 1
                PrimeFactor(i) = k
            End If
            n1 = n1 / k
        Else
            k = k + 1
        End If
    Loop
    B = 1
    
    
    For J = 1 To i
        B = B * PrimeFactor(J)
    Next J
    If n Mod 4 = 0 Then B = B * 2
    a = B + 1
    c = Int(n * 0.66)
    t = True
    
    
    Do While t
        t = False
        
        
        For J = 1 To i
            If ((c Mod PrimeFactor(J) = 0) Or (c Mod a = 0)) Then t = True
        Next J
        If t Then c = c - 1
    Loop
    Randomize
    s = Rnd(n)
End If
s = (a * s + c) Mod n
Shuffle = s + Lower
End Function
Private Sub Command1_Click()
    WordList.SetFocus
    TTS1.Speak "The word is  " & LCase(Words(WordKey).Caption) & " ."
End Sub
Private Sub Command2_Click()
    WordList.SetFocus
    LoadWords
End Sub

Private Sub Form_Load()
    btnFlat Command1
    btnFlat Command2
    ComWordLists.AddItem "Dolch Pre-Primer"
    ComWordLists.AddItem "Dolch Primer"
    ComWordLists.AddItem "Dolch First Grade"
    ComWordLists.AddItem "Dolch Second Grade"
    ComWordLists.AddItem "Dolch Third Grade"
    ComWordLists.AddItem "Types Of Birds"
    ComWordLists.AddItem "Parts Of The Body"
    ComWordLists.AddItem "Zoo Animals"
    ComWordLists.AddItem "Types Of Insects"
    ComWordLists.AddItem "Types Of Reptiles"
    ComWordLists.AddItem "Types Of Fruit"
    ComWordLists.AddItem "Vegetables"
    ComWordLists.AddItem "Weather Words"
    ComWordLists.AddItem "Winter Words"
    ComWordLists.AddItem "July 4Th Words"
    ComWordLists.AddItem "Animal Babys"
    ComWordLists.AddItem "Types Of Animals"
    ComWordLists.ListIndex = 0
    ComWordLists.ListIndex = GetSetting("TAS Site Words And Beyond", "Settings", "Word List", ComWordLists.ListIndex)
    LoadWords
End Sub
Private Sub LoadWords()
    Dim FName As String
    Select Case ComWordLists.ListIndex
    Case 0
        FName = "DolchPrePrimer.dat"
    Case 1
        FName = "DolchPrimer.dat"
    Case 2
        FName = "DolchFirstGrade.dat"
    Case 3
        FName = "DolchSecondGrade.dat"
    Case 4
        FName = "DolchThirdGrade.dat"
    Case 5
        FName = "Birds.dat"
    Case 6
        FName = "BodyParts.dat"
    Case 7
        FName = "ZooAnimals.dat"
    Case 8
        FName = "Insects.dat"
    Case 9
        FName = "Reptiles.dat"
    Case 10
        FName = "TypesOfFruit.dat"
    Case 11
        FName = "Vegetables.dat"
    Case 12
        FName = "Weather.dat"
    Case 13
        FName = "Winter.dat"
    Case 14
        FName = "July4Th.dat"
    Case 15
        FName = "AnimalBabys.dat"
    Case 16
        FName = "Animals.dat"
        
    End Select
    WordList.Clear
    Call LoadListFromFile(App.Path & "\" & FName, WordList)
    LabRight.Caption = 0
    LabWrong.Caption = 0
    SetTest
End Sub
Private Sub SetScore()
    If GotRight = True Then
        LabRight.Caption = LabRight.Caption + 1
    Else
        LabWrong.Caption = LabWrong.Caption + 1
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "TAS Site Words And Beyond", "Settings", "Word List", ComWordLists.ListIndex
    Unload Me
End Sub
Private Sub OneChance_Click()
    WordList.SetFocus
End Sub
Private Sub Words_Click(Index As Integer)
    TTS1.Speak Words(Index).Caption
    If Index = WordKey Then
        SetScore
        TTS1.Speak "Great job ! The next word is  "
        SetTest
    Else
        If OneChance.Value = 1 Then
            SetScore
            TTS1.Speak "Sorry that's the wrong word !"
            SetTest
        Else
            GotRight = False
            TTS1.Speak "Sorry that's the wrong word ! Try again ! The word is  " & LCase(Words(WordKey).Caption) & " ."
        End If
    End If
End Sub
Private Function btnFlat(Button As CommandButton)
    SetWindowLong Button.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
