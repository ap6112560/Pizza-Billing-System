VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "150"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   43
      Top             =   2520
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "80"
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   42
      Top             =   2520
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "250"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   41
      Top             =   3120
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "150"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   40
      Top             =   3120
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "170"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   39
      Top             =   3720
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "90"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   38
      Top             =   3720
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "250"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   37
      Top             =   4320
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "140"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   36
      Top             =   4320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rs"
      Height          =   615
      Index           =   4
      Left            =   4560
      TabIndex        =   35
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rs"
      Height          =   615
      Index           =   3
      Left            =   4560
      TabIndex        =   34
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rs"
      Height          =   615
      Index           =   2
      Left            =   4560
      TabIndex        =   33
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rs"
      Height          =   615
      Index           =   1
      Left            =   4560
      TabIndex        =   32
      Top             =   4080
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "180"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   31
      Top             =   1920
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "100"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   30
      Top             =   1920
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   8040
      TabIndex        =   29
      Top             =   4200
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      ItemData        =   "Form1.frx":0000
      Left            =   8040
      List            =   "Form1.frx":0010
      TabIndex        =   27
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Total Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   25
      Top             =   4680
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      ItemData        =   "Form1.frx":0044
      Left            =   8040
      List            =   "Form1.frx":0054
      TabIndex        =   20
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      ItemData        =   "Form1.frx":0088
      Left            =   8040
      List            =   "Form1.frx":0098
      TabIndex        =   19
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "Form1.frx":00CC
      Left            =   8040
      List            =   "Form1.frx":00DC
      TabIndex        =   18
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   6960
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   6960
      TabIndex        =   16
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   6960
      TabIndex        =   15
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   6960
      TabIndex        =   14
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   6960
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rs"
      Height          =   615
      Index           =   0
      Left            =   4560
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tomato Pizza"
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Chicken Pizza"
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Paneer Pizza"
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cheese Capsicum Pizza"
      Height          =   435
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cheese Pizza"
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label10 
      Height          =   495
      Index           =   1
      Left            =   10200
      TabIndex        =   28
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label15 
      Height          =   495
      Left            =   5880
      TabIndex        =   26
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label10 
      Height          =   375
      Index           =   3
      Left            =   10200
      TabIndex        =   24
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label10 
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   23
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label10 
      Height          =   375
      Index           =   4
      Left            =   10200
      TabIndex        =   22
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label10 
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   21
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Price per Serving"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Medium"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Large"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Servings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Toppings(Rs.20)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "PIZZAS"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s, i As Integer
Private Sub Form_Load()
s = 0
For i = 0 To 4
Option2(i).Enabled = False
Option1(i).Enabled = False
Text1(i).Enabled = False
Combo1(i).Enabled = False
Next i
End Sub
Private Sub Check1_Click(Index As Integer)
If Check1(Index).Value = vbChecked Then
Option1(Index).Enabled = True
Option2(Index).Enabled = True
If Option1(Index).Value = True Then
Text1(Index).Enabled = True
ElseIf Option2(Index).Value = True Then
Text1(Index).Enabled = True
End If
Else
Option1(Index).Enabled = False
Option2(Index).Enabled = False
Text1(Index).Text = ""
Text1(Index).Enabled = False
Label10(Index).Caption = ""
Combo1(Index).Enabled = False
End If
End Sub

Private Sub Check6_Click()
If (Check6.Value = vbChecked) Then
Label15.Caption = Val(Label10(0).Caption) + Val(Label10(1).Caption) + Val(Label10(2).Caption) + Val(Label10(3).Caption) + Val(Label10(4).Caption)
Else
Label15.Caption = ""
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
If Combo1(Index).ListIndex < 3 And s < 1 Then
Label10(Index).Caption = Val(Label10(Index).Caption) + 20 * Val(Text1(Index).Text)
s = s + 1
ElseIf Combo1(Index).ListIndex = 3 And s = 1 Then
s = s - 1
Label10(Index).Caption = Val(Label10(Index).Caption) - 20 * Val(Text1(Index).Text)
End If
End Sub

Private Sub Option1_Click(Index As Integer)
If (Text1(Index).Enabled = True) Then
Label10(Index).Caption = Val(Text1(Index).Text) * Val(Option1(Index).Caption)
If (Combo1(Index).ListIndex > -1 And Combo1(Index).ListIndex < 3) Then
Label10(Index).Caption = Val(Label10(Index).Caption) + 20 * Val(Text1(Index).Text)
End If
End If
If Option1(Index).Value Or Option2(Index).Value = vbChecked Then
Text1(Index).Enabled = True
End If
End Sub

Private Sub Option2_Click(Index As Integer)
If (Text1(Index).Enabled = True) Then
Label10(Index).Caption = Val(Text1(Index).Text) * Val(Option2(Index).Caption)
If (Combo1(Index).ListIndex > -1 And Combo1(Index).ListIndex < 3) Then
Label10(Index).Caption = Val(Label10(Index).Caption) + 20 * Val(Text1(Index).Text)
End If
End If
If Option2(Index).Value Or Option1(Index).Value = vbChecked Then
Text1(Index).Enabled = True
End If
End Sub

Private Sub Text1_Change(Index As Integer)
If Val(Text1(Index).Text) > 0 Then
Combo1(Index).Enabled = True
If Option1(Index).Value = True Then
Label10(Index).Caption = Val(Text1(Index).Text) * Val(Option1(Index).Caption)
Else
Label10(Index).Caption = Val(Text1(Index).Text) * Val(Option2(Index).Caption)
End If
If (Combo1(Index).ListIndex > -1 And Combo1(Index).ListIndex < 3) Then
Label10(Index).Caption = Val(Label10(Index).Caption) + 20 * Val(Text1(Index).Text)
End If
Else
Label10(Index).Caption = ""
Combo1(Index).Enabled = False
If (Combo1(Index).ListIndex > -1 And Combo1(Index).ListIndex < 3) Then
Label10(Index).Caption = Val(Label10(Index).Caption) + 20 * Val(Text1(Index).Text)
End If
End If
End Sub
