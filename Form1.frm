VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Recommanc�"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdTry 
      Caption         =   "Valider"
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtGuess 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "entrer votre Valeur"
      Top             =   240
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   6600
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label lblAttempts 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim secretNumber As Integer
Dim attempts As Integer
Const MaxNumber As Integer = 100

Private Sub InitSecretNumber()
    Randomize
    secretNumber = Int(Rnd * MaxNumber) + 1
End Sub


Private Sub ResetGame()
    InitSecretNumber
    attempts = 0
    lblMessage.Caption = "Devine un nombre entre 1 et " & MaxNumber
    Form1.Caption = "Devine un nombre entre 1 et " & MaxNumber
    lblAttempts.Caption = "Tentatives : 0"
    txtGuess.Text = ""
    cmdTry.Enabled = True
    txtGuess.Enabled = True
    cmdBegin.Visible = False
End Sub

Private Sub CheckGuess()
    Dim guess As Integer
    If IsNumeric(txtGuess.Text) Then
        guess = CInt(txtGuess.Text)
        attempts = attempts + 1
        lblAttempts.Caption = "Tentatives : " & attempts

        If guess < secretNumber Then
            lblMessage.Caption = guess & " est trop petit !"
        ElseIf guess > secretNumber Then
            lblMessage.Caption = guess & " est trop grand !"
        Else
            lblMessage.Caption = "Bravo ! Tu as trouv� en " & attempts & " tentatives !"
            cmdTry.Enabled = False
            txtGuess.Enabled = False
            cmdBegin.Visible = True
        End If
    Else
        lblMessage.Caption = "Entre un nombre valide !"
    End If
    txtGuess.Text = ""
End Sub

Private Sub cmdBegin_Click()
    Call ResetGame
End Sub

Private Sub cmdTry_Click()

    Call CheckGuess
End Sub

Private Sub Form_Load()
    Call ResetGame
        
End Sub


Private Sub txtGuess_Change()
    If txtGuess.Text <> "" Then
        If Val(txtGuess.Text) > MaxNumber Then
            MsgBox "Veuillez entrer un nombre entre 0 et 100.", vbExclamation
            txtGuess.Text = "100"
        End If
    End If
End Sub


Private Sub txtGuess_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call CheckGuess
    End If
End Sub

Private Sub txtGuess_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Bloque la touche
    End If
End Sub


