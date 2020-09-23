VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Class Module Tutorial"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Out Come"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   4335
      Begin VB.Label lblResult 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Number of Replacements"
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Result"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtReplacement 
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtChar 
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtString 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Replacement"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Character"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter String"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program:  ClassModuleTutorial
'Description:

'This program allows the user to replace a
'chosen character with another character in
'a given string.

'This project is designed to be tutorial for
'implementing a class module.  I wrote this in
'order to learn more about modules.  I used
'character replacement as the task since it
'may be of use after the project is entered.
'I hope that it will be of assistance to others.  If
'you have any suggestions, please feel free
'to contact me at jerry_m_barnes@hotmail.com.

Option Explicit
Dim WithEvents ReplacementString As ReplaceChar
'If you want to use events you must declare
'the object using the WithEvents declaration.
'If you do not want to use events, it is not
'necessary.

Private Sub cmdClear_Click()
    
    Set ReplacementString = Nothing
    'Destroy the object so resources
    'are not wasted.
    
    lblResult.Caption = ""
    lblCount.Caption = ""
    txtChar.Text = ""
    txtReplacement.Text = ""
    txtString.Text = ""
    'Clear the controls.
    
    txtString.SetFocus
    'return to the first text box.
End Sub

Private Sub cmdExit_Click()
    
    Set ReplacementString = Nothing
    'Tidy up.  Don't waste resources.
    
    End
End Sub

Private Sub cmdReplace_Click()

    Set ReplacementString = New ReplaceChar
    'Create a new object of the class that
    'was created.
    
    ReplacementString.ToBeReplaced = txtChar.Text
    'Send the property ToBeReplaced.  This
    'is a Let sub in the module.
    
    ReplacementString.ReplaceWith = txtReplacement.Text
    'Send the property ReplaceWith.  This
    'is a Let sub in the module.
    
    lblResult.Caption = ReplacementString.ReplaceChar(txtString.Text)
    'Set the caption of lblResult with the
    'results of the Replace method.
    
    lblCount.Caption = ReplacementString.Count
    'Get the count through the count property.
    'This is a Get sub in the module.
End Sub

Private Sub Replacementstring_NoSubstitute(strString As String)
    'This subs only purpose is to demonstrate
    'using an event.  StrString is passed
    'from the module back to the program.
    
    MsgBox "No substitutions were made in " & strString, vbOKOnly, "Warning"
End Sub
