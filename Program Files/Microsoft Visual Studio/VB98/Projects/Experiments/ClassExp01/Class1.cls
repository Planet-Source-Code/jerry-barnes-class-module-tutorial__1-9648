VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplaceChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The variables that are used in the properties
'are module level only not public.  Their
'values are manipulated through the public
'properties and method.  Please excuse the
'naming convention.  I thought it would be
'clearer to use m instead of mstr or mint.

Private mToBeReplaced As String * 1
    'This variable holds the character to be
    'replaced.  It's length is One character.
    
Private mReplaceWith As String * 1
    'This variable holds the replacement
    'character.  Its length is also one.
    
Private mCount As Integer
    'This variable represents the number
    'of substitutions made in a string.
    
Public Event NoSubstitute(strString As String)
    'This module will have one event,
    'named NoSubstitute.  If will fire
    'when no susbtitution is made.
    
'###########################################
'#                Properties               #
'###########################################

'Go to the Tool menu and select Add Procedure.
'Type the name and select property.  The scope
'should be public for this project.  Click
'OK.  This will create two subs.  One to
'send data to the main project (Get) and one
'to receive data (Let).

Public Property Get ToBeReplaced() As String
    'Get Property:  The program is fetching
    '               a value from the module.
    '               This could be useful in
    '               determining the last
    '               character that was
    '               replaced (maybe for
    '               and undo action).
    
    ToBeReplaced = mToBeReplaced
End Property

Public Property Let ToBeReplaced(ByVal strChoice As String)
    'Let Property:  The module is receiving
    '               a value.  In this case
    '               the value is the character
    '               to be replaced.
    
    mToBeReplaced = strChoice
End Property

Public Property Get ReplaceWith() As String
    'Get Property:  The program is fetching
    '               a value from the module.
    '               This could be useful in
    '               determining the last
    '               character that was
    '               used to replace another
    '               character (maybe for
    '               and undo action).
            
    ReplaceWith = mReplaceWith
End Property

Public Property Let ReplaceWith(ByVal strChoice As String)
    'Let Property:  The module is receiving
    '               a value.  In this case
    '               the module is receiving
    '               the value to substitute
    '               into the string.
    
    mReplaceWith = strChoice
End Property

Public Property Get Count() As Integer
    'Get Property:  The program is fetching
    '               a value from the module.
    '               In this case, the value
    '               fetched will be the
    '               number of replacements.
    '
    '               Note:  This is a read-
    '               only property since
    '               there is no let property.

    Count = mCount
End Property

'###########################################
'#                Methods                  #
'###########################################

'This module only has one method.  You can
'use the Tools menu to declare it.  I typed
'it in.

Public Function ReplaceChar(strString As String) As String
    Dim intLoop As Integer
    Dim intLen As Integer
    
    Dim strTemp As String
    Dim strTest As String
    Dim strHold As String
    
    mCount = 0
    'The replacement count should be zero.
    
    '#######################################
    '# The following code scans the string #
    '# and makes the desired replacements. #
    '#######################################
    intLoop = 1
    strTemp = ""
    strHold = strString
    intLen = Len(strString) + 1
    Do Until intLoop = intLen
        intLoop = intLoop + 1
        strTest = Left(strHold, 1)
        If strTest = mToBeReplaced Then
                'mTobeReplaced comes from
                'the properties.
            strTemp = strTemp & mReplaceWith
                'mReplaceWith comes from
                'the properties.
            mCount = mCount + 1
        Else
            strTemp = strTemp & Left(strHold, 1)
        End If
        strHold = Right(strHold, Len(strHold) - 1)
    Loop
    '#######################################
    '# Scanning and replacement code ends. #
    '#######################################
    
    If mCount <> 0 Then
        ReplaceChar = strTemp
        'Write the new string.
    Else
        RaiseEvent NoSubstitute(strTemp)
    End If
    'If mCount is zero the no replacements
    'were made.  This means that we want to
    'raise the event NoSubstitute.
    
End Function



