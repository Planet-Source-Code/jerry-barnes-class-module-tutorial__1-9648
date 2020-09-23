<div align="center">

## Class Module Tutorial


</div>

### Description

This project is designed to be tutorial for implementing a class module. I wrote this in

order to learn more about modules. I used character replacement as the task since it

may be of use after the project is entered. I hope that it will be of assistance to others.

This program allows the user to replace a chosen character with another character in

a given string.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-07-10 14:34:26
**By**             |[Jerry Barnes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jerry-barnes.md)
**Level**          |Beginner
**User Rating**    |4.4 (53 globes from 12 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD78647172000\.zip](https://github.com/Planet-Source-Code/jerry-barnes-class-module-tutorial__1-9648/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Class Module Tutorial</title>
</head>
<body>
<p align="center"><font size="4"><b>Class Module Tutorial</b></font></p>
<p align="center"><font size="4"><b>for Beginners</b></font></p>
<p align="left">This project is designed to be tutorial for implementing a class module. I wrote this in<br>
order to learn more about modules. I used character replacement as the task since it<br>
may be of use after the project is entered.  I hope that it will be of assistance to others.  </p>
<p align="left">I'll start by giving a brief explanation of what a class module
is.  When you create a class module, you are basically creating an
object.  This object has properties, methods, and events like the controls
that you put on form.  For example, Caption is a property of a label, Clear
is a method of a listbox, and Click is an event for a command button. </p>
<p align="left">
This class module allows the user to replace a chosen character with another character in<br>
a given string.  It has properties, a method, and one event.</p>
<p align="left">It should take the user between 30 minutes and 45 minutes to
complete this project.<br>
<br>
<br>
Steps.</p>
<p align="left">1.  Open Visual Basic and select a standard EXE project.</p>
<p align="left">2.  Rename the form frmMain and save the project to
whatever name you like.</p>
<p align="left">3.  Add the following controls to the form.</p>
<table border="1" width="100%">
 <tr>
 <td width="33%">Label1</td>
 <td width="33%">Caption</td>
 <td width="34%">Enter String</td>
 </tr>
 <tr>
 <td width="33%">Label2</td>
 <td width="33%">Caption</td>
 <td width="34%">Enter Character</td>
 </tr>
 <tr>
 <td width="33%">Label3</td>
 <td width="33%">Caption</td>
 <td width="34%">Enter Replacement</td>
 </tr>
 <tr>
 <td width="33%">txtString</td>
 <td width="33%">Text</td>
 <td width="34%">""</td>
 </tr>
 <tr>
 <td width="33%">txtChar</td>
 <td width="33%">Text</td>
 <td width="34%">""</td>
 </tr>
 <tr>
 <td width="33%"> </td>
 <td width="33%">Maxlength</td>
 <td width="34%">1</td>
 </tr>
 <tr>
 <td width="33%">txtReplacement</td>
 <td width="33%">Text</td>
 <td width="34%">""</td>
 </tr>
 <tr>
 <td width="33%"> </td>
 <td width="33%">Maxlength</td>
 <td width="34%">1</td>
 </tr>
 <tr>
 <td width="33%">Frame1</td>
 <td width="33%">Caption</td>
 <td width="34%">Out Come</td>
 </tr>
 <tr>
 <td width="33%">Label4</td>
 <td width="33%">Caption</td>
 <td width="34%">Result</td>
 </tr>
 <tr>
 <td width="33%">Label5</td>
 <td width="33%">Caption</td>
 <td width="34%">Number of Replacements</td>
 </tr>
 <tr>
 <td width="33%">lblResult</td>
 <td width="33%">Caption</td>
 <td width="34%">""</td>
 </tr>
 <tr>
 <td width="33%"> </td>
 <td width="33%">BorderStyle</td>
 <td width="34%">1-Fixed Single</td>
 </tr>
 <tr>
 <td width="33%">lblCount</td>
 <td width="33%">Caption</td>
 <td width="34%">""</td>
 </tr>
 <tr>
 <td width="33%"> </td>
 <td width="33%">BorderStyle</td>
 <td width="34%">1-FixedSingle</td>
 </tr>
 <tr>
 <td width="33%">cmdReplace</td>
 <td width="33%">Caption</td>
 <td width="34%">Replace</td>
 </tr>
 <tr>
 <td width="33%">cmdClear</td>
 <td width="33%">Caption</td>
 <td width="34%">Clear</td>
 </tr>
 <tr>
 <td width="33%">cmdExit</td>
 <td width="33%">Caption</td>
 <td width="34%">Exit</td>
 </tr>
</table>
<p align="center"><img border="0" src="http://www.geocities.com/jerry_m_barnes/images/cmt01.jpg" width="335" height="315"></p>
<p align="center">The form should be similar to this when you are finished.</p>
<p align="left">4.  Right Click on Project1 in the Project window. 
Select Add from the menu.  Select Class Module.  Select Class Module
again.</p>
<p align="left">5.  Right Click on the Class Module in the Project
Window.  Change the name property to ReplaceChar.  This will be the
name of the object.</p>
<p align="left">6.  Declare the following variables and events.</p>
<blockquote>
 <p align="left">Option Explicit<br>
 <br>
 Private mToBeReplaced As String * 1<br>
 <br>
 Private mReplaceWith As String * 1<br>
 <br>
 Private mCount As Integer<br>
 <br>
 Public Event NoSubstitute(strString As String)</p>
 <p align="left"><i>Notice that the variables are private and the the event is
 public.  The variables actually hold values for the properties. 
 Since they are private, the program itself cannot manipulate them.  Only
 the module can change them.  Two of the strings are limited to 1
 character in length.</i></p>
</blockquote>
<p align="left">7.  Go to the Tool menu and select Add Procedure. 
Type the name of the property (ToBeReplaced) and select property option.  The scope
should be public for this property. Click OK. This will create two subs. One to
send data to the main project (Get) and one to receive data (Let).  You
will have to change the parameters to the variable types listed below.</p>
<p align="left">8.  Enter the following code for the two properties. 
The ToBeReplaced property hold the value of the character that will be replaced.</p>
<blockquote>
 <p align="left">Public Property Get ToBeReplaced() As String<br>
     ToBeReplaced = mToBeReplaced<br>
 End Property</p>
 <p align="left">
 <i>Get is used to send information from the object to the program.  The
 program is getting information.  Notice, the properties equal the
 variable declared in the declartions section.</i></p>
 <p align="left">Public Property Let ToBeReplaced(ByVal strChoice As String)<br>
     mToBeReplaced = strChoice<br>
 End Property                  </p>
 <p align="left"><i>Let is used to retrieve value from the program.  The
 program lets the module have information. </i></p>
</blockquote>
<p align="left">9.  Repeat the above the process for the ReplaceWith
Property.  The ReplaceWith property holds the value to replace the desired
character with.</p>
<blockquote>
 <p align="left">Public Property Get ReplaceWith() As String<br>
     ReplaceWith = mReplaceWith<br>
 End Property<br>
 <br>
 Public Property Let ReplaceWith(ByVal strChoice As String)<br>
     mReplaceWith = strChoice<br>
 End Property</p>
</blockquote>
<p align="left">10.  Finally, add the Count Property.  It will be read
only so it does not have a let property.  The count property will return to
the program the number of substitutions made.</p>
<blockquote>
 <p align="left">Public Property Get Count() As Integer<br>
     Count = mCount<br>
 End Property</p>
</blockquote>
<p align="left">11.  Now, we are going to add a method to the class
module.  Methods can consist of funtions or procedures.  This method
scans the string and makes the replacements.  It also raises an
event.  Look toward the bottom of the code.  If no replacements are
made, an event is raised.  This will be used in the form's code. 
Enter the following code.</p>
<blockquote>
 <p align="left">Public Function ReplaceChar(strString As String) As String<br>
     Dim intLoop As Integer<br>
     Dim intLen As Integer<br>
 <br>
     Dim strTemp As String<br>
     Dim strTest As String<br>
     Dim strHold As String<br>
 <br>
     mCount = 0<br>
     <font color="#008000">'The replacement count should be zero.</font><br>
 <br>
 <font color="#008000">    '#######################################<br>
     '# The following code scans the string                   
 #<br>
     '# and makes the desired replacements.                
 #<br>
     '#######################################</font><br>
     intLoop = 1<br>
     strTemp = ""<br>
     strHold = strString<br>
     intLen = Len(strString) + 1<br>
     Do Until intLoop = intLen<br>
         intLoop = intLoop + 1<br>
         strTest = Left(strHold, 1)<br>
         If strTest = mToBeReplaced Then<br>
             <font color="#008000">'mTobeReplaced comes
 from the properties.</font><br>
             strTemp = strTemp & mReplaceWith<br>
             <font color="#008000">'mReplaceWith comes from
 the properties.</font><br>
             mCount = mCount + 1<br>
         Else<br>
             strTemp = strTemp & Left(strHold, 1)<br>
         End If<br>
         strHold = Right(strHold, Len(strHold) - 1)<br>
     Loop<br>
 <font color="#008000">    '#######################################<br>
     '# Scanning and replacement code ends.               
 #<br>
     '#######################################</font><br>
 <br>
     If mCount <> 0 Then<br>
         ReplaceChar = strTemp<br>
         'Write the new string.<br>
     Else<br>
         RaiseEvent NoSubstitute(strTemp)<br>
     End If<br>
 <font color="#008000">    'If mCount is zero the no replacements<br>
     'were made. This means that we want to<br>
     'raise the event NoSubstitute.</font><br>
 <br>
 End Function</p>
</blockquote>
<p align="left">12.  Provide everything was entered correctly, the class
module is fully functional now.  Save it and go back to the form.</p>
<p align="left">13.  Enter the following declaration.  This declares a
variable as a type of the created object.</p>
<blockquote>
 <p align="left">Option Explicit<br>
 Dim WithEvents ReplacementString As ReplaceChar</p>
 <p align="left"><i>Note that WithEvents is not required.  However, it is
 necessary if you want to use events.</i></p>
</blockquote>
<p align="left">14.  Enter the code for the cmdReplace_Click Event. 
You have to create a new instance of the object first.  Next, set the
properties ToBeReplaced and ReplaceWith.  Next call the ReplaceChar
method.  Finally use the Count property to get the number of replacements.</p>
<blockquote>
 <p align="left">Private Sub cmdReplace_Click()<br>
 <br>
     Set ReplacementString = New ReplaceChar<br>
 <font color="#008000">    'Create a new object of the class that<br>
     'was created.</font><br>
 <br>
     ReplacementString.ToBeReplaced = txtChar.Text<br>
 <font color="#008000">    'Send the property ToBeReplaced. This<br>
     'is a Let sub in the module.</font><br>
 <br>
     ReplacementString.ReplaceWith = txtReplacement.Text<br>
 <font color="#008000">    'Send the property ReplaceWith. This<br>
     'is a Let sub in the module.<br>
 </font><br>
     lblResult.Caption = ReplacementString.ReplaceChar(txtString.Text)<br>
 <font color="#008000">    'Set the caption of lblResult with the<br>
     'results of the Replace method.</font><br>
 <br>
     lblCount.Caption = ReplacementString.Count<br>
 <font color="#008000">    'Get the count through the count property.<br>
     'This is a Get sub in the module.</font><br>
 End Sub<br>
 </p>
</blockquote>
<p align="left">15.  Program the event procedure for the class
module.  The event fires if no replacements were made.  You can code
whatever actions want to transpire when the event happens.  I used a
message box to alert the user that no changes were made.</p>
<blockquote>
 <p align="left">Private Sub Replacementstring_NoSubstitute(strString As String)<br>
 <font color="#008000"> 'This subs only purpose is to demonstrate using an event. StrString is passed<br>
 'from the module back to the program.</font><br>
 <br>
     MsgBox "No substitutions were made in " & strString, vbOKOnly, "Warning"<br>
 End Sub</p>
</blockquote>
<p align="left">16.  Enter code for the final two command buttons.</p>
<blockquote>
 <p align="left">Private Sub cmdClear_Click()<br>
 <br>
     Set ReplacementString = Nothing<br>
    <font color="#008000"> 'Destroy the object so resources<br>
     'are not wasted.</font><br>
 <br>
     lblResult.Caption = ""<br>
     lblCount.Caption = ""<br>
     txtChar.Text = ""<br>
     txtReplacement.Text = ""<br>
     txtString.Text = ""<br>
 <font color="#008000">   </font> <font color="#008000">'Clear the controls.</font><br>
 <br>
     txtString.SetFocus<br>
 <font color="#008000">   </font> <font color="#008000">'Return to the first text box.</font><br>
 End Sub<br>
 <br>
 Private Sub cmdExit_Click()<br>
 <br>
     Set ReplacementString = Nothing<br>
     '<font color="#008000">Tidy up. Don't waste resources.</font><br>
 <br>
     End<br>
 End Sub</p>
</blockquote>
<p align="left">17.  That's it.  The program should run.  The
module can be inserted in other programs now.  It does not have to be used
with text box or labels.  It can be used purely in code.  For example.</p>
<blockquote>
 <p align="left">Dim WithEvents RepStr As ReplaceChar</p>
 <p align="left">Set RepStr = New ReplaceChar</p>
 <p align="left">RepStr.ToBeReplace = " "</p>
 <p align="left">RepStr.ReplaceWith = "_"</p>
 <p align="left">strString = RepStr.ReplaceChar(strString)</p>
 <p align="left">if RepStr.Count = 0 then </p>
 <p align="left">    msgbox "No subs made"</p>
 <p align="left">End if</p>
</blockquote>
<p align="left">This would replace all space in a string with an
underscore.  Pretty useful.</p>
<p align="left"> </p>
<p align="left"> If you have any suggestions, please feel free to contact me at
jerry_m_barnes@hotmail.com.</p>
</body>
</html>

