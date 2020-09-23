<div align="center">

## Doing Strings in VB Part 1


</div>

### Description

Strings play an important role in every software. This tutorial is big, it has everything that a beginner wants. This version is modified to remove some mistakes in the previous one, and adds some information about instr. I'd be glad to see your feedback. Thanks! Note: Since I am no guru, this could be all wrong, use it at your own risk.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TeknikForce](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/teknikforce.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/teknikforce-doing-strings-in-vb-part-1__1-24583/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type"
content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title>Doing Strings In VB</title>
</head>
<body bgcolor="#FFFFFF" link="#0000FF" vlink="#800080">
<p><font size="5" face="Verdana"><strong>Doing Strings In VB Part
1</strong></font><font size="2" face="Verdana"><br>
By Cyril &#145;Razoredge&#146; Gupta<br>
Mail: </font><a href="mailto:cyril@icnol.com"><font size="2" face="Verdana">cyril@icnol.com</font></a><font size="2"
face="Verdana"><br>
Warning: The code presented here is not indented properly because
HTML won't let me put a space or a tab character before the text.
Please indent the code if you plan to use the reuse the code in
your program.</font></p>
<p><font size="2" face="Verdana">Strings are an indispensable
part of almost all VB software; you will need to use strings in
almost all the software you ever make.</font></p>
<p><font size="2" face="Verdana"><b>Let&#146;s start with <br>
What is a string and where do you use it?</b><br>
In VB String is a length of text assigned to a variable of type
Variant or of type String. A string can store a maximum of around
2 billion characters between ASCII value 32 to 256. Strings mean
a lot to a programmer. They can hold important data, which the
user reads, intermediate values, comments, or can be used simply
to test if the software works correctly. People store text in
strings in .INI files, in the windows registry .RES files and
other text resources. </font></p>
<p><font size="2" face="Verdana"><b>Strings in a file<br>
</b>You may often need to store and retrieve text from a file.
Here&#146;s how<br>
Retrieving text from a file<br>
VB6 and VB5 introduced the new File object handling system but
moldy programmers like me still prefer the old Open Statement.
Here&#146;s sample code that does that</font></p>
<p><font color="#800000" size="2" face="Courier">Dim MyFileText
As String &#145;Makes a String Variable Called MyFileText <br>
Open &quot;MYFILE.TXT&quot; for input as #1 &#145;Opens The File
And Names It #1<br>
MyFileText = Input$(Lof(1),1) &#145;Assigns The Text In The File
To MyFileText<br>
Close #1 &#145;Closes The File</font><font color="#800000"
size="2" face="Verdana"><br>
</font></p>
<p><font color="#800000" size="2" face="Courier">Open
&quot;MYFILE.TXT&quot; for Input as #1 &#145;Opens The File And
Names It #1</font><font color="#800000" size="2" face="Verdana"><br>
</font><font size="2" face="Verdana">This line does the actual
opening bit. Myfile.Txt is the name of the file to be opened. You
can open a file in many ways for many purposes. I&#146;ve used
Input Mode here because I just want to read the contents of the
file. If you want <b>to write to a file use Output</b>, <b>use
Append to add in the end of the file and Random if you have a
Database in the file. Binary Mode can be used to load Bitmap or
Sound Files. </b>#1 is the number of the file. Whenever you want
to work on the file you will access it using that number.</font></p>
<p><font color="#800000" size="2" face="Courier">MyFileText =
Input$(Lof(1),1) &#145;Assigns The Text In The File To MyFileText<br>
</font><font size="2" face="Verdana">This line assigns the
contents of the file to MyFileText variable. Input$ Function
reads data from a file using the file number. </font></p>
<p><font size="2" face="Verdana">The first argument of Input$ is <b>Lof(1).
</b>The LOF function retrieves the length of a file in number of
characters. The second argument <b>1 </b>is the number of file,
which has to be read. So in practice we tell VB to read the
entire length [LOF(1)] of file number 1 in the variable
MyFileText.</font></p>
<p><font color="#800000" size="2" face="Courier">Close #1
&#145;Closes The File<br>
</font><font size="2" face="Verdana">This statement closes the
file and frees file number 1. It&#146;s a good practice to close
the file immediately after you&#146;ve read the contents in a
variable to free resources and avoid problems caused by a file
that remains open all the while the software is running. </font></p>
<p><font size="2" face="Verdana"><b>Problems with Opening File </b><br>
For most problems VB gives a self evident error message which
documents in detail the problem and allows the error to be
trapped and rectified. However there&#146;s a special case which
forced me to rack my brains for quite a while when I was new to
programming.</font></p>
<p><font size="2" face="Verdana">VB won&#146;t recognize and read
a file with a null terminated string in the normal input mode. Now in most editors like
NotePad etc., no null terminated string is added at the end of
the file but in some special cases, specially when the files has
been used for Binary purposes there may be a null terminated
string at the end of the file, and the file has to be opened in Binary mode in
VB, if you try to open it in input mode, there will be some cryptic error.Rectifying this problem is quite easy, just remove the last
character from the file and it gets opened fine.</font></p>
<p><font size="2" face="Verdana"><b>Writing Strings to Files<br>
</b>To put your string in a file use Output instead of Input to
open the file. To save your string into the file you can either
use Write # or Print # in this way.</font></p>
<p><font color="#800000" size="2" face="Courier">Write
#FileNumber, TheText<br>
</font><font size="2" face="Verdana">Or<br>
</font><font color="#800000" size="2" face="Courier">Print
#FileNumber, TheText</font></p>
<p><font size="2" face="Verdana"><b>Searching Stuff in Strings<br>
</b>You may often need to search for a word in lengths of text.
Visual Basic&#146;s Instr function does this great.</font></p>
<p><font color="#800000" size="2" face="Courier">Dim WordPos<br>
WordPos = Instr(1, MyText, MyWord, VbTextCompare)</font></p>
<p><font size="2" face="Verdana">Here WordPos holds the position
of the first character of the word if it is found in the file. </font></p>
<p><font size="2" face="Verdana"><b>The first argument
&#145;1&#146;</b> specifies the character no. from where Instr
should start looking. This is useful when you need to do multiple
searches or search from the middle of the text. You can also
leave this option blank if you want to search from the beginning
of the text.</font></p>
<p><font size="2" face="Verdana"><b>The second argument
&#145;MyText&#146; </b>specifies the name of the string variable
that has to be searched. You can also use a string length like
this one (&quot;I can use This String Instead Of MyText&quot;)
instead of the variable.</font></p>
<p><font size="2" face="Verdana"><b>The third argument
&#145;MyWord&#146;</b> is the word or character that has to be
searched in MyText. MyWord can also be a string instead of a
variable.</font></p>
<p><font size="2" face="Verdana"><b>The fourth argument
&#145;VbTextCompare&#146; </b>decides the mode of the comparison.
By default the mode is Binary. Here I am doing a comparison
between two strings, that&#146;s why I have used VbTextCompare
instead of the default VbBinaryCompare.</font></p>
<p><font size="2" face="Verdana">VbTextCompare is inferior to
Binary compare in speed. In fact when I ran a test which tried
finding the letter &#145;A&#146; in a string comprising of all
alphabets VbTextCompare took twice the time needed by
VbBinaryCompare to finish the searches. However I still prefer
using VbTextCompare in most cases because VbBinaryCompare thinks
Capital &#145;A&#146; and small &#145;a&#146; are different
characters and won&#146;t provide a match if the case is
different in the searched word and original string.</font></p>
<p><font size="2" face="Verdana">If Instr is successful in
finding a match it returns the position of the first character in
the word. If it is unsuccessful the function returns 0.</font></p>
<p><font size="2" face="Verdana"><b>Extracting parts from a
string<br>
</b>You may often to extract specific portion of a string and use
them. VB has three functions for extracting string parts. Left,
Mid &amp; Right.</font></p>
<p><font size="2" face="Verdana">VB Pros and Code invigilators
recommend using Mid for all types of extraction. It is entirely
possible to do almost everything with Mid, but they won&#146;t
have made Left &amp; Right if they weren&#146;t supposed to be
used.</font></p>
<p><font color="#800000" size="2" face="Courier">TheText =
Left(MyText,NoOfCharacters)<br>
</font><font size="2" face="Verdana">Left function retrieves
specified number of characters from the left of the specified
string for e.g. if you wrote </font><font color="#800000"
size="2" face="Verdana">MyText = Left(&quot;ABCD&quot;,3)</font><font
size="2" face="Verdana"> then left would give you
&quot;ABC&quot;. </font></p>
<p><font size="2" face="Verdana">Right returns the specified
number of characters from the rightmost part of the string.<br>
Mid is by far the most versatile, useful function which can serve
the function of both Left, Right and also extract text from the
middle of the document.</font></p>
<p><font color="#800000" size="2" face="Courier">MyText =
Mid(TheText,StartPos,LenOfText)<br>
</font><font size="2" face="Verdana">The first argument
&#145;TheText&#146; is the name of the string from which the text
has to be extracted. <br>
The second argument &#145;StartPos&#146; is the character
position from which Mid should start taking the text.<br>
The third argument &#145;LenOfText&#146; is the no of characters
that have to be picked up.</font></p>
<p><font size="2" face="Verdana"><b>Replacing Text In Strings<br>
</b>You can include this feature in your software using the Left,
Right, Mid and Instr functions. Let&#146;s see some sample code
which &#145;B&#146; with &#145;F&#146; in a string ABCD in this
fashion.</font></p>
<p><font color="#800000" size="2" face="Courier">Dim TheText as
String = &quot;ABCD&quot;<br>
Dim WordPos as Integer<br>
Dim MyTextLeft as String<br>
Dim MyTextRight as String</font></p>
<p><font size="2" face="Verdana">First find the text using Instr<br>
</font><font color="#800000" size="2" face="Courier">WordPos =
Instr(TheText, &quot;B&quot;) &#145;returns 2</font></p>
<p><font size="2" face="Verdana">Use Left to take text before the
searched character or word<br>
</font><font color="#800000" size="2" face="Courier">MyTextLeft =
Left(TheText, WordPos-1)</font></p>
<p><font size="2" face="Verdana">Use Right to take text after the
searched character<br>
</font><font color="#800000" size="2" face="Courier">MyTextRight
= Right(TheText, len(&quot;ABCD&quot;)-WordPos)<br>
</font><font size="2" face="Verdana">Or<br>
</font><font color="#800000" size="2" face="Courier">MyTextRight
=
Mid(TheText,WordPos+len(&quot;B&quot;),len(TheText)-WordPos+len(&quot;B&quot;))</font></p>
<p><font size="2" face="Verdana">Put The Two Strings Together
with the replaced character<br>
</font><font color="#800000" size="2" face="Courier">TheText =
MyTextLeft &amp; &quot;F&quot; &amp; MyTextRight</font></p>
<p><font size="2" face="Verdana">The Modus Operandi here is quite
simple. We look for the string in the text, take all the text
that is prior to the string with the left function, and all the
text that is present after the string using the Right or Mid
function. The two strings are then put together with the
replacement text or no text if the part of the string has to be
deleted.</font></p>
<p><font size="2" face="Verdana"><b>Replacing Easily<br>
</b>If you were intimidated by the long length and seemingly
complex code, you can do this much more easily if you have VB6.
The new Replace function eliminates several lines of code with a
single line.<br>
For e.g. if I want to replace all &quot;BBBB&quot; with
&quot;C&quot; I would use <br>
</font><font color="#800000" size="2" face="Courier">Replace(&quot;BBBB&quot;,&quot;B&quot;,&quot;C&quot;)</font></p>
<p><font size="2" face="Verdana">Here the first argument is the
original text, Second is the text to be searched and the third is
the alternative text. <br>
You can also specify the number of found words to be replaced
using an extra Count argument, i.e. set count as 1 if you want to
replace only the first find and none other or leave it to the
default to replace all finds. </font></p>
<p><font size="2" face="Verdana"><b>Encyrpting Strings<br>
</b>If you&#146;ve ever though about storing passwords or other
sensitive data in a file or a string you must have thought
Encrypting it. Several algorithms of encryption exist in the
market and some of them are very complex. You can make a simple
algorithm of your own by replacing the ASCII value of the
characters, however the approach provides a weak form of
encryption and can be broken very easily. However you can do
quality encryption very easily using the VB Xor function.
Here&#146;s a Function Which Encrypts text using the numerical
keys provided by the user.</font></p>
<p><font color="#800000" size="2" face="Courier">Public Function
XorEncrypt(Byval TheText As String, Byval Key1 As Integer, Byval
Key2 As Integer) As String<br>
For I = 1 to Len(TheText)<br>
XorEncrypt = XorEncrypt &amp; Asc(Mid(TheText, I, 1)) Xor Key1
Xor Key2 &amp; &quot;.&quot;<br>
Next<br>
End Function</font></p>
<p><font size="2" face="Verdana">This extremely small function
uses the unique features of Xor to provide good quality
Encryption. First the ASCII value of the character is Xor&#146;d
with Key1 and then the resultant value is Xor&#146;d with Key2
resulting in a random number that&#146;s very hard to decrypt,
the number is delimited by the period sign to distinguish two
characters from each other. Xor performs a bitwise calculation.
If you perform a Xor on two numbers and then Xor the resultant
figure with any of the two numbers Xor returns the other number.</font></p>
<p><font color="#800000" size="2" face="Courier">Public Function
XorDecrypt(Byval TheText As String, Byval Key1 As Integer, Byval
Key2 As Integer) As String<br>
Dim PeriodPos as Integer<br>
Do<br>
PeriodPos = instr(TheText,&quot;.&quot;)<br>
If Not PeriodPos=0 Then<br>
TheXordNum=Mid(TheText,1,PeriodPos-1)<br>
XorDecrypt = XorDecrypt &amp; Chr(Xor(Xor(TheXordNum, Key2),
Key1))<br>
TheText = Mid(TheText, PeriodPos+1) <br>
Else<br>
Exit Do<br>
Endif<br>
End Function</font></p>
<p><font size="2" face="Verdana">There&#146;s still a lot more to
strings, in fact a lot-lot more, we could talk about storing
Strings in .INI files, strings in registry, strings in Random
Access Files, Strings Compiled in .EXEs with resources and a
whole lot of other types of strings, but, I guess we won&#146;t
be covering all that in this article. If you found this of any
help please drop me a mail and I&#146;ll try to write all the
other parts as quick as possible.</font></p>
<p><font size="2" face="Verdana"><b>Searching for Stuff<br>
</b>The most common functionality needed by any user is searching. You can use
the 'Instr' statement for performing searched in VB. This is how a typical instr
looks.</font></p>
<p><font color="#800000" size="2" face="Courier">SearchPos</font><font color="#800000" size="2" face="Courier">
= instr(1,&quot;ABCD&quot;,&quot;C&quot;,vbTextCompare)</font></p>
<p><font size="2" face="Verdana">Most of you should already be familiar with the
instr statement, so I am not going to explain it here. The thing that needs a
though is the last parameter, vbTextCompare. What parameter you pass to this
option decides how fast your search will be. If you use vbTextCompare, instr
ignore case and search strings in both upper case and lower case, but the speed
will be slowed tremendously. If you use vbBinaryCompare, it speeds up the search
more than 10 times, but will match case will searching. Personally I recommend
you use vbBinaryCompare, if you can, the speed gained is tremendous.&nbsp;</font></p>
<p><font size="2" face="Verdana">Thanks<br>
Razoredge<br>
E-mail: </font><a href="mailto:psl@nde.vsnl.net.in"><font
size="2" face="Verdana">psl@nde.vsnl.net.in</font></a></p>
</body>
</html>

