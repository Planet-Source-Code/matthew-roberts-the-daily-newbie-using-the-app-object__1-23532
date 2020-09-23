<div align="center">

## The Daily Newbie \- Using the App Object


</div>

### Description

Explains the basics of using the App Object.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-the-daily-newbie-using-the-app-object__1-23532/archive/master.zip)





### Source Code


<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Daily Newbie - 05/01/2001</TITLE>
</HEAD>
<BODY bgcolor="#ffffff">
<P></P>
<P class="MsoTitle"><IMG width="100%" height="3" v: shapes="_x0000_s1027"></P>
<P align="center" class="MsoTitle"><FONT size="7"><STRONG>The
Daily Newbie</STRONG></FONT></P>
<P align="center" class="MsoTitle"><STRONG>&#8220;To Start Things
Off Right&#8221;</STRONG></P>
<P align="center" class="MsoTitle"><FONT size="1">
May 8,
2001
</FONT></P>
<P align="center" class="MsoTitle"></P>
<P align="left" class="MsoNormal" style="TEXT-ALIGN: left">Love it, hate it, or just don't care, the Daily Newbie is back. I have decided to change the format a little. Although
	the layout is going to be the same as it always was, I am going to start using the PSC Ask A Pro discussion forum
	to choose my topics. I find that newbies make up a large part of that forum and they ask some pretty good questions.
	Also, if you have a question, email me and I will try to work it in.</P>
<P align="center" class="MsoNormal" style="TEXT-ALIGN: center"></P>
<P class="MsoNormal"><FONT face="Arial"></FONT></P>
<P class="MsoNormal"><FONT size="2" face="Arial"></FONT></P>
<P class="MsoNormal"><FONT size="2" face="Arial"></FONT></P>
<P class="MsoNormal" style="MARGIN-LEFT: 135pt; TEXT-INDENT: -135pt"><FONT size="2" face="Arial"><STRONG>Today&#8217;s Topic:</STRONG>
        </FONT><FONT size="4" face="Arial"> The App Object</FONT></P>
<P class="MsoNormal" style="MARGIN-LEFT: 135pt; TEXT-INDENT: -135pt"><FONT size="2" face="Arial"><STRONG>Name Derived
From:  </STRONG>   </FONT>
 <FONT size="2" face="Arial">"Application"</A></I> </EM></FONT></P>
<P></P>
<P class="MsoNormal" style="MARGIN-LEFT: 135pt; TEXT-INDENT: -135pt; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto"><FONT size="2" face="Arial"><STRONG>Used for: </STRONG>
Retriving information about your application at runtime.</FONT></P>
<P class="MsoNormal" style="MARGIN-LEFT: 135pt; TEXT-INDENT: -135pt; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto"><FONT size="2" face="Arial"><STRONG>VB Help Description: </STRONG>It determines or specifies information
about the application's title, version information, the path and name of its executable file and Help files,
and whether or not a previous instance of the application is running.
</FONT></P><FONT size="2" face="Arial"><STRONG>Plain
English: </STRONG>Returns information about the running application.
<P class="MsoNormal" style="MARGIN-LEFT: 135pt; TEXT-INDENT: -135pt; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto"><FONT< pre>
<FONT size="2" face="Arial"><STRONG>Syntax:  </STRONG>X =    App.{Property}   </FONT>
<PRE></PRE>
<P></P>
<FONT size="4" face="Arial"><STRONG><br><br>Properties:  </STRONG><BR>
<P class="MsoNormal" style="MARGIN-LEFT: 135pt; TEXT-INDENT: -135pt; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto"><FONT size="2" face="Arial"><STRONG>Usage:  </STRONG>   MsgBox "This application is named: " &amp; App.Title   </FONT></P>
<P class="MsoNormal" style="MARGIN-LEFT: 135pt; TEXT-INDENT: -135pt; mso-margin-top-alt: auto; mso-margin-bottom-alt: auto"><FONT face="arial" size="2">
	<I>Note: This article shows the most common and useful properties for the App object. There are a total of 30
	properties that you can access from code.</I>
<BLOCKQUOTE>
<BLOCKQUOTE>
<LI>Comments - The comments that were added in the Make tab of the project before compiling.
<LI>Company Name = The company name that was added in the Make tab before compiling. This is useful for copyright
	protection when creating reusable objects (.dll's or .ocx's)
	ActiveX .dll's or
<LI>EXEName - The name of the executable file that is running.
<LI>FileDescription - Again, entered in the Make tab before compiling. A general description of a project.
<LI>HelpFile - The Windows help file associated with this application. This property could be used to make sure
		the help file exists before trying to open it.
<LI>Major - The Major application version. In MyApp Version 2.5.34, the Major Version would be "2".
<LI>Minor - The Minor application version. In MyApp Version 2.5.34, the Minor Version would be "5".
<LI>Revision - The Revision (or "Build") number of they application version. In MyApp Version 2.2.34, the Revision would be "34"
<LI>Path - Probably the most commonly used property. Returns the full path to the folder that the executable was
				run from.
<LI>Title - The name of the application (i.e. MyCoolApp or whatever you compiled it as). This is not necessarily the same as the
			EXEName, since EXE's can be renamed at will.
</blockquote></blockquote>
<FONT size="4" face="Arial"><STRONG><br><br>Methods:  </STRONG></font><BR>
<LI>StartLogging - Sets the execution log path to a log file. Can also be set to log to the NT Event Log.
<LI>LogEvent - Causes a log event to be written to the
  log path that was specified in the StartLogging method.</LI></BLOCKQUOTE></BLOCKQUOTE>
Example:<BR><BR>To find out what path the .exe is running from:<BR><BR>
<BLOCKQUOTE>
<PRE style="MARGIN-LEFT: 1.25in; TEXT-INDENT: 0.35pt; tab-stops: 45.8pt 91.6pt 183.2pt 229.0pt 274.8pt 320.6pt 366.4pt 412.2pt 458.0pt 503.8pt 549.6pt 595.4pt 641.2pt 687.0pt 732.8pt"><FONT size="3" face="Arial">
<PRE>		MsgBox "The application is running from " &amp; App.Path
	</PRE></BLOCKQUOTE><BR><BR>
<BLOCKQUOTE>
<P></P>Today's code snippet returns a list of information about the current
application:
</FONT>
<P></P>
<P class="MsoNormal" style="MARGIN-LEFT: 135.35pt; TEXT-INDENT: -135.35pt"><FONT size="2" face="Arial"><STRONG>Copy &amp; Paste Code:</STRONG></FONT></P>
<P class="MsoNormal" style="MARGIN-LEFT: 135.35pt; TEXT-INDENT: -135.35pt"><FONT size="2" face="Arial"></FONT></P>
<PRE>
<FONT size="2" face="Arial">
<CODE></CODE></FONT></PRE>
<PRE style="MARGIN-LEFT: 1.25in; TEXT-INDENT: 0.35pt; tab-stops: 45.8pt 91.6pt 183.2pt 229.0pt 274.8pt 320.6pt 366.4pt 412.2pt 458.0pt 503.8pt 549.6pt 595.4pt 641.2pt 687.0pt 732.8pt"><FONT size="3" face="Arial">
<CODE><BR><BR>
	Debug.Print "Application Name: " &amp; App.Title
    Debug.Print "Running From: " &amp; App.Path
    Debug.Print "Version = " &amp; App.Major &amp; "." &amp; App.Minor &amp; App.Minor<BR><BR>
</CODE></FONT></PRE>
<BR><BR><b><font face="arial" size="3">Some Notes about the App Object:</b>
<br>
<LI>You can use the App.PrevInstance property to prevent your application from being run more than once on a single machine:
<PRE style="MARGIN-LEFT: 1.25in; TEXT-INDENT: 0.35pt; tab-stops: 45.8pt 91.6pt 183.2pt 229.0pt 274.8pt 320.6pt 366.4pt 412.2pt 458.0pt 503.8pt 549.6pt 595.4pt 641.2pt 687.0pt 732.8pt"><FONT size="3" face="Arial">
<CODE><BR><BR>
If App.PrevInstance = True Then
	MsgBox App.Title &amp; " is already running.
End If
<BR><BR>
</CODE></FONT></PRE></LI></FONT>
</blockquote></blockquote>
<li>You can open a local file from the application's folder without knowing what path the application is running from:
<CODE><BR><BR>
<BR><BR>
</CODE></FONT></PRE></LI></FONT>
<PRE style="MARGIN-LEFT: 1.25in; TEXT-INDENT: 0.35pt; tab-stops: 45.8pt 91.6pt 183.2pt 229.0pt 274.8pt 320.6pt 366.4pt 412.2pt 458.0pt 503.8pt 549.6pt 595.4pt 641.2pt 687.0pt 732.8pt"><FONT size="3" face="Arial">
<CODE><BR><BR>
Open App.Path & "customer.dat" For Input As #1
</CODE></FONT></PRE>
<br><br>
The App Object makes it easy to do some things that otherwise would be very difficult to do in VB. The App.Path property is
especially helpful when creating applications that manipulate files. Any comments about this article are welcome.
</BODY>
</HTML>

