<%
	Dim x : x = 0
	If Request.QueryString("x") = 1 Then
		x = 1
	End If
	If Request.Form("send")<>"" Then
		Dim my_from 	: my_from 		= Request.Form("INPNAME_FIRST") + " " + Request.Form("INPNAME_Last")
		Dim my_fromAddress: my_fromAddress	= Request.Form("INPEMAIL")
		Dim my_subject 	: my_subject 	= "This is a test page"
		Dim my_re 		: my_re 		= Request.Form("INPRE")
		Dim my_copy 	: my_copy		= Request.Form("INPBODY")
		Dim fullBody 	: fullBody		= ""
		Dim myMail

		Set myMail=CreateObject("CDO.Message")
		myMail.Subject 	= my_subject
		myMail.From 	= my_fromAddress
		myMail.To 		= "editor@calvarycomics.com"
		
		fullBody		= fullBody + "RE: " + my_re
		fullBody		= fullBody + " "
		fullBody		= fullBody + my_copy

		myMail.TextBody	= fullBody 

		myMail.Send
		set myMail 		= nothing
		Response.Redirect("contact.asp?x=1")
	End If
%>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="90%">
<%if x = 0 then%>
	<TR>
		<TD class="textBold" ALIGN=Right VALIGN=Top NOWRAP>Name:</TD>
		<TD WIDTH=10><IMG SRC="/img/space.gif" HEIGHT=1 WIDTH=10></TD>
		<TD>
		<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
			<TR>
				<TD><INPUT TYPE=text NAME="INPNAME_FIRST" SIZE=10 MAXLENGTH=20></TD>
				<TD WIDTH=10><IMG SRC="/img/space.gif" HEIGHT=1 WIDTH=10></TD>
				<TD><INPUT TYPE=text NAME="INPNAME_LAST" SIZE=40 MAXLENGTH=20></TD>
			</TR>
			<TR>
				<TD class="text1">First Name</TD>
				<TD WIDTH=10><IMG SRC="/img/space.gif" HEIGHT=1 WIDTH=10></TD>
				<TD class="text1">Last Name</TD>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=3><IMG SRC="/img/space.gif" HEIGHT=5 WIDTH=1></TD>
	</TR>
	<TR>
		<TD class="textBold" ALIGN=Right NOWRAP>Email:</TD>
		<TD WIDTH=10><IMG SRC="/img/space.gif" HEIGHT=1 WIDTH=10></TD>
		<TD><INPUT TYPE=text NAME="INPEMAIL" SIZE=50 MAXLENGTH=40></TD>
	</TR>
	<TR>
		<TD COLSPAN=3><IMG SRC="/img/space.gif" HEIGHT=5 WIDTH=1></TD>
	</TR>
	<TR>
		<TD COLSPAN=3><SPAN class="textBold">Email Content:</SPAN>
		<TABLE BORDER=1 CELLPADDING=4 CELLSPACING=0>
			<TR>
				<TD>
				<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD class="textBold" ALIGN=Right NOWRAP>Subject:</TD>
		<TD WIDTH=10><IMG SRC="/img/space.gif" HEIGHT=1 WIDTH=10></TD>
		<TD><INPUT TYPE=text NAME="INPSUBJECT" SIZE=60 MAXLENGTH=40></TD>
	</TR>
	<TR>
		<TD COLSPAN=3><IMG SRC="/img/space.gif" HEIGHT=5 WIDTH=1></TD>
	</TR>
	<TR>
		<TD class="textBold" ALIGN=Right NOWRAP>Regarding:</TD>
		<TD WIDTH=10><IMG SRC="/img/space.gif" HEIGHT=1 WIDTH=10></TD>
		<TD><SELECT NAME="INPRE" SIZE="1">
			<OPTION VALUE="INFORMATION">I need some information
		</SELECT></TD>
	</TR>
	<TR>
		<TD COLSPAN=3><IMG SRC="/img/space.gif" HEIGHT=5 WIDTH=1></TD>
	</TR>
	<TR>
		<TD class="textBold" ALIGN=Right VALIGN=Top NOWRAP>Body:</TD>
		<TD WIDTH=10><IMG SRC="/img/space.gif" HEIGHT=1 WIDTH=10></TD>
		<TD><TEXTAREA NAME="INPBODY" ROWS=10 COLS=60 WRAP="virtual"></TEXTAREA></TD>
	</TR>
				</TABLE>
				</TD>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=3><IMG SRC="/img/space.gif" HEIGHT=5 WIDTH=1></TD>
	</TR>
	<TR>
		<TD></TD>
		<TD COLSPAN=2><INPUT TYPE=submit NAME="send" VALUE="Send"> <INPUT TYPE=reset VALUE="Reset"> <INPUT TYPE=button VALUE="Close Window" onClick="javascript: opener.focus();self.close();"></TD>
	</TR>
<% else %>
			<TR>
				<TD ALIGN=Center class="header">***Your message was sent successfully***</TD>
			</TR>
			<TR>
				<TD ALIGN=Center><INPUT TYPE=submit NAME="back" VALUE="Back to Form"></TD>
			</TR>
<% End If%>
</TABLE>