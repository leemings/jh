<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
dim abgcolor
dim username
dim TotalUseTable
dim AnnounceID
dim PostBuyUser,replyid
AnnounceID=0
TotalUseTable=NowUseBBS
abgcolor=Forum_body(5)
if boardid=0 then
	redim Board_Setting(23)
	Board_Setting(6)=1
	Board_Setting(5)=0:Board_Setting(7)=1
	Board_Setting(8)=1:Board_Setting(9)=1
	Board_Setting(10)=0:Board_Setting(11)=0
	Board_Setting(12)=0:Board_Setting(13)=0
	Board_Setting(14)=0:Board_Setting(15)=0
	Board_Setting(23)=0
end if
stats="预览帖子"
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=Forum_info(0)%>--<%=stats%></title>
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>>
<p>&nbsp;</p>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY> 
<TR>
<Th height=25>帖子预览</Th>
</TR>
<TR> 
<TD class=tablebody1 height=24>
<%
response.write "<b>"& htmlencode(request.form("title")) &"</b><br>"& dvbcode(request.form("body"),4,1) 
%></TD>
</TR>
</TBODY>
</TABLE>
<%
call footer()
%>