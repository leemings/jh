<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<%
id=request("id")
newbt=""
newxj=""
newsjrq=""
a=""
b=""
newsjr=""
if id<>"" then
	if not isnumeric(id) then
		Response.Write "<script Language=Javascript>alert('提示：ID错误！');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select id from 用户 where 姓名='"&aqjh_name&"'",conn
	myid=rs("id")
	rs.close
	rs.open "select * from letter where yhid="&myid&" and id="&id&" and del=false",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：您没有收到过这封短信！');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	newsjrq=rs("rq")
	newbt="回复："&rs("bt")
	a="☆☆☆☆☆"&rs("fjr")&"在"&rs("rq")&"给你的信件中写道：☆☆☆☆☆"&chr(13)
	b="====================================="
	newxj=rs("xj")&chr(13)
	newsjr=rs("fjr")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
%>
<HTML><HEAD><TITLE>鸿雁传情</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../jhimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=../jhimg/back002.gif bgColor=#7c6244 leftMargin=0 oncontextmenu="return false;" text=#cccccc topMargin=0 marginheight="0" marginwidth="0">
<div align=center><br>
<table border="0" cellpadding="0" cellspacing="0" width="504">
<tr>
<td width="22"><img src="../jhimg/jiao_z_s.gif" width="21" height="21"></td>
<td width="468" background="../jhimg/line_s.gif"></td>
<td width="14"><img src="../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr>
<td width="22" background="../jhimg/line_z.gif">&nbsp;</td>
<td width="468" align=center> 
<TABLE bgColor=#675138 border=0 cellPadding=5 cellSpacing=1 width=521 height="195" style="border-left: 1 solid #816445; border-right: 1 solid #A98761; border-top: 1 solid #A98761">
<TBODY> 
<TR align=middle> 
<TD height="15" bgcolor="#333300" colspan="2" width="445"> <SPAN class=title><IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17> 
发送短消息 <IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17></SPAN>  
</TD>
</TR>
<form name="form1" method="post" action="writeok.asp">
<TR> 
<TD height="41" bgcolor="#675138" width="82" style="border: 1 solid #7C6244">
<p style="text-indent: 19; line-height: 150%" align="center">收件人</TD>
<TD height="41" bgcolor="#675138" width="407" style="border: 1 solid #7C6244">
<input type="text" name="name" value="<%=newsjr%>" size="10" maxlength="5">
</TD>
</TR>
<TR> 
<TD height="41" bgcolor="#675138" width="82" style="border: 1 solid #7C6244"> 
<p style="text-indent: 19; line-height: 150%" align="center">标&nbsp;&nbsp;题</TD>
<TD height="41" bgcolor="#675138" width="407" style="border: 1 solid #7C6244"> 
<input type="text" name="bt" size="30" maxlength="20" value="<%=newbt%>">
</TD>
</TR>
<TR> 
<TD height="41" bgcolor="#675138" width="82" valign="top" style="border: 1 solid #7C6244"> 
<p style="text-indent: 19; line-height: 150%" align="center">内&nbsp;&nbsp;容
</TD>
<TD height="41" bgcolor="#675138" width="407" style="border: 1 solid #7C6244"> 
<textarea name="txt" cols="55" rows="8"><%=a%>
<%=newxj%>
<%=b%></textarea>
</TD>
</TR>
<TR> 
<TD height="90" bgcolor="#675138" colspan="2" width="445"> 
<div align="left">说明：<br>
① 您可以使用Ctrl+Enter键快捷发送短信<br>
                  ② 标题最多20个字符，内容最多200个字符<br>
③ 在发言内容中请尽量使用全角字符，否则有些半角的标点符号会被自动屏蔽<br>
④ 每发一条短信收费1金币<br>
</div> 
<br>
<div align="center">
<input type="submit" name="Submit" value="发送">
<input type="reset" name="Submit2" value="清除">
<input type="button" name="Submit3" value="返回" onclick='javascript:history.go(-1)'>
<input type="button" value="关闭" onClick="window.close()" name="button">
<br>
</div>
</TD>
</TR>
</form>
</TBODY> 
</TABLE>
<!--#include file="bottom.htm"-->
</td>
<td width="14" background="../jhimg/line_y.gif">&nbsp;</td>
</tr>
<tr>
<td width="22"><img src="../jhimg/jiao_z_x.gif" width="22" height="21"></td>
<td width="468" background="../jhimg/line_x.gif"></td>
<td width="14"><img src="../jhimg/jiao_y_x.gif" width="21" height="21"></td>
</tr>
</table>
</div>
</BODY></HTML>
