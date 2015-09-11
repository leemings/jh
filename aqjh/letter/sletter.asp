<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<%
id=request("id")
if not isnumeric(id) then
	Response.Write "<script Language=Javascript>alert('提示：ID错误！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你不是江湖中人！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
myid=rs("id")
rs.close
if aqjh_grade<10 then
	rs.open "select * from letter where yhid="&myid&" and id="&id&" and del=false",conn
else
	rs.open "select * from letter where id="&id,conn
end if
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您要查看的短信已删除或不存在！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
bt=rs("bt")
rq=rs("rq")
xj=rs("xj")
yhid=rs("yhid")
fjr=rs("fjr")
rs.close
if yhid=myid then
conn.execute "update letter set cs=true where id="&id
end if
set rs=nothing
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE>雪落枫飞江湖→鸿雁传情</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../jhimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=../jhimg/back002.gif bgColor=#7c6244 leftMargin=0 oncontextmenu="return false;" text=#cccccc topMargin=0 marginheight="0" marginwidth="0">
<div align=center><br>
<table border="0" cellpadding="0" cellspacing="0" width="508">
<tr>
<td width="22"><img src="../jhimg/jiao_z_s.gif" width="21" height="21"></td>
<td width="401" background="../jhimg/line_s.gif"></td>
<td width="85"><img src="../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr>
<td width="22" background="../jhimg/line_z.gif">&nbsp;</td>
<td width="401" align=center> 
<TABLE bgColor=#675138 border=0 cellPadding=5 cellSpacing=1 width=479 height="195" style="border-left: 1 solid #816445; border-right: 1 solid #A98761; border-top: 1 solid #A98761">
<TBODY> 
<TR align=middle> 
  
<TD height="15" bgcolor="#333300" width="464"> <SPAN class=title><IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17> 
  雪落枫飞江湖→我的信件<IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17></SPAN>   
</TD>
</TR>
<TR> 
  
<TD height="41" bgcolor="#675138" width="464"> 
<p style="text-indent: 19; line-height: 150%" align="center"> <a href="write.asp?id=<%=id%>"><img src="../jhimg/m_reply.gif" width="40" height="40" border="0"></a>
<a href="del.asp?id=<%=id%>"><img src="../jhimg/m_delete.gif" width="40" height="40" border="0"></a> 
<a href="zf.asp?id=<%=id%>"><img src="../jhimg/m_fw.gif" width="40" height="40" border="0"></a> 
<a href="write.asp"><img src="../jhimg/m_write.gif" width="40" height="40" border="0"></a>
</TD>
</TR>
<TR>
<TD height="90" bgcolor="#675138" width="464"> 
<div align="center">标题：<%=bt%>
<br>发信人：<%=fjr%> 日期：<%=rq%> <br>
<hr></div>
<br>
<%=xj%><br><br><br>
<div align="center">
<input type="button" value="关闭窗口" onClick="window.close()" name="button">
<br>
<br>
</div>
</TD>
</TR>
</TBODY> 
</TABLE>
<!--#include file="bottom.htm"-->
</td>
<td width="85" background="../jhimg/line_y.gif"></td>
</tr>
<tr>
<td width="22"><img src="../jhimg/jiao_z_x.gif" width="22" height="21"></td>
<td width="401" background="../jhimg/line_x.gif"></td>
<td width="85"><img src="../jhimg/jiao_y_x.gif" width="21" height="21"></td>
</tr>
</table>
</div>
</BODY></HTML>
