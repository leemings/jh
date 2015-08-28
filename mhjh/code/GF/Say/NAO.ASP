<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
dim conn,rs,userconn,users
username=session("yx8_mhjh_username")
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
if session("yx8_mhjh_username")=""  then
Response.Redirect "../error.asp?id=016"
else
%>
<!--#include file="data.asp"-->
<html>
<head>
<title>申冤状纸一览</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet>
</head>
<body background='../../chatroom/bg1.gif' leftmargin="5" marginwidth="5" oncontextmenu="self.event.returnValue=false">
<table width="544" border="0" cellspacing="2" cellpadding="2" align="center">
<tr>
<td colspan="2" height="26" width="534">
<p align="center"><b><font size="3">公      
正 廉 明</font></b></p>      
</td>      
</tr>      
</table>      
      
<form action=over.asp method=get>      
<%      
set rs=server.createobject("adodb.recordset")      
rs.open "SELECT * FROM bbs  ORDER BY id DESC",Conn,1,1      
if not rs.eof and not rs.bof then      
RS.PageSize=10      
Dim TotalPages      
TotalPages = RS.PageCount      
      
If CurPage>RS.Pagecount Then      
CurPage=RS.Pagecount      
end if      
      
RS.AbsolutePage=CurPage      
      
rs.CacheSize = RS.PageSize      
      
Dim Totalcount      
Totalcount =INT(RS.recordcount)      
%>      
<table width="554" border="1" cellspacing="0" cellpadding="2" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">      
<tr>      
<td width="229"><font  color="#000000">○ 共有<%=Totalcount%>张状纸，共<%=TotalPages%>页，目前为第<%=CurPage%>页</font></td>      
<td align="right" width="311"> <font color="#000000"> <a href="index.asp">添加状纸</a>&nbsp;&nbsp;&nbsp;<a href=nao.asp?owner=<%=owner%>>[刷新]</a>&nbsp;<a href=nao.asp?owner=<%=owner%>&Curpage=<%=Curpage-1%>>[上一页]</a>&nbsp;<a href=nao.asp?owner=<%=owner%>&Curpage=<%=Curpage+1%>>[下一页]</a>&nbsp;<a href=nao.asp?owner=<%=owner%>&Curpage=1>[首页]</a>&nbsp;<a href=nao.asp?owner=<%=owner%>&Curpage=<%=TotalPages%>>[尾页]</a>&nbsp;</font></td>     
</tr>     
</table>     
<table width=553 border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" >     
<tr>     
<td height="25" width="281" align="center"><b>状纸标题</b></td>     
<td align="center" colspan="2" height="25" width="170"><b>受害者</b></td>     
<td align="center" height="25" width="80"><b>被告者</b></td>     
<td align="center" height="25" width="72"><b>审判结果</b></td>     
</tr>     
<%I=0     
p=RS.PageSize*(Curpage-1)     
do while (Not RS.Eof) and (II<RS.PageSize)     
p=p+1%>     
<tr>     
<td height="25" width="281" align="center"><a><%=rs("标题")%></a>     
</td>     
<td colspan="2" height="25" width="170"><b><%=left(rs("原告"),10)%></b></td>     
<td width="80" height="25"><b><%=left(rs("被告"),10)%></b></td>     
<td width="72" height="25"><b><%if rs("结果")="N" then%>未审理<% end if %><%if  rs("结果")="0" then %>不接受<% end if %><% if  rs("结果")="1" then %>接受<% end if %></b></td>     
</tr>     
<%rs.movenext     
II=II+1     
loop%>     
</table>     
<%StartPageNum=1     
do while StartPageNum+15<=CurPage     
StartPageNum=StartPageNum+15     
Loop     
     
EndPageNum=StartPageNum+14     
     
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount     
%>     
<table width="552" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">     
<tr>     
<td width="342">     
<div align="center">○ 页次： <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%>      
，每页：<font color="#CC0000"><%=RS.PageSize%></font>帖 </div>      
</td>      
<td align="right" width="192">      
<div align="center">页数： <a href="nao.asp?owner=<%=owner%>&CurPage=<%=StartPageNum-1%>"><font color="#0000FF"><<</font></a>      
<font color="#0000FF">      
<% For I=StartPageNum to EndPageNum      
if I<>CurPage then %> <a href="nao.asp?owner=<%=owner%>&CurPage=<%=I%>">[<%=I%>]</a>      
<% else %>  </font> <font color="#CC0000"><b><%=I%></b></font> <% end if %> <% Next %>      
<% if EndPageNum<RS.Pagecount then %> <font color="#0000FF"> <a href="nao.asp?owner=<%=owner%>&CurPage=<%=EndPageNum+1%>">[更多...]</a>      
</font>      
<%end if%> </div>      
</tr>      
</table>      
<%else%> <br>      
<div align="center"><font color="#FFFFFF">暂时无状纸&nbsp; </font><a href="index.asp"><font color="#FFFFFF">添加新状纸</font></a>      
<font color="#0000FF">      
</font>      
<%      
end if%> </div>      
</form>      
      
<% end if%>      
<div align="center">      
<hr width="80%">      
</div>      
<p align="center">爱火神话&nbsp; Copyright 2001-2005 [ <a href="manage.asp"><font color="#000000">官府审案</font></a>     
<font color="#FFFFFF"> </font>      
]</p>      
</body>      
</html>      
<%      
rs.close      
set rs=nothing      
%>      
