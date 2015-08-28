<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("yx8_mhjh_usercorp")<>"官府" or session("yx8_mhjh_username")<>Application("yx8_mhjh_admin")  then Response.Redirect "../../error.asp?id=046"
dim conn,rs,userconn,users
username=session("yx8_mhjh_username")
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
if session("yx8_mhjh_username")=""  then
Response.Redirect "../../error.asp?id=016"
else
%>
<!--#include file="data.asp"-->
<html>
<head>
<title>申冤状纸一览</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--td {  font-family: 宋体; font-size: 9pt}body {  font-family: 宋体; font-size: 9pt}select {  font-family: 宋体; font-size: 9pt}A {text-decoration: none; font-family: "宋体"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "宋体"; font-size: 9pt} .big {  font-family: 宋体; font-size: 12pt}
--></style>

</head>
<body leftmargin="0" topmargin="0" background='../../chatroom/bg1.gif'>
<table width="557" border="0" cellspacing="2" cellpadding="2" align="center">
<tr>
<td colspan="2" height="26" width="547">&nbsp;</td>
</tr>
</table>

<form action=over.asp method=get>
<%
set rs=server.createobject("adodb.recordset")
rs.open "SELECT * FROM bbs  ORDER BY id DESC",Conn,1,1
if not rs.eof and not rs.bof then
RS.PageSize=15
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
<table width="560" border="1" cellspacing="0" cellpadding="2" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td width="255"><font size="2" color="#000000">○ 共有<%=Totalcount%>张状纸，共<%=TotalPages%>页，目前为第<%=CurPage%>页</font></td> 
<td align="right" width="291"> <font color="#0000FF"> <a href="index.asp"><font size="2"> 添加状纸</font></a><font size="2" color="#FFFF00">&nbsp;&nbsp;&nbsp;</font></font><font size="2" color="#000000">&nbsp;&nbsp;<a href=manage.asp?owner=<%=owner%>>[刷新]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage-1%>>[上一页]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage+1%>>[下一页]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=1>[首页]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=TotalPages%>>[尾页]</a>&nbsp;</font></td>
</tr>
</table>
<table width="558" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">        <tr height=22>
<td height="25" width="290"><font size="2"><b><font color="#000000">&nbsp;&nbsp;&nbsp;</font></b><font>状纸标题</font></font></td>
<td align="center" colspan="2" height="25" width="79"><font size="2">受害者</font></td>
<td align="center" height="25" width="85"><font size="2">被告者
</font></td>
<td align="center" height="25" width="86"><font size="2">审判结果</font></td>
</tr>
<%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (II<RS.PageSize)
p=p+1%>
<tr>
<td height="25" width="290"><font size="2"><a href=rootdisp.asp?ID=<%=rs("id")%>><%=rs("标题")%></a>
</font></td>
<td colspan="2" height="25" width="79"><b><font size="2"><%=left(rs("原告"),10)%></font></b></td>
<td width="85" height="25"><b><font size="2"><%=left(rs("被告"),10)%></font></b></td>
<td width="86" height="25"><b><font size="2"><%if rs("结果")="N" then%>未审理<% end if %><%if  rs("结果")="0" then %>不接受<% end if %><% if  rs("结果")="1" then %>接受<% end if %></font></b></td>
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
<table width="559" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td width="274">
<div align="center">○ 页次： <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%> 
，每页：<font color="#CC0000"><%=RS.PageSize%></font>帖 </div> 
</td> 
<td align="right" width="267"> 
<div align="center">页数： <a href="manage.asp?owner=<%=owner%>&CurPage=<%=StartPageNum-1%>"><<</a> 
<% For I=StartPageNum to EndPageNum 
if I<>CurPage then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=I%>">[<%=I%>]</a> 
<% else %> <font color="#CC0000"><b><%=I%></b></font> <% end if %> <% Next %> 
<% if EndPageNum<RS.Pagecount then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=EndPageNum+1%>">[更多...]</a> 
<%end if%> </div> 
</tr> 
</table> 
<%else%> <br> 
<div align="center"><font color="#0000FF">暂时无状纸，请<a href="index.asp">添加新状纸</a> 
  </font>
<%
end if%> </div>
</form>

<% end if%>
</body>
</html>
<%
rs.close
set rs=nothing
%>
