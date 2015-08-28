
<html>

<head>
<title>cctony首页更新</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet>
<script language="javaScript" type="text/javascript" SRC="js/pz_chromeless_2.1.js"></SCRIPT>
<script language="javaScript" type="text/javascript" SRC="js/openScript.js"></script>
</head>
<body><!-- #include file="connect.asp" -->
<%
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
%>
<%rs.Open "Select * From data Order By time DESC", conn, 1,1%>
<%if rs.eof and rs.bof then%>
暂时没有任何记录！！
<%else
RS.PageSize=50'设置每页记录数,可修改
Dim TotalPages
TotalPages = RS.PageCount

If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if

RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'设置最大记录数
Dim Totalcount
Totalcount =INT(RS.recordcount)

StartPageNum=1
do while StartPageNum+10<=CurPage
StartPageNum=StartPageNum+10
Loop

EndPageNum=StartPageNum+9

If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount       %>
<table border="1" width="80%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" align="center">
<%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (I<RS.PageSize)
p=p+1%>
<tr>
<td bordercolorlight="#C0C0C0" align=left><img src=images/dian.gif><a href=javascript:openScript("dismemo.asp?id=<%=rs("id")%>",500,400)><%=rs("name") %></a><font size=1 color=#CCCCCC>[<%=rs("time")%>]</font></td>
</tr>
<%I=I+1
RS.MoveNext
Loop%>
<tr><td colspan=4 align=middle bordercolorlight="#C0C0C0">页次： <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%>
页数： <a href="oldnews.asp?CurPage=<%=StartPageNum-1%>">&lt;&lt;</a>
<% For I=StartPageNum to EndPageNum
if I<>CurPage then %> <a href="oldnews.asp?CurPage=<%=I%>"><%=I%></a> <% else %>
<%=I%> <% end if %> <% Next %>
<% if EndPageNum<RS.Pagecount then %> <a href="oldnews.asp?CurPage=<%=EndPageNum+1%>">下页</a>
<%end if%>|<a href="oldnews.asp">刷 新</a></td></tr>
</table>
</body>
</html>
<%
end if
rs.close
set rs=nothing
%>
