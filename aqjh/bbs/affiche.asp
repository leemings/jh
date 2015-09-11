<!-- #include file="setup.asp" -->
<title>公告栏 - Powered By BBSxp</title>
<%
top
%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="affiche.asp">社区公告</a></td>
</tr>
</table><br>
<%

sql="select * from affiche order by posttime Desc"
rs.Open sql,Conn,1
pagesetup=10 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
%>
<div align="center">
<table border="0" width="97%" cellspacing="1" cellpadding="5" height="64" class="a2">
	<tr>
		<td width="100%" height="20" class="a1">
		<p align="center"><%=rs("title")%></p>
		</td>
	</tr>
	<tr>
		<td width="100%" class="a3"><%=rs("content")%>
		</td>
	</tr>
	<tr>
		<td width="100%" class="a4" height="18">
		<p align="right">发布人 <%=rs("username")%>　发布时间 
		<font style="family:arial; font-size: 7pt"><%=rs("posttime")%></font> </p>
		</td>
	</tr>
</table></div>
<br>
<%
RS.MoveNext
loop
RS.Close
%><center>
[<b>
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"")
</script>
</b>]
<%
htmlend
%>