<!-- #include file="setup.asp" -->
<title>������ - Powered By BBSxp</title>
<%
top
%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="affiche.asp">��������</a></td>
</tr>
</table><br>
<%

sql="select * from affiche order by posttime Desc"
rs.Open sql,Conn,1
pagesetup=10 '�趨ÿҳ����ʾ����
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '��ҳ��
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��
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
		<p align="right">������ <%=rs("username")%>������ʱ�� 
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