<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
dim conn,rs,id
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("awjh_l_usermdb")
id=request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open ("SELECT * FROM 申冤 WHERE ID=" & id),conn,0,1
%>

<html>
<head>
<title>查看状词♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--td {  font-family: 宋体; font-size: 9pt}body {  font-family: 宋体; font-size: 9pt}select {  font-family: 宋体; font-size: 9pt}A {text-decoration: none; font-family: "宋体"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "宋体"; font-size: 9pt} .big {  font-family: 宋体; font-size: 12pt}.txt {  font-family: "宋体"; font-size: 10.8pt}
--></style>
</head>
<body bgcolor="#3a4b91" text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../images/8.jpg">
<hr size="1">
<table width="590" border="0" cellspacing="0" cellpadding="3" align="center">
<tr>
<td>受害人<b>:</b><b><%=rs("原告")%></b>&nbsp;悲惨地写道：</td>
<td align="right"><a href=javascript:history.back()>返 回</a></td>
</tr>
</table>
<table width="588" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000">
<tr >
<td height="67">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
<p><span class="txt"><br>
<%=changechr(rs("状词"))%><br>
<br>
<br>
对被告这种恶劣行径，希望给于<%=rs("要求")%> <br>
<br>
<font size="2" color="#ffffff"><%if info(5)="六扇门" or info(2)>=10 or awjh_l_nickname=admin then%></font>&gt;&gt;<a href="agree.asp?result=<%=rs("要求")%>&id=<%=id%>&bg=<%=rs("被告")%>&my=<%=rs("原告")%>">同意</a>
&gt;&gt;<a href="noagree.asp?id=<%=id%>">不同意</a><%end if %></span></p>
<p align="right">&nbsp;</p>
</td>
</tr>
</table>
</td>
</tr>
</table>
<%rs.close
set rs=nothing
%>
<hr size="1">
<p align="center">&nbsp;</p>
</body>
</html>
<%
function getorder(theid)
dim tmprs
tmprs=conn.execute("select [Order] from 申冤 Where ID=" & theid)
getorder=tmprs("Order")
set tmprs=nothing
end function

function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%>