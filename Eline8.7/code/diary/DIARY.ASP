<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "Select * From j where a='"& sjjh_name &"' Order by b DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title><%=Application("sjjh_chatroomname")%>-我的日记本</title>
<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#66CCCC">
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="15" width="79%" bgcolor="#669999"><font color="#669966"> <font color="#FFFFFF"><b>心情日记，记载您的心情！
</b></font><font color="#669966"><font color="#669966"><font color="#669966"><font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新此页面</a></font></font></font></font><font color="#FFFFFF"><b>
</b></font></font></td>
<td height="15" width="21%" bgcolor="#669999">
<div align="right"><font color="#669966"><font color="#FFFFFF"><b><a href="write.asp" target="_blank">开始写日记</a></b></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
<table width="92%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
        <table border="1" cellspacing="1" cellpadding="0" width="732"
bordercolorlight="#EFEFEF" background="backgroup1.gif">
          <tr bgcolor="#FFFFFF">
<td width="19%" height="16">
<div align="center"><font color="#FF6600"> 日期</font></div>
</td>
<td width="11%" height="16">
<div align="center"><font color="#FF6600"> 天气</font></div>
</td>
<td width="16%" height="16">
<div align="center"><font color="#FF6600"> 心情</font></div>
</td>
<td width="39%" height="16">
<div align="center"><font color="#FF6600">标题</font></div>
</td>
<td width="15%" height="16">
<div align="center"><font color="#FF6600">操作</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr>
<td width="19%" height="21">
<div align="center"> <font color="#0000FF"><%=rs("b")%></font>
</div>
</td>
<td width="11%" height="21">
<div align="center"> <%=rs("d")%> </div>
</td>
<td width="16%" height="21">
<div align="center"> <%=rs("c")%> </div>
</td>
<td width="39%" height="21">
<div align="center"><%=rs("e")%></div>
</td>
<td width="15%" height="21">
<div align="center">
<a href="view.asp?id=<%=rs("id")%>"><font color="#0000FF">查看</font></a>
<a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">删除</font></a>
</div>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
 <td align="left" width="37%" height="2">[共<font color="red"><b><%=pa%></b></font>页][<font
color="red"><b><%=mu%></b></font>条日记]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="diary.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a href="diary?page=<%=page+1%>">下一页</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 数量 from j where a='"& sjjh_name&"'")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%>

</body>