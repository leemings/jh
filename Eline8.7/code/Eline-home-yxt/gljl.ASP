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
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from l where a<now()-7",conn,3,3
rs.open "Select * From l where d='管理记录' Order by a DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title><%=Application("sjjh_chatroomname")%>-江湖操作</title>
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

<body bgcolor="#CCCCCC" background="../jhimg/bk_hc3w.gif" link="#000000" vlink="#FF0000" alink="#000000">
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="15" width="100%" bgcolor="#669999"><font color="#669966"> <font color="#FFFFFF"><b>管理记录</b></font></font></td>
</tr>
</tbody>
</table>
<div align="center">
  <table width="92%" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
<td width="100%">
        <table border="1" cellspacing="1" cellpadding="0" width="99%"
bordercolorlight="#EFEFEF" bordercolor="#000000">
          <tr bgcolor="#FFFFFF">
<td width="11%">
              <div align="center"><font color="#FF6600"> 管理员名</font></div>
</td>
            <td width="23%"> 
              <div align="center"><font color="#FF6600"> 时 间 </font></div>
</td>
            <td width="17%"> 
              <div align="center"><font color="#FF6600"> ip记录</font></div>
</td>
            <td width="49%"> 
              <div align="center"><font color="#FF6600"> 事 件 数 据</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#85C2E0';">
<td width="11%">
<div align="center"> <font color="#0000FF"><%=rs("b")%></font>
</div>
</td>
            <td width="23%"> 
              <div align="center"> <%=rs("a")%> </div>
</td>
            <td width="17%"> 
              <div align="center"> <%=rs("c")%> </div>
</td>
            <td width="49%"> 
              <div align="center"> <%=Replace(rs("e"),"red","blue")%> </div>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[共<font color="red"><b><%=pa%></b></font>页]
</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="gljl.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a href="gljl.asp?page=<%=page+1%>">下一页</a>]</div>
</td>
</tr>
</table>
</table>
</div>
</body>
