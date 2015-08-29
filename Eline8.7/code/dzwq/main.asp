<%@ LANGUAGE=VBScript codepage ="936" %>
<html>
<head>
<title>武器♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--a:hover {  color: #0000FF; cursor: hand}-->
</style>
</head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0">
<table width="100%" border="0" height="461">
  <tr>
<td height="397" valign="top" align="center"> <br>
      <b><font color="#FFFF00">可修练物品列表</font></b>(按名字进行修练)<br>
      <table width="100%" border="1" cellpadding="0" cellspacing="0" align="left">
        <tr>
<td width="19%">
<div align="center"><font color="#FFCC66"><b>名称</b></font></div>
</td>
<td width="58%">
<div align="center"><font color="#FFCC66"><b>所需材料</b></font></div>
</td>
<td width="33%">
            <div align="center"><font color="#FFCC66"><b>说明</b></font></div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM x where f='锻造' order by g",conn
do while not rs.bof and not rs.eof
%>
<tr>
<td width="19%"><div align="center"><a href="dzwq.asp?wq=<%=rs("a")%>" target="cz"><font color="#FFFFFF"><%=rs("a")%></font></a><img src="../hcjs/jhjs/images/<%=rs("d")%>"></div></td>
<td width="58%">
            <div align="center"><font size="-1"><%=replace(rs("b"),"|"," ")%></font></div>
          </td>
<td width="33%"><div align="center"><%=rs("c")%></div></td>
</tr>
<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
    </td>
</tr>
</table>
</body>
</html>