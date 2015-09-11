<%@ LANGUAGE=VBScript codepage ="936" %>
<html>
<head>
<title>练药</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
TD { FONT-SIZE: 9pt } A:link { FONT-SIZE: 9pt; TEXT-DECORATION: none } A:visited { FONT-SIZE: 9pt; TEXT-DECORATION:
none } A:hover { COLOR: #cc9900; CURSOR: hand; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt;TEXT-DECORATION: underline blink }
</style>
</head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0">
<table width="640" border="0" height="461" background="images/back.jpg">
<tr>
<td height="397" valign="top" align="center"> <br>
      <b><font color="#FFFF00">可修练物品列表</font></b>(按名字进行修练)<br>
      <table width="86%" border="1" cellpadding="0" cellspacing="0" align="left">
        <tr>
<td width="19%">
<div align="center"><font color="#FFCC66"><b>名称</b></font></div>
</td>
<td width="58%">
<div align="center"><font color="#FFCC66"><b>所需材料</b></font></div>
</td>
<td width="33%">
<div align="center"><font color="#FFCC66"><b>功效</b></font></div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM x where f='药品'",conn
do while not rs.bof and not rs.eof
%>
<tr>
<td width="19%"><div align="center"><a href="xlwp.asp?xlwpname=<%=rs("a")%>" target="cz"><font color="#FFFFFF"><%=rs("a")%></font></a></div></td>
<td width="58%"><div align="center"><%=replace(rs("b"),"|"," ")%></div></td>
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
<div id="Layer1" style="position:absolute; width:114px; height:90px; z-index:1; left: 28px; top: 375px"><img src="images/fire.gif" width="82" height="70"></div>
<div id="Layer2" style="position:absolute; width:62px; height:53px; z-index:2; left: 128px; top: 419px">
<table width="75%" border="0" height="27">
<tr>
<td><img src="images/firejj.gif" width="70" height="13"></td>
</tr>
<tr>
<td>
<div align="right"><img src="images/huochai2.gif" width="26" height="22"></div>
</td>
</tr>
</table>
</div>
<div id="Layer3" style="position:absolute; width:200px; height:115px; z-index:3; left: 286px; top: 344px">
<table width="75%" border="0" height="27" cellspacing="00" cellpadding="0">
<tr>
<td>
<div align="center"><img src="images/taobar2.gif" width="34" height="40"></div>
</td>
</tr>
<tr>
<td>
<div align="center"><img src="images/ping.gif" width="51" height="65"></div>
</td></tr></table></div></td></tr></table></body></html>