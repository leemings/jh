<%@ LANGUAGE=VBScript codepage ="936" %>
<html>
<head>
<title>����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0">
<table width="100%" border="0" height="461">
  <tr>
<td height="397" valign="top" align="center"> <br>
      <b><font color="#FFFF00">��������Ʒ�б�</font></b>(�����ֽ�������)<br>
      <table width="100%" border="1" cellpadding="0" cellspacing="0" align="left">
        <tr>
<td width="19%">
<div align="center"><font color="#FFCC66"><b>����</b></font></div>
</td>
<td width="58%">
<div align="center"><font color="#FFCC66"><b>�������</b></font></div>
</td>
<td width="33%">
            <div align="center"><font color="#FFCC66"><b>˵��</b></font></div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM x where f='��Ʒ' order by g",conn
do while not rs.bof and not rs.eof
%>
<tr>
<td width="19%"><div align="center"><a href="dzwq.asp?wq=<%=rs("a")%>" target="cz"><font color="#FFFFFF"><%=rs("a")%></font></a><img src="../jhjs/images/<%=rs("d")%>"></div></td>
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
</table></td></tr></table></body></html>