<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%><html>
<head>
<title><%=Application("aqjh_chatroomname")%>���ߵ�</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="../../bg.gif">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../../images/dj.gif" width="101" height="304"></td>
<td colspan="2" valign="top" align="center"> [ <font color="#009900">�� �� ��
�� ��</font> ]<br>
<br>
<table width="600" border="0" cellspacing="2" cellpadding="0">
<tr>
<td width="17%" nowrap>
<div align="center"><a href="BINQI.asp"><font color="#6666FF"><b>�ֳֵ���</b></font></a>
</div>
</td>
<td width="15%" nowrap>
<div align="center"><a
href="anqi.asp"><font color="#6666FF"><b>��������</b></font></a></div>
</td>
</tr>
</table>
<br>
      <table border="1" align="center" width="561" cellpadding="0" cellspacing="0" bordercolor="#000000">
        <tr align="center"> 
          <td width="85"><font color="#FF6600">װ �� ��</font></td>
          <td width="44"><font color="#FF6600">�� �� </font> 
          <td width="263"><font color="#FF6600">�� ��</font></td>
          <td width="58"><font color="#FF6600">�� ��</font></td>
          <td width="46"><font color="#FF6600">�� ��</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM b where b<>'����' and b<>'��Ƭ' and b<>'ҩƷ' and b<>'��ҩ' and b<>'�ʻ�'order by b,h ",conn
do while not rs.bof and not rs.eof
%>
        <tr> 
          <form method=POST action='binqi1.asp?id=<%=rs("id")%>'>
            <td width="85">
<div align="center"><%=rs("a")%></div></td>
            <td width="44">
<div align="center"><%=rs("b")%></div></td>
            <td width="263">
<div align="center">�ӷ���<%=abs(rs("g"))%>���ӹ���<%=abs(rs("f"))%></div></td>
            <td width="58"> 
              <div align="center"><%=rs("h")%>��</div></td>
            <td width="46"> 
              <div align="center">
              <input type="SUBMIT" name="Submit" value="����"></div>
            </td>
          </form>
        </tr>
        <%
rs.movenext
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
<br>
</body>
</html>
