<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<!--#include file="function.asp"-->
<%
set rs=server.createobject("adodb.recordset")
if request("id")="" then 
sql="select * from user where UserName='"&request("UserName")&"'"
else
sql="select * from user where id="&request("id")
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
	response.write "<p align='center'>�Ҳ����������</p>" 
else
	id=rs("id")
	username=rs("username")
	Sex=rs("Sex")
	email=rs("Email")
	oicq=rs("oicq")
	homepage=rs("homepage")
	Address=rs("Address")
	tel=rs("tel")
	regDate=rs("regDate")
	lastlogin=rs("lastlogin")
%>
<script src=JS/Js.js></script>
<style type="text/css">
<!--
BODY {FONT-FAMILY: ����; FONT-SIZE: 9pt}
td {font-size: 9pt}
a:link {text-decoration: none; color: #000000}
a:visited {color: #000000; text-decoration: none}
a:active {color: #0000FF; text-decoration: underline}
a:hover {color: #0000FF; text-decoration: underline}
div {font-size: 9pt}
-->
</style>
<meta http-equiv="Content-Language" content="zh-cn">
<title>�û�<%=rs("UserName")%>����ϸ����</title>
<body topmargin="0" leftmargin="0" bgcolor=menu style="border:none" scroll=no>

<div align="center">
  <center>
  <table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C4BEB3" width="100%" id="AutoNumber1" bordercolorlight="#DDDDDD" bordercolordark="#DDDDDD">
    <tr>
      <td width="100%" height="122">��<div align="center">
             <table bgcolor=#000000 border=1 cellpadding=0 cellspacing=0 
            width="100%" valign="top" style="border-collapse: collapse" bordercolor="#7C7461" height="116">
        <tr bgcolor="#FFB400"> 
          <td align=middle colspan=2 bgcolor="#DED7E3">--==�û�<%=username%>����ϸ����==--</td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>�û�����</td>
          <td bgcolor=#F5F3F8><%=username%>��</td>
        </tr>
        <tr>
          <td align=right bgcolor=#F5F3F8 noWrap>�Ա�</td>
          <td bgcolor=#F5F3F8><%if sex=true then%>��<%else%>Ů<%end if%></td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>��ϵ��ַ��</td>
          <td bgcolor=#F5F3F8><%=Address%>��</td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>�绰���ֻ���</td>
          <td bgcolor=#F5F3F8><%=tel%>��</td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>���䣺</td>
          <td bgcolor=#F5F3F8><%=email%>��</td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>OICQ��</td>
          <td bgcolor=#F5F3F8><%=OICQ%>��</td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>��ַ��</td>
          <td bgcolor=#F5F3F8><a href=<%=homepage%> target=_blank><%=homepage%></a>��</td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>����½��</td>
          <td bgcolor=#F5F3F8><%=lastlogin%>��</td>
        </tr>
        <tr> 
          <td align=right bgcolor=#F5F3F8 noWrap>ע�����ڣ�</td>
          <td bgcolor=#F5F3F8><%=regDate%>��</td>
        </tr>
        </table></div>
      </td>
    </tr>
    </table>
  </center>
</div>
</body>
</html>
<%end if%>