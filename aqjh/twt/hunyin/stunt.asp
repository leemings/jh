<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ��ż<>'��' and ��ż<>���� and ����='"&aqjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��["&aqjh_name&"]�㻹û����ż�أ�����ʲô��');window.close();</script>"
	response.end
end if
rs.close
%>
<html>
<head>
<title>�Դ����޺��弼</title>
<style></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../css.css">
</head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"> <font color="#FFFFFF" ><b><font color="#000000">�� �� �� �� ��
�� ��</font></b></font><font color="#000000"><br>
<br>
</font> </div>
<table cellpadding="0" cellspacing="1" border="1" align="center" width="45%" bgcolor="#00CCFF" bordercolor="#000000">
  <tr> 
    <td height="11" width="42"> 
      <div align="center">ID</div>
    </td>
    <td height="11" width="97"> 
      <div align="center"> <font color="#000000" size="2">�� ʽ ��</font> </div>
    </td>
    <td height="11" width="93"> 
      <div align="center"> <font color="#000000" size="2">�� �� �� ��</font> </div>
    </td>
    <td height="11" width="104"> 
      <div align="center"> <font color="#000000" size="2">�� ��</font> </div>
    </td>
  </tr>
  <%
rs.open "SELECT * FROM t where b='" & aqjh_name & "' or c='" & aqjh_name & "'"
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='stuntok.asp?a=m'>
    <tr> 
      <td height="2" width="42"> 
        <div align="center"><font color="#000000" size="2"><%=rs("id")%></font></div>
      </td>
      <td height="2" width="97"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="wg" size="10" value="<%=rs("a")%>" maxlength="8">
          <input type="hidden" name="id" size="10" value="<%=rs("id")%>" maxlength="8">
          </font> </div>
      </td>
      <td height="2" width="93"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="nl" size="10" value="<%=rs("d")%>" maxlength="8">
          </font> </div>
      </td>
      <td height="2" width="104"> 
        <div align="center"> <font color="#000000" size="2">
<input type="submit" value="�޸�" name="submit">
          <a href="del.asp?id=<%=rs("id")%>">ɾ��</a></font> </div>
      </td>
    </tr>
  </form>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<5 then
%>
  <form method=POST action='stuntok.asp?a=n'>
    <tr> 
      <td width="42" height="2"> 
        <div align="center"><font color="#000000" size="2">�½�</font></div>
      </td>
      <td width="97" height="2"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="wg" size="10" maxlength="8">
          </font> </div>
      </td>
      <td width="93" height="2"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="nl" size="10" maxlength="8">
          </font> </div>
      </td>
      <td width="104" height="2"> 
        <div align="center"> <font color="#000000" size="2">
<input type="submit" value="���" name="submit">
          </font> </div>
      </td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center"><font color="#000000" size="2">[ע��ֻ���ѻ��������Ů�ſ����Դ���ÿ����һ����ȡ10000����������˫������ʹ�ã�������弼ʧЧ��]<br>
<br>
  [ɱ����=<b><font color="#0000FF">((�з�����+Ů������)</font></b><font color="#FF0000">-�Է�����</font>)+��������]<br>
<br><br>
<center>��Ȩ���С����齭����վ��
</font></p>
</body></html>