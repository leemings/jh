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
rs.open "select ʦ�� from �û� where ����='"&sjjh_name&"'",conn
sf=rs("ʦ��")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>E�߽�����wWw.51eline.com��</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=sjjh_name%></font>��ӭ����ϰ�䳡</b></font></div>
<form method=POST action='wuguanok.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
            <td align=center><font color=blue>ʦ��:<%=sf%></font>
            <%if trim(sf)<>"��" then%>
              <select name=money size=1>
                <option value="1000" selected> ����</option>
                <option value="10000">����Ӳ��</option>
                <option value="100000">��ü�ķ�</option>
                <option value="1000000">��ң��</option>
                <option value="2000000">��ҹƶ�</option>
                <option value="10000000">��Ȼ����</option>
              </select>
              <%end if%>
</td>
</tr>
<tr>
<td  align=center>
            <%if trim(sf)<>"��" then%>
              <input type=submit value=��ʼ���� name="submit">
              <%end if%>
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
                �������</div>
</td>
</tr>
<tr>
            <td valign="top" > 
              <p><br>
            <%if trim(sf)="��" then%><div align="center">
            <font color=red>û��ʦ�����ܲ�����</font><BR></div>
            <%end if%>ʦ������ȥ������ѡ�����ʺ��Լ����ķ�(��Ǯ�ǲ�ͬ��Ӵ��)�ſ������쵽�������޵��书��</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF">E�߽���</font></b></font>
</div>
</body>
</html>



