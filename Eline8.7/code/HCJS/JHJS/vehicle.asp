<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML><HEAD><TITLE><%=Application("sjjh_chatroomname")%>��ͨ����ר����</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type="text/css">
<!-- 
a { text-decoration: none} 
--> 
A:link {COLOR: #ff00ff;FONT-FAMILY: ����; }
A:visited {COLOR: #ff00ff; FONT-FAMILY: ����; }
A:active {FONT-FAMILY: ����; }
A:hover {FONT-FAMILY: ����;COLOR: #FF0000; }
BODY {FONT-FAMILY: ����; FONT-SIZE: 10pt;COLOR: #ffffff;}
TABLE {FONT-FAMILY: ����; FONT-SIZE:10pt;COLOR : #ffffff;}
</style>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY bgColor=#85C2E0 text=#ffffff topMargin=5>
<DIV align=center>
<p align="center">[ <font color="#FF3300" size=4><%=Application("sjjh_chatroomname")%>��ͨ����ר����</font> ]</p>
<font color="#FF00FF" size=4><a href="myvh.asp" target="_self" title="װ����ͨ���߼����ƽ����˳������ҹ���">����������[װ������]ҳ��</a></font><br><br>
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b WHERE b='��ͨ'",conn
hh=-1
do while not rs.eof and not rs.bof 
hh=hh+1
ha=hh mod 4
if ha<>0 then
%>
  <td width=25% align="center" valign="bottom" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
  <img src="images/<%=rs("i")%>"><br><br>[<%=rs("a")%>]�����;ö�[<%=rs("j")%>]<br>����:<%=rs("h")%>�������:<%=rs("m")%>��<br>����:[<%=rs("l")%>]����<a href="buyVH.asp?wpname=<%=rs("a")%>"><font color="#0000FF">����</font></a> 
  </td>
<% else %>
 </tr><tr>
  <td width=25% align="center" valign="bottom" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
  <img src="images/<%=rs("i")%>"><br><br>[<%=rs("a")%>]�����;ö�[<%=rs("j")%>]<br>����:<%=rs("h")%>�������:<%=rs("m")%>��<br>����:[<%=rs("l")%>]����<a href="buyVH.asp?wpname=<%=rs("a")%>"><font color="#0000FF">����</font></a> 
  </td>
<%end if
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</tr>
</table>
</div>
<div align="center">
<br><font color="#FF00FF" size=4><a href="myvh.asp" target="_self"  title="װ����ͨ���߼����ƽ����˳������ҹ���">����������[װ������]ҳ��</a></font><br><br>
<font color="#ff00ff" size=4>E�߽���</font></div>

</div>
</body>
</html>
