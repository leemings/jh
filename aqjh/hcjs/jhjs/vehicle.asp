<%@ LANGUAGE=VBScript codepage ="936" %>
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
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
%>
<HTML><HEAD><TITLE><%=Application("aqjh_chatroomname")%>��ͨ����ר����</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK href="../../css.css" type=text/css rel=stylesheet>
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
<p align="center">[ <font color="#FF3300"><%=Application("aqjh_chatroomname")%>��ͨ����ר����</font> ]</p>
<font color="#FF00FF"><a href="myvh.asp" target="_self" title="װ����ͨ���߼����ƽ����˳������ҹ���">����������[װ������]ҳ��</a></font><br><br>
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
<br><font color="#FF00FF"><a href="myvh.asp" target="_self"  title="װ����ͨ���߼����ƽ����˳������ҹ���">����������[װ������]ҳ��</a></font><br><br>
<font color="#ff00ff"><FONT color=#0000ff>&copy; ��Ȩ���� 2004-2005 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>���ֽ�����</FONT></A></font></div>
</div></body></html>