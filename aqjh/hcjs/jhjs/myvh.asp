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
%>
<HTML><HEAD><TITLE><%=Application("aqjh_chatroomname")%>--��ͨ��������</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type="text/css">
<!-- 
a { text-decoration: none} 
--> 
A:link {COLOR: #0000ff;FONT-FAMILY: ����; }
A:visited {COLOR: #b8860b; FONT-FAMILY: ����; }
A:active {FONT-FAMILY: ����; }
A:hover {FONT-FAMILY: ����;COLOR: #FF0000; }
BODY {FONT-FAMILY: ����; FONT-SIZE: 10pt;COLOR: #ffffff;}
TABLE {FONT-FAMILY: ����; FONT-SIZE:10pt;COLOR : #ffffff;}
</style>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY bgColor=#85C2E0 text=#ffffff topMargin=5>
<DIV align=center>
<p align="center">[ <font color="#FF3300"><%=Application("aqjh_chatroomname")%>--�ҵĽ�ͨ����</font> ]</p>
<font color="#FF00FF"><a href="vehicle.asp" target="_self"  title="ȥ�����ʽ�����Ľ�ͨ����">����������[��ͨ����ר����]ҳ��</a></font><br><br>
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM vh WHERE b='"&aqjh_name&"'",conn
do while not rs.eof and not rs.bof
%>
  <td width=25% align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
  <img src="images/<%=rs("h")%>"><br><%=rs("a")%>����<font color="#ff0000">ʹ��״����<br></font><%if rs("c") then%>���룺[ʹ��]<%else%> ���룺[����]<%end if%> <%if rs("e") then%>���˳���[ʹ��] <%else%>���˳���[����]<%end if%></td>
  <td valign="middle"><table><tr><td colspan=2>������ڣ�<%=rs("l")%>�����;öȣ�<%=rs("i")%>����״̬��<%if rs("j") then%> �� <%else%> ����<%end if%>����ά�޴�����<%=rs("k")%>����<a href="rpvh.asp?id=<%=rs("id")%>" target="_self" title="ά����Ҫ֧��һ���ķ���">[����ά��]</a> <a href="rrvp.asp?id=<%=rs("id")%>" target="_self" title="ά����Ҫ֧��һ���ķ���">[���۾ɳ�]</a></td></tr><tr>
  
  <td colspan=2>
  <form  method="post" action="myvhok.asp">
  <div>�����ǳƣ�<input type="text" name="vhname" value="<%=rs("g")%>" maxlength=50 size=18>
  <select name="jrtc" >
  <%if not rs("c") and rs("e")  then%>
  <option value=true >����������ʱ</option><option value=false selected >�˳�������ʱ</option></select>
  <%else%>
  <option value=true  selected >����������ʱ</option><option value=false>�˳�������ʱ</option></select>
  <%end if%>
  <select name="jr"><option value=false>����</option><option value=true selected>ʹ��</option></select>
  <tr>
  <td >
  <%if not rs("c") and rs("e") then%>
  ���棺<input type="text" name="jrtg" value="<%=rs("f")%>" maxlength=255 size=60>
  <%else%>
  ���棺<input type="text" name="jrtg" value="<%=rs("d")%>" maxlength=255 size=60>
  <%end if%>
  </td>
  <td>
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <input type="submit" name="submit" value="ȷ��"></td>
  </tr>
  </div>
  </form></td></tr></table>
  </td>
  </tr><tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
</tr>
</table>
</div><div align="left">
<br> <font color="#0000ff">
�ᡡ��ʾ������ʱ�������ʹ��##�����Լ������֣�$$���������ǳơ�<br>
�����ʱ��##���Ÿ��򲻾õ�$$��������<%=Application("aqjh_chatroomname")%>��������λ��Ϻ������Ի������λ��Ϻ��С�������ˡ�<br>
���˳�ʱ��##����ģ���Խ�����ĸ�λ��Ϻ����������������һ���ˡ����������䣬����Լҵ�$$��ֻ��ƨ��ð�̣�������ʧ��Ӱ�١�<br>
<br>
ע�����⣺���뼰�˳�����Ȳ����԰�������ַ����ո�ȣ���ʹ��ȫ�Ǳ����ţ�ÿ������ֻ����һ���ǳơ�<br>
����������������˳�������ʱ���ֻ��һ�ֽ�ͨ���ߴ���ʹ��״̬����������һ��Ϊʹ��״̬�����������Ϊ��ʹ��״̬
<br>������ڣ���ά�޴���Ϊ0ʱ�ǹ������ڣ����������һ��ά�޵�����</font><br><br>
<font color="#ff0000">
��    �棺���빫��Ҫ��д��̫���Ļ��޳ܵĻ����ID��ɾIDû�����������������</font> 
<br><br><DIV align=center>
<font color="#FF00FF"><a href="#" onClick="window.open('vehicle.asp',target='_self')"  title="ȥ�����ʽ�����Ľ�ͨ����">����������[��ͨ����ר����]ҳ��</a></font><br><br><FONT color=#0000ff>&copy; ��Ȩ���� 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>���ֽ�����</FONT></A></div>
</body>
</html>