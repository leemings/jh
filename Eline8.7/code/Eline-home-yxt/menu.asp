<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>");
body{font-family:"����";font-size:12pt;}td{font-family:"����";font-size:10pt;line-height:125%;FILTER: dropshadow(color=#ffffff,offx=1,offy=1);}
A{color:black;text-decoration:none;}
A:Hover{color:blue;text-decoration:underline;cursor: crosshair;}
A:Active {color:black;text-decoration:none;}
</style>
<title>�������ġ�wWw.happyjh.com��</title>
</head>
<body bgcolor="#CCCCCC" background="../jhimg/bk_hc3w.gif" link="#000000" vlink="#000000" alink="#000000">
<table border="1" width="70%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="147" align="center">
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="addmp.asp" target="txt">�������</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="adminmpp.asp" target="txt">���ɹ���</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="adminwg.asp" target="txt">�书����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="adminxx.asp" target="txt">��ԯ����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="roomlist.asp" target="txt">�������</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="fine.asp" target="txt">�û�����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="gladmin.asp" target="txt">�ٸ�����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="seeuser.asp" target="txt">������ѯ</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="sqlcomm.asp" target="txt">SQl 
      ָ��</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="hyadmin.asp" target="txt">��Ա����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="wpadmin.asp" target="txt">��Ʒ����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="wpmoney.asp" target="txt">��۵���</a></td>
  </tr>
<tr>
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="hyk.asp" target="txt">�� Ա ��</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="ip.asp" target="txt">IP 
      ����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="manacc.asp" target="txt">�˺Ź���</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="gp/gp.asp" target="txt">��Ʊ����</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="../hc3w/compmdb.asp" target="txt">ѹ������</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="gljl.asp" target="txt">������¼</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="hysm.asp" target="txt">��ϵ��Ա</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="ggsm.asp" target="txt">�������</a></td>
  </tr>
<tr>
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="ZZFF.ASP" target="txt">վ������</a></td>
  </tr>
  <tr>
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="../yhy/admin1.asp" target="txt">������Ժ</a></td>
  </tr>
  <tr>
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="../TOP/MENU2.HTM" target="txt">��������</a></td>
  </tr>
  <tr>
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="100%" align="center" nowrap><a href="questionlist.asp" target="txt">�Զ�����</a></td>
  </tr>
</table>
<div align="center"><br>
</div>
</body>
</html>

