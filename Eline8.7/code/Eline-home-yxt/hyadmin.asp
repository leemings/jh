<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
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
<title>��Ա���ݿ�����wWw.51eline.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><%=Application("sjjh_chatroomname")%>���ݿ�</p>
<p align="center"><a href="hygl.asp">��Ա�����������</a></p>
<p align="center"><b><font color="#FF0000">ע�⣺</font></b><font color="#FF0000">�����Ѿ��ǻ�Ա�ģ�������ɺ�ʱ��ᰴ��Ա����ʱ���<br>
  ������ʱ�䣬��Ա�ȼ������ó��µĵȼ�!</font></p>
<p align="center"><font color="#000000">˵�����ȼ��ƻ�Ա��������˵�����ݵ�ɱ��ƻ�Ա��Ϊ6.2���������ܣ����óɻ�Ա��<br>
  �����������ݵ�Ϊ������N�� ��N������paodian.asp�����ã�</font><font color="#FF0000"><br>
  <br>
  (<font color="#0000FF"><b>�ݵ�ɱ��ƻ�Ա���ֵȼ���</b></font>)<br>
  </font><a href="hydj.asp"><font size="+2" face="����_GB2312"><b>�ȼ��ƻ�Ա</b></font></a></p>
<p align="center"><a href="hypd.asp"><font size="+2" face="����_GB2312"><b>�ݵ��ƻ�Ա</b></font></a></p>
<p align="center"><a href="../jhmp/hy.asp">���л�Ա�б�</a> <a href="fine.asp">�����û�����</a></p>
</body>
</html>
