<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<html>
<head>
<title>��Ա���ݿ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>���ݿ�</p>
<p align="center"><b><font color="#FF0000">ע�⣺</font></b><font color="#FF0000">�����Ѿ��ǻ�Ա�ģ�������ɺ�ʱ��ᰴ��Ա����ʱ���<br>
  ������ʱ�䣬��Ա�ȼ������ó��µĵȼ�!</font></p>
<p align="center"><font color="#000000">˵�����ȼ��ƻ�Ա��������˵�����ݵ�ɱ��ƻ�Ա��Ϊͼ�ΰ��������ܣ����óɻ�Ա��<br>
  �����������ݵ�Ϊ������N�� ��N������paodian.asp�����ã�</font><font color="#FF0000"><br>
  <br>
  (<font color="#0000FF"><b>�ݵ�ɱ��ƻ�Ա���ֵȼ���</b></font>)<br>
  </font><a href="hydj.asp"><b>�ȼ��ƻ�Ա</b></a></p>
<p align="center"><a href="hypd.asp"><b>�ݵ��ƻ�Ա</b></font></a></p>
<p align="center"><a href="hylist.asp">���л�Ա�б�</a> <a href="fine.asp">�����û�����</a> <a href="kfaddadmin.asp">��Ա�����Ϳ� </a> <a href="jiangli.asp"><font color="#FF0000">���Ž���</font></a></p>
</body></html>