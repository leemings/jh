<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
conn.execute "update �û� set ����ʱ��=now() where ����='"&aqjh_name&"'"
conn.close
set conn=nothing
Session("diaoyu")=true
ff="<font color=#ff0000>��Ϣ</font>"& aqjh_name &"�ں�ɽ�ڱ��أ�����ǿ���˶�����,�����˵��ֵĲƱ�."			'��������
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>������</title></head>
<body background="../../bg.gif">
<div align="center">
<p>&nbsp;</p>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="200">
<tr>
<td bgcolor=#C6BD9B>
<table>
<tr>
<td valign="top">
<div align="center">
<font color="#000000"><img border="0" src="../../chat/img/Error.gif">
��ʹ��,���Ҫ���ֵı��ر�ǿ��������...</font>
</div>
</table>
<div align="center"><br>
<input type=button value=�رմ��� onClick='window.close()' name="button">
</div></td></tr></table></div></body></html>