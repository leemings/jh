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
%>
<HTML>
<HEAD>
<title>�����ӻ�</title>
<link rel="stylesheet" href="../../css.css">
</HEAD>
<body bgcolor="#8EB4D9" background="../../bg.gif">
<p align=center> <b> <font color="#000000" size="4">��Ҫ�ӻ�</font></b></p>

<p align="center"><img border="0" src="no_27.gif"></p>
<p align="center"><font lang="ZH-CN" color="#FFFFFF" face="����" size="2">���ǳ�ʱ�ղ�ʿ���������ǻص����ʱ�������ǽ��</font></p>
<p align="center"><font lang="ZH-CN" color="#FFFFFF" face="����" size="2">���⣬����㲻��Ҫ���ڵ���ż����������ټ����</font></p>
<p align="center"><font lang="ZH-CN" color="#FFFFFF" face="����" size="2">��ż�����ü�į�����ˡ���ô���ҿ��԰���<a href="panjue1.asp">ʵ���ӻ�</a></font></p>
<p align="center">��</p>
</body></HTML> 