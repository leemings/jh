<html>
<head>
<%
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
%>
<title>������ȡ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
BODY {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
SELECT {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#CCCCCC" text="#000000" leftmargin="0" background="../bg.gif">
<div align="center">
<p><b><img border="0" src="../chat/img/menoy.gif"><font color="#008080"><b>������ȡ��</b></font></b> <br>
<font color="#800000"><font size="2">��Աÿ������쵽һ���������ͻ�Ա��,�ȼ�Խ�����ý�Խ�ࣻ����ÿ����쵽һ���Ľ�ң����ս���ȼ�20�����Ͽ���ȡ���1������Ա�����뵽��Ա��Ǯ����ȡ����,нˮ���㷽���״δ������ɼ�ת���������㣬ÿ��һ�ζ����쵽1�㹤��С�㣬�ô��ǳ���</font>
</font>
</p>
<p><a href="MONEY1.ASP"><font color="#008000">һ���ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY2.ASP"><font color="#008000">�����ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY3.ASP"><font color="#008000">�����ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY4.ASP"><font color="#008000">�ļ��ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY10.ASP"><font color="#008000">�弶�ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY06.ASP"><font color="#008000">�����ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY07.ASP"><font color="#008000">�߼��ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY08.ASP"><font color="#008000">�˼��ȼ��ƻ�Ա</font></a>
</p>
<p><a href="MONEY5.ASP"><font color="#008000">�ݵ��ƻ�Ա��Ǯ</font></a>
</p>
<p><a href="MONEY6.ASP"><font color="#008000">������</font></a><font color="#008000">-</font><a href="MONEY7.ASP"><font color="#008000">�ٸ���</font></a>
</p>
<p><a href="MONEY8.ASP"><font color="#008000">��ͨ����Ǯ</font></a>
</p>
<p><a href="MONEY9.ASP"><font color="#008000">20���Ϸ�����</font></a>
</p>
<p align="center">
<input type=button value=�رմ��� onClick='window.close()' name="button">
</p>
</div>
</body>
</html>