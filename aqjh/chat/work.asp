<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>ְҵ</title>
<style type="text/css">
<!--
body {  font-size: 9pt}
td {  font-size: 9pt}
a:link {  text-decoration: none}
a:hover {  text-decoration: underline}
-->
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p>
<font color="#000000">ְҵѡ��</font><br></font>
<span style='font-size:9pt'><br>
��<a href="#"onClick="window.open('../work/ice/icemain.asp','caibing','scrollbars=yes,resizable=yes,width=700,height=420')"
title="�ɱ�׬Ǯ��">�� ��</a>��<br>
<br><span style='font-size:9pt'>��<a href="#"onClick="window.open('../work/mine/minemain.asp','caikuang','scrollbars=yes,resizable=yes,width=700,height=420')"
title="�ɱ�׬Ǯ��">�� ��</a>��</p>
<p><span style='font-size:9pt'>��<a href="#"onClick="window.open('../work/tie/tiemain.asp','liantie','scrollbars=yes,resizable=yes,width=700,height=420')"
title="�ɱ�׬Ǯ��">�� ��</a>��</p>
<p>��ͬ��ְҵ�����в�ͬ�����룬��Ȼ����¡�������������������Ҳ�ǲ���ͬ�ģ���ҿ��Ը����Լ���ϲ����ѡ���Լ���ְҵ��<br>
</p></div></body></html>