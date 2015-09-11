<%@ LANGUAGE=VBScript.Encode%>
<% 
Response.Expires=0
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>MTV 欣赏</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link rel="stylesheet" href="../readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false background="../<%=chatimage%>" bgproperties="fixed" text="#FFFFFF">
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=list+document.af1.song.value;parent.f2.document.af.sytemp.focus();}</script>
<div align="center"><font color="#FFFF00" size=2>MTV欣赏 | <a href=song.asp><font color="#FFFF00" style="font-size:9pt">Mp3点歌</font></a></font></div>
<hr size=1 color=FFFF00><br>
<form name=af1 method=POST>
<table border="0" width="100%">
<tr><td><font size=2>选择好歌曲后按“确定”按钮，然后点击发言栏的“发言”按钮即可！(需要10万银两)</font></td></tr>
<tr>
<td align=center>
<br><br><font color="#FFFF00" size=2>选 择 曲 目</font><br><br>
<select name="song" style="font-size:9pt">
<%
Set connt=Server.CreateObject("ADODB.CONNECTION")
connt.open Application("aqjh_usermdb")
sql="select * from swf"
Set Rs=connt.Execute(sql)
do while not rs.eof
%>
<option value="<%=rs("id")%>"><%=rs("歌名")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td></tr>
<tr><td align=center>
<br><input type=button value=' 确 定 ' onclick="javascript:s('/送歌$')"></td></tr>
</table></form></html>