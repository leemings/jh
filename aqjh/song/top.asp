<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Buffer=true
Response.Expires=0
%>
<head>
<STYLE type=text/css>A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #000000; FONT-FAMILY: ����
}
A:visited {
	COLOR: #000000; FONT-FAMILY: ����
}
A:active {
	FONT-FAMILY: ����
}
A:hover {
	COLOR: #ff0000; FONT-FAMILY: ����
}
BODY {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
TABLE {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
A.aoyoo_style_1:link {
	COLOR: #ffffff; FONT-FAMILY: ����
}
A.aoyoo_style_1:visited {
	COLOR: #ffffff; FONT-FAMILY: ����
}
A.aoyoo_style_1:active {
	FONT-FAMILY: ����
}
A.aoyoo_style_1:hover {
	COLOR: #ffcc00; FONT-FAMILY: ����
}
A.aoyoo_style_1:unknown {
	COLOR: #ffffff; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
A.aoyoo_style_1:unknown {
	COLOR: #ffffff; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
</STYLE>
</head>
<body bgcolor=#FBF4E2 oncontextmenu=self.event.returnValue=false>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select top 6 * from music order by -dateandtime",conn
x=0
if rs.bof and rs.eof then
	response.write "<p align=center><font color=red>��ʱû���͸���Ϣ</font></p>"
else
%>
<%do while not rs.eof%>
<img src="music.gif"><bgsound src=chat/wav/xintie.wav loop=1> <acronym title="���:<%=rs("name")%>&#13&#10�͸�:<%=rs("toname")%>&#13&#10ף��:<%=rs("zhufu")%>&#13&#10�ͳ�ʱ��:<%=rs("dateandtime")%>"><font color=red><%=rs("name")%></font>Ϊ<font color=red><%=rs("toname")%></font>�ͳ�����[<a href="#" title="���:<%=rs("name")%>&#13&#10�͸�:<%=rs("toname")%>&#13&#10ף��:<%=rs("zhufu")%>&#13&#10�ͳ�ʱ��:<%=rs("dateandtime")%>" onClick="window.open('Play.asp?id=<%=rs("id")%>','url','resizable=no,width=250,height=80')"><font color=blue><%=rs("songname")%></font></a>]</acronym><br>
<%
x=x+1
if x=6 then exit do
rs.movenext
loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%><div align=right> [<a href="#" onClick="window.open('add.asp','add','width=200,height=240')">��Ҫ���</a>] [<a href="#" onClick="window.open('index.asp','more','width=440,height=280')">�鿴����</a>]...</div>