<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<LINK href="../../css.css" rel=stylesheet>
<body bgcolor='<%=bgcolor%>' background='../../bg.gif' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
  <tr align=center> 
    <td bgcolor=#FFFF00><A href="jiehun.asp" target=main>���Ǽ�</a></td>
    <td><A href="lihun.asp" target=main>���Ǽ�</a></td>
    <td bgcolor=#FFFF00><A href="yuelao.ASP" target=main>�������</a></td>
    <td><A href="yuanou.ASP" target=main>����Թż</a></td><td bgcolor=#FFFF00><A href="lingyang.ASP" target=main>����С��</a></td></tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("aqjh_chatroomname")%>����</b></font></p>