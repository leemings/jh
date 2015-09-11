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
    <td bgcolor=#FFFF00><A href="jiehun.asp" target=main>结婚登记</a></td>
    <td><A href="lihun.asp" target=main>离婚登记</a></td>
    <td bgcolor=#FFFF00><A href="yuelao.ASP" target=main>江湖求婚</a></td>
    <td><A href="yuanou.ASP" target=main>江湖怨偶</a></td><td bgcolor=#FFFF00><A href="lingyang.ASP" target=main>领养小孩</a></td></tr></table>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("aqjh_chatroomname")%>月老</b></font></p>