<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 性别 from 用户 where 姓名='" & sjjh_name &"'",conn
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("sjjh_chatroomname")%>[<%=sjjh_name%>]状态</title>
</head oncontextmenu=self.event.returnValue=false >

<frameset rows="25,*" framespacing="0" border="0" frameborder="0">
<frame name="header" target="main" <%if rs("性别")="男" then%>src="cd.asp"<%else%>src="cd1.asp"<%end if%> noresize scrolling="no">
<frame name="main" <%if rs("性别")="男" then%>src="zb.asp"<%else%>src="zb2.asp"<%end if%> scrolling="no">
<noframes>
<body>

<p>此网页使用了框架，但您的浏览器不支持框架。</p>

</body>
</noframes>
</frameset>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>