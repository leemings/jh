<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=Application("aqjh_chatroomname")%>-->点歌祝福</title>
<link href="CSS.CSS" rel="stylesheet" type="text/css">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
dim chinannid
chinannid=request("id")

set rs = Server.CreateObject("adodb.recordset")
sql="select * from music where ID="+cstr(chinannid)+""
rs.open sql,conn,1,1
if rs("toname")="大家" then
aqjh_name="大家"
end if
if aqjh_name<>rs("name") and aqjh_name<>rs("toname") then
 Response.Write "<script language=JavaScript>{alert('提示：对不起!此歌是["&rs("name")&"]送给["&rs("toname")&"]的!');window.close();}</script>"
    Response.End
end if
%>
<table align="center">
<object id="mPlayer1" width=200 height=60 classid="CLSID:6BF52A52-394A-11D3-B153-00C04F79FAA6" type=application/x-oleobject standby="Loading Windows Media Player components...">
<param name="URL" value="<%=rs("songurl")%>">
<param name="rate" value="1">
<param name="balance" value="0">
<param name="currentPosition" value="0">
<param name="defaultFrame" value="">
<param name="playCount" value="100">
<param name="autoStart" value="-1">
<param name="currentMarker" value="0">
<param name="invokeURLs" value="-1">
<param name="baseURL" value="">
<param name="volume" value="100">
<param name="mute" value="0">
<param name="uiMode" value="full">
<param name="stretchToFit" value="0">
<param name="windowlessVideo" value="-1">
<param name="enabled" value="-1">
<param name="enableContextMenu" value="0">
<param name="fullScreen" value="0">
<param name="SAMIStyle" value="">
<param name="SAMILang" value="">
<param name="SAMIFilename" value="">
<param name="captioningID" value="">
<param name="enableErrorDialogs" value="0">
<param name="_cx" value="12435">
<param name="_cy" value="1640">
</object>
</table>
<marquee scrollamount='1' scrolldelay='1' direction= 'left' id=zhufu onmouseover=zhufu.stop() onmouseout=zhufu.start()><%=rs("zhufu")%></marquee>
</body>
</html>
