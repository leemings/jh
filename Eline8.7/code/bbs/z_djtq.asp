<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<%
Response.Buffer=true
Response.CacheControl="no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires=0
if membername="" or not founduser then%> 
<script LANGUAGE="JavaScript">
if (window != window.top)
top.location.href = location.href;
</script>
<script language=javascript>
alert("对不起，请先登录论坛再来玩联线对战台球游戏。");  
window.close()
</script>
<% end if
if myarticle<10 then%> 
<script LANGUAGE="JavaScript">
if (window != window.top)
top.location.href = location.href;
</script>
<script language=javascript>
alert("对不起，发帖数少于10帖的用户不能玩联线对战台球游戏。");  
window.close()
</script>
<% end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=forum_info(0)%>台球世界</title>
<script language="JavaScript" src='tq/system/activex/gamexver.js'></script>
<SCRIPT language=JavaScript src="tq/game/files_dir.js"></SCRIPT>
<SCRIPT language=javascript event=OnCmd(type,cmd) for=GX>
switch(type){
case "appinit":
	GX.Cmd("<sys I0d<KKBOBKGLHQIDPHNNHHOJNMMIOKLRJPDLMECDCNHNODCFDLJINNJOGIGSNJAMCLPHGEBNLFNHIQBPDEEOHEIMDDONPGGOEKIDJFPQIQLKCDEDCIMIFHJLMOLJPNKOEQDKEPOJMFOJMMPGNRCNMEDFMIOIBEBGDNGNBHDIDILFGIFQMKLDBLKHJSINOHEH>");
	break;
}
</SCRIPT>
</head>
<BODY bgColor=#c0c0c0 leftMargin=0 topMargin=0 scroll=no>
<OBJECT id=GX codeBase=tq/system/activex/gamex.cab#Version="+top.gamexver+" classid=CLSID:2D4851FD-0BFE-11D4-9260-9AF666D52059 width="100%" height="100%" VIEWASTEXT></OBJECT>
</BODY>
</html>
