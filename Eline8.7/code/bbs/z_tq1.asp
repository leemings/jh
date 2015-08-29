<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_plus_check.asp"-->
<%
Response.Buffer=true
Response.CacheControl="no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires=0
dim sjjh_name,sjjh_grade,sjjh_jhdj
sjjh_name=membername
sjjh_grade=memberword 
sjjh_jhdj=memberclass 
if sjjh_name="" then%> 
<script LANGUAGE="JavaScript">
if (window != window.top)
top.location.href = location.href;
</script>
<script language=javascript>
alert("对不起，请先登录论坛再来玩联线对战台球游戏。");  
window.close()
</script>
<% end if
if myarticle<1 then%> 
<script LANGUAGE="JavaScript">
if (window != window.top)
top.location.href = location.href;
</script>
<script language=javascript>
alert("对不起，发帖数少于1帖的用户不能玩联线对战台球游戏。");  
window.close()
</script>
<% end if
Session("sjjh_name")=sjjh_name
Session("sjjh_grade")=sjjh_grade
Session("sjjh_jhdj")=sjjh_jhdj%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=forum_info(0)%>台球世界</title>
<script language="JavaScript" src='tq/system/activex/gamexver.js'></script>
<script language="JavaScript" src="tq/game/files_dir.js"></script>
</head>
<body bgcolor=black scroll=no topmargin="0" leftmargin="0">
<script>
function _GXonCmd_FBOX(type,cmd)
{
	switch(type){
		case "appinit":FBOX.OnGameInit();break;
		case "runfun":eval(cmd);break;
		case "script":execScript(cmd,"javascript");break;
	}
}
function StdGameXMain()
{
	var web="<SCRIPT LANGUAGE=javascript FOR=FBOX EVENT=OnCmd(type,cmd)>_GXonCmd_FBOX(type,cmd)</"+"SCRIPT>"
	web+="<OBJECT ID=FBOX CLASSID='CLSID:2D4851FD-0BFE-11D4-9260-9AF666D52059' CODEBASE='tq/system/activex/gamex.cab#Version="+top.gamexver+"' width=100% height=100%  VIEWASTEXT></OBJECT>"
	document.write(web);
	document.close();
	var obj=document.all["FBOX"]
	obj.OnGameInit=StdGameXInit;
}
function StdGameXInit(){

	var cmd="<sys setid="+this.id+"><sys roothtm="+document.location.href+">\
	<: if=(isobj(&gvar))?{;}:{gvar=new text{name=gvar;lock=1;src=param;show=0}};\
		&gvar.IP=210.34.9.198;&gvar.port=10003;\
		&gvar.loginname='<%=sjjh_name%>';&gvar.pass='<%=sjjh_grade%>';&gvar.gname=<%=forum_info(0)%>台球小世界;\
		_g_gnick=<%=forum_info(0)%>台球小世界;\
	>\
	<sys import=desktop,tq\\game\\in\\jh.gml>";
	this.Cmd(cmd);
}
StdGameXMain()
</script>
</body>
</html>
