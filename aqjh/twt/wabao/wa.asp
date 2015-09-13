<%@ LANGUAGE=VBScript%>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<8 then
ss=8-sj
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来挖宝贝吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>挖掘宝藏</title>
<meta http-equiv="refresh" content="10">
<style></style>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
</head>
<BODY oncontextmenu=self.event.returnValue=false BGCOLOR="#000000" text="#FFFFFF">
<DIV ID="Layer1" STYLE="position:absolute; left:46px; top:139px; width:192px; height:165px; z-index:3">
<%
if Session("diaoyu")=true then
	Session("diaoyu")=now()
end if 
abc=DateDiff("s",Session("diaoyu"),now())
if abc>=115 then
	Response.Write "<script Language=Javascript>alert('唉，让强盗发现了……');location.href = 'wano.asp';</script>"
	response.end
end if
if abc<=90 then
%>
<IMG SRC="wa.gif" width="50" height="50"><br>
    <font color="#FFFFFF" size="2">努力挖啊,我要发财!</font>
<%
else
%>
    <a href="waok.ASP"><IMG SRC="waok.gif" border="0" width="30" height="50"></a> 
    <font color="#0000FF"><br>
    </font> 
    <font color="#FFFFFF" size="2">哇,挖到大宝藏啊,快跑呀强盗要来了!</font>
<%
end if
%>
</div>
</DIV>
<DIV ID="Layer2" STYLE="position: absolute; left: 76px; top: 101px; width: 101px; height: 64px; z-index: 1"> 
  <p align="center"><img border="0" src="wolf-admin.gif"></DIV>  
<span id="liveclock" style="position: absolute; left: 196; top: 139; width: 76; height: 29"> 
<div align="center"><font color="#FFFFFF" size="2">消耗内力: <%=abc%></font> </div>
</span> 
<div id="Layer5" style="position:absolute; width:262px; height:21px; z-index:6; left: 119px; top: 382px"><font color="#FFFFFF" size="2">快乐江湖网&copy;版权所有 
  2001-2002</font><br>
</div></body></html>