<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then 
	Response.Write "<script Language=Javascript>top.location.href='http://www.51eline.com';alert('提示：对不起，您还没有登陆江湖！');</script>"
	Response.End
end if
if request("songid")<>"" then
    set rs=server.createobject("adodb.recordset")
    id=request("songid")
	sql="select * from MusicDJ where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
	LF_Path=rs("LF_Path")
	rs("Hits")=rs("Hits")+1
	source=LF_Path
%>
<%
if replace(request("songid")," ","")="" then
Songid=request.cookies("playlist")
else
songid=replace(request("songid")," ","")
end if
if songid="" then
response.write("操作错误，请先选择要试听的舞曲！")
response.end
end if
if urgent<>"" then
response.write urgent
response.end
end if
if freelisten<>true then
response.write"<script>alert(""禁止游客听歌，请注册！"");window.close();</script>"
response.end
end if
%>
<html>
<script>
songid="<%=songid%>"
if(window.name!='HeiRui_Studio_Player')
{
alert("干你妈啦！吓看你妈了逼啊！再看整死你！草你妈的！");top.location="";
}
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>E线江湖总站→DJ舞吧→舞曲视听</title>
<style type=text/css>
.font1  { color:red; }
body, td, p  { font-size:12px; }
</style>
<body leftMargin='0' topMargin='0' MARGINHEIGHT='0' MARGINWIDTH='0' oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<table border="0" cellpadding="0" cellspacing="0" width="276">
  <tr>
   <td colspan="4"><img border="0" src="images/HeiRui_r1_c1.gif"></td>
  </tr>
  <tr>
   <td rowspan="5"><img name="HeiRui_r5_c1" src="images/HeiRui_r5_c1.gif" width="28" height="325" border="0" alt=""></td>
   <td><img name="HeiRui_r5_c2" src="images/HeiRui_r5_c2.gif" width="219" height="20" border="0" alt=""></td>
   <td rowspan="5"><img name="HeiRui_r5_c6" src="images/HeiRui_r5_c6.gif" width="29" height="325" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="20" border="0" alt=""></td>
  </tr>
  <tr>
   <td bgcolor="#000000">
         <script LANGUAGE="javaScript">              
var ERR_NonePlayer="播放器提示您:您的系统中没有安装Real Player播放器，请在主页上下载。\n\nMovie播放器将关闭。";
var ERR_FileNotFind="播放器提示您:对不起，未发现你要点播的舞曲。\n\n请更换其他舞曲！";
var ERR_NotLocateServer="播放器提示您:无法定位Movie服务器。请更换其他舞曲！";
var ERR_UnkownError="播放器提示您:在线用户过多，服务器限制，请抢线或稍后观看！";
         </script>
      <script LANGUAGE="VBScript">                                    
on error resume next
RealPlayerG2 = (NOT IsNull(CreateObject("rmocx.RealPlayer G2 Control")))\n'); 
RealPlayer5 = (NOT IsNull(CreateObject("RealPlayer.RealPlayer(tm) ActiveX Control (32-bit)")))
RealPlayer4 = (NOT IsNull(CreateObject("RealVideo.RealVideo(tm) ActiveX Control (32-bit)")))
if not RealPlayerG2 and RealPlayer5 and RealPlayer4 then
		if MsgBox("您的浏览器无法自动下载最新的浏览器插件，是否要下载播放器来播放？", vbYesNo) = vbYes then
			window.location = "http://51cd.com/2.exe"
		end if
end if

Sub player_OnBuffering(lFlags,lPercentage)
	if (lPercentage=100) then
		StartPlay=false
		if (FirstPlay) then
			FirstPlay=false
		end if	
		exit sub
	end if
End Sub
Sub player_OnErrorMessage(uSeverity, uRMACode, uUserCode, pUserString, pMoreInfoURL, pErrorString)
select case player.GetLastErrorRMACode()
		case -2147221496
			window.alert(ERR_FileNotFind)
		case -2147221433,-2147221428,-2147221417,-2147217468
			window.alert(ERR_NotLocateServer)
		case else
			window.alert(ERR_UnkownError)
	end select
End Sub
         </script>
   <p align="center">
      <object id="player" name="player" classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" width="216" height="178">
        <param name="_ExtentX" value="5715">
        <param name="_ExtentY" value="4710">
        <param name="AUTOSTART" value="-1">
        <param name="SHUFFLE" value="0">
        <param name="PREFETCH" value="0">
        <param name="NOLABELS" value="-1">
        <param name="CONTROLS" value="imagewindow,controlpanel">
        <param name="CONSOLE" value="clip1">                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  <param name="SRC" value="<%=source%>">
        <param name="LOOP" value="-1">
        <param name="NUMLOOP" value="0">
        <param name="CENTER" value="0">
        <param name="MAINTAINASPECT" value="0">
        <param name="BACKGROUNDCOLOR" value="#000000">
      </object>
 </td>
   <td><img src="images/spacer.gif" width="1" height="178" border="0" alt=""></td>
  </tr>
  <tr>
   <td><img name="HeiRui_r7_c2" src="images/HeiRui_r7_c2.gif" width="219" height="6" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="6" border="0" alt=""></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#000000">
<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
 codebase="http://active.macromedia.com/flash2/cabs/swflash.cab#version=4,0,0,0" WIDTH=219 height=70>
  <PARAM NAME=movie VALUE="images/elinedj.swf" ref> 
  <PARAM NAME=quality VALUE=High>
  <param name="_cx" value="5556">
  <param name="_cy" value="4710">
  <param name="FlashVars" value="-1">
  <param name="Src" ref value="images/elinedj.swf">
  <param name="WMode" value="Window">
  <param name="Play" value="-1">
  <param name="Loop" value="-1">
  <param name="SAlign" value>
  <param name="Menu" value="-1">
  <param name="Base" value>
  <param name="AllowScriptAccess" value="always">
  <param name="Scale" value="ShowAll">
  <param name="DeviceFont" value="0">
  <param name="EmbedMovie" value="0">
  <param name="BGColor" value>
  <param name="SWRemote" value>
  <EMBED src="images/elinedj.swf" loop=false menu=false quality=high WIDTH=400 height=200 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED></OBJECT>
    </td>
   <td><img src="images/spacer.gif" width="1" height="74" border="0" alt=""></td>
  </tr>
  <tr>
    <td><img name="HeiRui_r9_c2" src="images/HeiRui_r9_c2.jpg" width="219" height="47" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="47" border="0" alt=""></td>
  </tr>
</table>
<%
rs.movenext
wend
rs.Close
set rs=nothing
end if
conn.close
set conn=nothing
%>