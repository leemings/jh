<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then 
	Response.Write "<script Language=Javascript>top.location.href='http://www.happyjh.com';alert('提示：对不起，您还没有登陆江湖！');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：想听歌请先进入江湖聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 银两,银币 FROM 用户 WHERE 姓名='"& session("sjjh_name") &"'",conn,1,3
if rs("银两")<50000 or rs("银币")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "身上没50000银两或银币不足1个，不能听歌！"
	response.end
end if
rs("银两")=rs("银两")-50000
rs("银币")=rs("银币")-1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%PageName="Box"%>
<!--#include file="conn.asp"-->
<body leftmargin="0" topmargin="0" onLoad="drop(300)">
<%
if request("id")<>"" then
set rs=server.createobject("adodb.recordset")
id=request("id")
sql="select * from MusicList where id="&id
rs.open sql,conn,1,3
if rs.eof then
errmsg="<li>对不起！该歌曲不存在，可能已经被管理员删除。</li>"
call error()
Response.End 
else
rs("hits")=rs("hits")+1
rs.Update
id=rs("id")
Wma=rs("Wma")
hits=rs("hits")
Singer=rs("Singer")
MusicName=rs("MusicName")
end if
rs.Close

else
errmsg= "<li>请选择歌曲！</li>"
call error()
Response.End 
end if

set rs=nothing
conn.close
set conn=nothing
%><html>
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>一线视听 ← wWw.happyjh.com</title>
<style type="text/css">

 <!--

 td {  font-size: 9pt; color: #000000}

 a:link {  font-size: 9pt; color:#000000; text-decoration: none}

 a:active {  font-size: 9pt; ; text-decoration: underline}

 a:visited {  font-size: 9pt; color: #000000; text-decoration: none}

 a:hover {  font-size: 9pt; color:#000000; text-decoration: underline; background-color: #00FFFF}

 -->

 </style>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0"><BODY bgColor=#000000 leftMargin=0 topMargin=0 oncontextmenu="return false" ondragstart="return false" onselectstart="return false"> 
<table width="300" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><img src="files/play.jpg" width="300" height="13"></td>
  </tr>
  <tr> 
    <td><object id="mPlayer1" width=300 height=250
 classid="CLSID:6BF52A52-394A-11D3-B153-00C04F79FAA6" type=application/x-oleobject standby="Loading Windows Media Player components...">
        <param name="URL" value="Yxwma.asp?id=<%=id%>">
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
                        </object></td>
  </tr>
  <tr> 
    <td><table width="300" height="32" border="0" cellpadding="2" cellspacing="0">
        <tr>
          <td width="180"><marquee>
                        <font color="#FFFFFF"><b>本站音乐全为WMA格式，如不能正常播放，请下载升级播放程序，方可收听</b></font></marquee></td>
          <td width="120"><a href="http://windowsmedia.com/download" target="_blank"><img src="images/WMP9series_Free.gif" alt="点击进入下载站点" width="120" height="32" border="0"></a></td>
        </tr>
      </table></td>
  </tr>
</table>
</BODY></HTML>
<script>
      function rf()
      {return false; }
    document.oncontextmenu = rf
    function keydown()
      {if(event.ctrlKey ==true || event.keyCode ==93 || event.shiftKey ==true){return false;} }
      document.onkeydown =keydown
    function drag()
      {return false;}
    document.ondragstart=drag 
function stopmouse(e) { 
		if (navigator.appName == 'Netscape' && (e.which == 3 || e.which == 2))  
 		return false; 
      else if  
      (navigator.appName == 'Microsoft Internet Explorer' && (event.button == 2 || event.button == 3)) {  
 		alert("请不要盗用本站连接，你的IP已经被记录了");
		return false;  
 		}
		return true; 
      } 
      document.onmousedown=stopmouse;  
      if (document.layers) 
      window.captureEvents(Event.MOUSEDOWN); 
       window.onmousedown=stopmouse;  
      </script>
<script language="javascript">kstatus();function kstatus(){self.status="一线网络→ | wWw.happyjh.com | ←以后请大家使用这个域名访问本站.";setTimeout("kstatus()",0);}</script>