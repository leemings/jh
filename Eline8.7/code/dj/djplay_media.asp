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
	Response.Write "<script Language=Javascript>top.location.href='http://www.51eline.com';alert('��ʾ���Բ�������û�е�½������');</script>"
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
response.write("������������ѡ��Ҫ������������")
response.end
end if
if urgent<>"" then
response.write urgent
response.end
end if
if freelisten<>true then
response.write"<script>alert(""��ֹ�ο����裬��ע�ᣡ"");window.close();</script>"
response.end
end if
%>
<html>
<script>
songid="<%=songid%>"
if(window.name!='HeiRui_Studio_Player')
{
alert("�����������ſ������˱ư����ٿ������㣡������ģ�");top.location="";
}
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>E�߽�����վ��DJ��ɡ���������</title>
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
var ERR_NonePlayer="��������ʾ��:����ϵͳ��û�а�װReal Player��������������ҳ�����ء�\n\nMovie���������رա�";
var ERR_FileNotFind="��������ʾ��:�Բ���δ������Ҫ�㲥��������\n\n���������������";
var ERR_NotLocateServer="��������ʾ��:�޷���λMovie�����������������������";
var ERR_UnkownError="��������ʾ��:�����û����࣬���������ƣ������߻��Ժ�ۿ���";
         </script>
      <script LANGUAGE="VBScript">                                    
on error resume next
RealPlayerG2 = (NOT IsNull(CreateObject("rmocx.RealPlayer G2 Control")))\n'); 
RealPlayer5 = (NOT IsNull(CreateObject("RealPlayer.RealPlayer(tm) ActiveX Control (32-bit)")))
RealPlayer4 = (NOT IsNull(CreateObject("RealVideo.RealVideo(tm) ActiveX Control (32-bit)")))
if not RealPlayerG2 and RealPlayer5 and RealPlayer4 then
		if MsgBox("����������޷��Զ��������µ������������Ƿ�Ҫ���ز����������ţ�", vbYesNo) = vbYes then
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
<object id="mPlayer1" width=216 height=178 classid="CLSID:6BF52A52-394A-11D3-B153-00C04F79FAA6" type=application/x-oleobject standby="Loading Windows Media Player components...">                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         <PARAM NAME="URL" VALUE="<%=source%>">
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
</OBJECT>
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