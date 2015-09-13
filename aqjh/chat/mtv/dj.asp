<%@ LANGUAGE=VBScript.Encode codepage ="936" %>
<%Response.Expires=0
if session("aqjh_name")="" then response.end
song=Request.QueryString("song")
'response.write "<center><bgsound src="&song&" loop=-1 width=150 height=60 autostart=true></center>"
%>
<title>øÏ¿÷µ„∏ËÃ®</title><body background=../bg.gif oncontextmenu=self.event.returnValue=false>
<center>
<object id="mPlayer1" width=200 height=60 classid="CLSID:6BF52A52-394A-11D3-B153-00C04F79FAA6" type=application/x-oleobject standby="Loading Windows Media Player components...">
<param name="URL" value="<%=song%>">
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