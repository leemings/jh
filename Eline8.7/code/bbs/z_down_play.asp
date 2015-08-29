<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp"-->
<!--#include file="inc/FORUM_CSS.asp"-->
<%dim ldown,URL
set rs=server.createobject("adodb.recordset")
sql="select * from download where id="&request("id")         
rs.open sql,conndown,1,1
ldown=rs("ldown")
if request("downid")="1" then
	if ldown=1 then
		URL=downpath&rs("filename")
	else
		URL=rs("filename")
	end if
elseif request("downid")="2" then
	URL=rs("filename1")
elseif request("downid")="3" then
	URL=rs("filename2")
end if
dim LastHits
lasthits=rs("lasthits")
dim tdate
tdate=year(Now()) & "-" & month(Now()) & "-" & day(Now())
if datediff("d",lasthits,date())=0 then
	sql="update download set dayhits=dayhits+1 where id="&request("id")
	conndown.Execute(sql)
else
	sql="update download set dayhits=1 where id="&request("id")
	conndown.Execute(sql)
end if
sql="update download set hits=hits+1,lasthits='"&tdate&"' where ID="&request("id")
conndown.Execute(sql)
if datediff("d",lasthits,date())<=7 then
	sql="update download set weekhits=weekhits+1 where id="&request("id")
	conndown.Execute(sql)
else
	sql="update download set weekhits=1 where id="&request("id")
	conndown.Execute(sql)
end if
rs.close
set rs=nothing
if request("action")="rm" then
	call playrm()
elseif request("action")="mpeg" then
	call plaympeg()
elseif request("action")="swf" then
	call playswf()
elseif request("action")="mp3" then
	call playrm()
end if

sub plaympeg()%>
	<script language="JavaScript">
		if (window.Event) document.captureEvents(Event.MOUSEUP); 
		
		function nocontextmenu() {
			event.cancelBubble = true
			event.returnValue = false;
			return false;
		}
		 
		function norightclick(e) {
			if (window.Event) {
				if (e.which == 2 || e.which == 3)
				return false;
			}
			else
				if (event.button == 2 || event.button == 3)	{
					event.cancelBubble = true
					event.returnValue = false;
					return false;
				}
		}
		 
		document.oncontextmenu = nocontextmenu;  // for IE5+
		document.onmousedown = norightclick;  // for all others
	
		function runClock() {
			theTime = window.setTimeout("runClock()", 1000);
			var today = new Date();
			var display= today.toLocaleString();
			status=display;
		}
	</script>
	<table border="0" width="470" cellspacing="0" cellpadding="0" >
		<tr>
			<td width="100%" valign="top">
				<object classid="clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" id="MediaPlayer1" width="450" height="380">
				<param name="AudioStream" value="-1">
				<param name="AutoSize" value="0">
				<param name="AutoStart" value="-1">
				<param name="AnimationAtStart" value="-1">
				<param name="AllowScan" value="-1">
				<param name="AllowChangeDisplaySize" value="-1">
				<param name="AutoRewind" value="0">
				<param name="Balance" value="0">
				<param name="BaseURL" value>
				<param name="BufferingTime" value="5">
				<param name="CaptioningID" value>
				<param name="ClickToPlay" value="-1">
				<param name="CursorType" value="0">
				<param name="CurrentPosition" value="-1">
				<param name="CurrentMarker" value="0">
				<param name="DefaultFrame" value>
				<param name="DisplayBackColor" value="0">
				<param name="DisplayForeColor" value="16777215">
				<param name="DisplayMode" value="0">
				<param name="DisplaySize" value="2">
				<param name="Enabled" value="-1">
				<param name="EnableContextMenu" value="-1">
				<param name="EnablePositionControls" value="-1">
				<param name="EnableFullScreenControls" value="0">
				<param name="EnableTracker" value="-1">
				<param name="Filename" value="<%=url%>">
				<param name="InvokeURLs" value="-1">
				<param name="Language" value="-1">
				<param name="Mute" value="0">
				<param name="PlayCount" value="1">
				<param name="PreviewMode" value="0">
				<param name="Rate" value="1">
				<param name="SAMILang" value>
				<param name="SAMIStyle" value>
				<param name="SAMIFileName" value>
				<param name="SelectionStart" value="-1">
				<param name="SelectionEnd" value="-1">
				<param name="SendOpenStateChangeEvents" value="-1">
				<param name="SendWarningEvents" value="-1">
				<param name="SendErrorEvents" value="-1">
				<param name="SendKeyboardEvents" value="0">
				<param name="SendMouseClickEvents" value="0">
				<param name="SendMouseMoveEvents" value="0">
				<param name="SendPlayStateChangeEvents" value="-1">
				<param name="ShowCaptioning" value="0">
				<param name="ShowControls" value="-1">
				<param name="ShowAudioControls" value="-1">
				<param name="ShowDisplay" value="0">
				<param name="ShowGotoBar" value="0">
				<param name="ShowPositionControls" value="-1">
				<param name="ShowStatusBar" value="-1">
				<param name="ShowTracker" value="-1">
				<param name="TransparentAtStart" value="0">
				<param name="VideoBorderWidth" value="0">
				<param name="VideoBorderColor" value="0">
				<param name="VideoBorder3D" value="0">
				<param name="Volume" value="-40">
				<param name="WindowlessVideo" value="0">
				</object>
			</td>
		</tr>
	</table>
<%end sub

sub playrm()%>
	<script language="JavaScript">
		if (window.Event) document.captureEvents(Event.MOUSEUP); 
		
		function nocontextmenu() {
			event.cancelBubble = true
			event.returnValue = false;
			return false;
		}
		 
		function norightclick(e) {
			if (window.Event) {
				if (e.which == 2 || e.which == 3)
				return false;
			}
			else
				if (event.button == 2 || event.button == 3)	{
					event.cancelBubble = true
					event.returnValue = false;
					return false;
				}
		}
		 
		document.oncontextmenu = nocontextmenu;  // for IE5+
		document.onmousedown = norightclick;  // for all others
	
		function runClock() {
			theTime = window.setTimeout("runClock()", 1000);
			var today = new Date();
			var display= today.toLocaleString();
			status=display;
		}
	</script>
	<table border="0" width="470" cellspacing="1" cellpadding="0" >
		<tr>
			<td width="100%" valign="top">
				<object ID="video2" CLASSID="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" HEIGHT="340" WIDTH="450">
				<param name="_ExtentX" value="11906">
				<param name="_ExtentY" value="8996">
				<param name="AUTOSTART" value="-1">
				<param name="SHUFFLE" value="0">
				<param name="PREFETCH" value="0">
				<param name="NOLABELS" value="0">
				<param name="SRC" value="<%=url%>">
				<param name="CONTROLS" value="ImageWindow">
				<param name="CONSOLE" value="Clip1">
				<param name="LOOP" value="0">
				<param name="NUMLOOP" value="0">
				<param name="CENTER" value="0">
				<param name="MAINTAINASPECT" value="0">
				<param name="BACKGROUNDCOLOR" value="#000000">
				<embed SRC="4.rpm" type="audio/x-pn-realaudio-plugin" CONSOLE="Clip1" CONTROLS="ImageWindow" HEIGHT="240" WIDTH="352" AUTOSTART="false">
				</object>
				<object ID="video1" CLASSID="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" HEIGHT="60" WIDTH="450">
				<param name="_ExtentX" value="11906">
				<param name="_ExtentY" value="1588">
				<param name="AUTOSTART" value="-1">
				<param name="SHUFFLE" value="0">
				<param name="PREFETCH" value="0">
				<param name="NOLABELS" value="0">
				<param name="CONTROLS" value="ControlPanel,StatusBar">
				<param name="CONSOLE" value="Clip1">
				<param name="LOOP" value="0">
				<param name="NUMLOOP" value="0">
				<param name="CENTER" value="0">
				<param name="MAINTAINASPECT" value="0">
				<param name="BACKGROUNDCOLOR" value="#000000">
				<embed type="audio/x-pn-realaudio-plugin" CONSOLE="Clip1" CONTROLS="ControlPanel,StatusBar" HEIGHT="60" WIDTH="275" AUTOSTART="false">
				</object>
			</td>        
		</tr>
	</table>
<%end sub

sub playswf()%>
	<script language="JavaScript">
		if (window.Event) document.captureEvents(Event.MOUSEUP); 
		
		function nocontextmenu() {
			event.cancelBubble = true
			event.returnValue = false;
			return false;
		}
		 
		function norightclick(e) {
			if (window.Event) {
				if (e.which == 2 || e.which == 3)
				return false;
			}
			else
				if (event.button == 2 || event.button == 3)	{
					event.cancelBubble = true
					event.returnValue = false;
					return false;
				}
		}
		 
		document.oncontextmenu = nocontextmenu;  // for IE5+
		document.onmousedown = norightclick;  // for all others
	
		function runClock() {
			theTime = window.setTimeout("runClock()", 1000);
			var today = new Date();
			var display= today.toLocaleString();
			status=display;
		}
	</script>
	<table border="0" cellspacing="1" cellpadding="0" >
		<tr>
			<td width="100%" valign="top">
				<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 width=440 height=320 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000><PARAM NAME=movie VALUE="<%=url%>">
				<PARAM NAME=quality VALUE=high>
				<embed src="<%=url%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width=440 height=320><%=url%>
				</OBJECT>
			</td>
		</tr>
	</table>
<%end sub%>
