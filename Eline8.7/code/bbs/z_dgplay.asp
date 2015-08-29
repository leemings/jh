<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'---------------------
' write by 绿水青山
'---------------------
	stats="正在播放..."&htmlencode(Request.QueryString("medianame"))

	call nav()
	call head_var(0,0,"欣赏点歌","#")

	if not founduser then
		response.redirect "login.asp"
	end if
	dim lurl
	lurl=Request.QueryString("url")

	select case Lcase(right(lurl,4))
		case ".swf"
%>		
			<p align=center>
			<OBJECT CLASSID=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 CODEBASE=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 WIDTH=720 HEIGHT=576>
				<PARAM NAME=movie VALUE=<%=lurl%>> <PARAM NAME=quality VALUE=high> <EMBED SRC=<%=lurl%> QUALITY=high PLUGINSPAGE=http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash TYPE=application/x-shockwave-flash WIDTH=720 HEIGHT=576></EMBED>
			</OBJECT>
<%	
		CASE ".mp3",".mid",".wmv",".mpg"	
%>
			<p align=center>
			<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=500 height=350 >
				<param name=ShowStatusBar value=-1>
				<param name=Filename value=<%=replace(lurl,chr(32),"%20",1)%>>
				<embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=<%=replace(lurl,chr(32),"%20",1)%>  width=500 height=100></embed>
			</object>
<%
		case else

			select case Lcase(right(lurl,3))
			  case ".rm"
%>			
				<p align=center>
				<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=500 height=350>
					<PARAM NAME=SRC VALUE=<%=lurl%>><PARAM NAME=CONSOLE VALUE=Clip1>
					<PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true>
				</OBJECT>
				<br>
				<OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=500>
					<PARAM NAME=SRC VALUE=<%=lurl%>><PARAM NAME=AUTOSTART VALUE=-1>
					<PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1>
				</OBJECT>
<%	
			 case ".ra"
%>			
				<p align=center>
				<OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=500>
					<PARAM NAME=SRC VALUE=<%=lurl%>><PARAM NAME=AUTOSTART VALUE=-1>
					<PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1>
				</OBJECT>
<%			 	
			 case else 
%>
				<p align=center>
				<embed name=player src=<%=lurl%> type=audio/x-pn-realaudio-plugin width=500 height=150 border=0 autostart=1></embed> 
<%
			end select 
	end select
	call footer()
%>