
<!--#include file="conn.asp"-->
<%
id=request("id")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

if request("id")<>"" then
'set rs=server.createobject("adodb.recordset")
conn.execute("Sp_MusicListHits @MusicListID="&cstr(id) )

    set rs=server.createobject("adodb.recordset")
    id=request("id")
    sql="select * from MusicList where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
mm=rs("Wma")
rs.movenext
wend
conn.close
set conn=nothing
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>一线视听</title>
<style type="text/css">  
A{text-transform: none; text-decoration:none;color:#0000FF}
a:hover {text-decoration:underline;color:#ff0033}
.new {  border-left:0px dotted; border-right:0px dotted #0F6B01; border-top:0px dotted; border-bottom:0px dotted; font-family:"宋体"; font-size: 10pt; vertical-align: 10%; line-height: 20px}
</style>
<SCRIPT>
function test()
{
if(event.ctrlKey) alert("一线网络 wWw.happyjh.com");
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" style="OVERFLOW: hidden; scroll: no"
 onkeydown="test()" oncontextmenu="return false" onselectstart="return false" ondragstart="return false" onclick="if(event.shiftKey&&event.srcElement.tagName=='A')return false;" background="image/33.jpg" >

<table border="0" cellpadding="0" cellspacing="0" width="469" height="270">
  <tr> 
    <td height="32">
      <p>　</p>
      <p>　</p>
      <p>　</p>
    </td>
  </tr>
  <tr> 
    <td></td>
  </tr>
  <tr> 
    <td> <br>
      <table width="469" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="208"> 
            <table width="147" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="48">　</td>
                <td width="99">
				
				<MARQUEE onmouseover=this.stop() onmouseout=this.start() 
                  scrollAmount=1 scrollDelay=50 direction=up height=120>&nbsp; <span class="new"><a href="http://windowsmedia.com/download" target=_blank>　　本站音乐文件均为WMA格式,须安装Windows&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Media Player方可正常听歌,如果您尚未安装,请点此下载安装后即可正常听歌!</a></span></marquee>
				
				</td>
              </tr>
            </table>
            <p>　</p>
            <p>　</p>
            <p>　</p>
            <p>　</p>
            <p>　</p>
          </td>
          <td width="261" valign="middle"> 
            <p><br>
              <br>
<script language="javascript">
<!--
var str
str="http://scat.ting88.com/ting88/<%=mm%>";

document.write ("<OBJECT id=Player2 name=Player classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 border=\"0\" width=100% height=70 type=application/x-oleobject standby=\"Loading Windows Media Player components...\">\n")
document.write ("<param name=\"AudioStream\" value=\"-1\">\n")
document.write ("<param name=\"AutoSize\" value=\"0\">\n")
document.write ("<param name=\"AutoStart\" value=\"-1\">\n")
document.write ("<param name=\"AnimationAtStart\" value=\"-1\">\n")
document.write ("<param name=\"AllowScan\" value=\"-1\">\n")
document.write ("<param name=\"AllowChangeDisplaySize\" value=\"-1\">\n")
document.write ("<param name=\"AutoRewind\" value=\"0\">\n")
document.write ("<param name=\"Balance\" value=\"10\">\n")
document.write ("<param name=\"BaseURL\" value>\n")
document.write ("<param name=\"BufferingTime\" value=\"5\">\n")
document.write ("<param name=\"CaptioningID\" value>\n")
document.write ("<param name=\"ClickToPlay\" value=\"-1\">\n")
document.write ("<param name=\"CursorType\" value=\"0\">\n")
document.write ("<param name=\"CurrentPosition\" value=\"-1\">\n")
document.write ("<param name=\"CurrentMarker\" value=\"0\">\n")
document.write ("<param name=\"DefaultFrame\" value>\n")
document.write ("<param name=\"DisplayBackColor\" value=\"0\">\n")
document.write ("<param name=\"DisplayForeColor\" value=\"16777215\">\n")
document.write ("<param name=\"DisplayMode\" value=\"0\">\n")
document.write ("<param name=\"DisplaySize\" value=\"4\">\n")
document.write ("<param name=\"Enabled\" value=\"-1\">\n")
document.write ("<param name=\"EnableContextMenu\" value=\"0\">\n")
document.write ("<param name=\"EnablePositionControls\" value=\"-1\">\n")
document.write ("<param name=\"EnableFullScreenControls\" value=\"0\">\n")
document.write ("<param name=\"EnableTracker\" value=\"-1\">\n")
document.write ("<param name=\"Filename\" value='" + str + "'>\n")
document.write ("<param name=\"InvokeURLs\" value=\"-1\">\n")
document.write ("<param name=\"Language\" value=\"-1\">\n")
document.write ("<param name=\"Mute\" value=\"0\">\n")
document.write ("<param name=\"PlayCount\" value=\"0\">\n")
document.write ("<param name=\"PreviewMode\" value=\"0\">\n")
document.write ("<param name=\"Rate\" value=\"1\">\n")
document.write ("<param name=\"SAMILang\" value>\n")
document.write ("<param name=\"SAMIStyle\" value>\n")
document.write ("<param name=\"SAMIFileName\" value>\n")
document.write ("<param name=\"SelectionStart\" value=\"0\">\n")
document.write ("<param name=\"SelectionEnd\" value=\"0\">\n")
document.write ("<param name=\"SendOpenStateChangeEvents\" value=\"-1\">\n")
document.write ("<param name=\"SendWarningEvents\" value=\"-1\">\n")
document.write ("<param name=\"SendErrorEvents\" value=\"-1\">\n")
document.write ("<param name=\"SendKeyboardEvents\" value=\"0\">\n")
document.write ("<param name=\"SendMouseClickEvents\" value=\"0\">\n")
document.write ("<param name=\"SendMouseMoveEvents\" value=\"0\">\n")
document.write ("<param name=\"SendPlayStateChangeEvents\" value=\"-1\">\n")
document.write ("<param name=\"ShowCaptioning\" value=\"0\">\n")
document.write ("<param name=\"ShowControls\" value=\"-1\">\n")
document.write ("<param name=\"ShowAudioControls\" value=\"-1\">\n")
document.write ("<param name=\"ShowDisplay\" value=\"0\">\n")
document.write ("<param name=\"ShowGotoBar\" value=\"0\">\n")
document.write ("<param name=\"ShowPositionControls\" value=\"-1\">\n")
document.write ("<param name=\"ShowStatusBar\" value=\"-1\">\n")
document.write ("<param name=\"ShowTracker\" value=\"-1\">\n")
document.write ("<param name=\"TransparentAtStart\" value=\"0\">\n")
document.write ("<param name=\"VideoBorderWidth\" value=\"0\">\n")
document.write ("<param name=\"VideoBorderColor\" value=\"0\">\n")
document.write ("<param name=\"VideoBorder3D\" value=\"0\">\n")
document.write ("<param name=\"Volume\" value=\"0\">\n")
document.write ("<param name=\"WindowlessVideo\" value=\"0\">\n")
document.write ("<embed type=\"application/x-mplayer2\" pluginspage=\"http://www.microsoft.com/windows/mediaplayer/download/default.asp\" Name=\"Player\" width=\"100%\" height=\"70\" border=\"0\" SHOWSTATUSBAR=\"-1\" SHOWCONTROLS=\"0\" SHOWGOTOBAR=\"0\" SHOWDISPLAY=\"-1\" INVOKEURLS=\"-1\" AUTOSTART=\"1\" CLICKTOPLAY=\"0\" DisplayForeColor=\"12945678\">\n")
document.write ("</OBJECT>\n")
//-->
            </script>
          </td>
        </tr>
      </table>
      <br>
      <br>
      <br>
    </td>
  </tr>
  <tr> 
    <td></td>
  </tr>
</table>         
</body>        
</html>