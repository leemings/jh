
<!--#include file="conn.asp"-->
<%
id=request("id")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Redirect "../error.asp?id=120"
end if

if request("id")<>"" then
conn.execute("Sp_MusicListHits @MusicListID="&cstr(id) )
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
.new {  font-family:"宋体"; font-size: 10pt; border-color: 0 #0F6B01 0 0; vertical-align: 10%; border-style: dotted; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 0px; border-left-width: 0px; line-height: 20px}
</style>
<SCRIPT>
function test()
{
if(event.ctrlKey) alert("一线网络 | HTTP://wWw.happyjh.com ");
}
</SCRIPT>


</head>
<body topmargin="0" leftmargin="0" style="OVERFLOW: hidden; scroll: no"
 onkeydown="test()" oncontextmenu="return false" onselectstart="return false" ondragstart="return false" oncontextmenu="window.event.returnValue=false" onclick="if(event.shiftKey&&event.srcElement.tagName=='A')return false;" background="image/33.jpg" >
<table border="0" cellpadding="0" cellspacing="0" width="469" height="270">
  <tr> 
    <td height="32">
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
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
                <td width="48">&nbsp;</td>
                <td width="99">
				
				<MARQUEE onmouseover=this.stop() onmouseout=this.start() 
                  scrollAmount=1 scrollDelay=50 direction=up height=120>
				  <span class="new"><a href="http://windowsmedia.com/download" target=_blank>　　本站音乐文件均为WMA格式,须安装Windows 
                  Media Player方可正常听歌,如果您尚未安装,请点此下载安装后即可正常听歌!</a></span>
				  </marquee>
				
				</td>
              </tr>
            </table>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
          </td>
          <td width="261" valign="middle"> 
            <p><br>
              <br>
              <object id="mPlayer1" width=243 height=64         
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
                <param name="volume" value="50">
                <param name="mute" value="0">
                <param name="uiMode" value="full">
                <param name="stretchToFit" value="0">
                <param name="windowlessVideo" value="0">
                <param name="enabled" value="-1">
                <param name="enableContextMenu" value="0">
                <param name="fullScreen" value="0">
                <param name="SAMIStyle" value="">
                <param name="SAMILang" value="">
                <param name="SAMIFilename" value="">
                <param name="captioningID" value="">
              </object><br>
            </p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp; </p>
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