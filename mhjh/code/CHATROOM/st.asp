<html><head><title>图片</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+"[IMG]"+list+"[/IMG]";parent.f2.document.af.sytemp.focus();}</script>
<style TYPE="text/css">
td{line-height:110%;font-size:12px;}
body{ font-family: 宋体; font-size: 12px;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
.text {
	FONT-SIZE: 9pt; COLOR: #ffcc33; TEXT-DECORATION: none
}
.p9{line-height:150%;font-size:9pt;}
A{color:FFFFFF;text-decoration:none;}
A:Visited{color:FFFFFF;}
A:Active{color:FFFF00;}
A:Hover{color:Black;}
</style>
</head>
<body text="#FFFFFF" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false background="bg1.gif"> 
<SCRIPT language=JavaScript>
//ServerName
NetShowServer = "" ;
//Path
//var mPath = NetShowServer + "/entercp/mnet/";
var mPath = NetShowServer + "";
var bgImage = "bg.gif";
var iName = bgImage;

function LoadVideo(c1,c2) {
 if (c1=="") return;
 video_filename = c1;

 MusicPlayer.Cancel()
 timer=window.setTimeout("video_play()",2000)
}
function video_play(fn) {
 var mName = mPath + video_filename;
 MusicPlayer.Open(mName)
}
</SCRIPT>
<div align="center">
  <center>
<table border="1" width="114%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" style="border-collapse: collapse" bordercolor="#111111">
<br>  <tr> 
        <td width="150%" height="28" valign="bottom" bgcolor="#0058C0"> <div align="center">
          <font color="#CCCCFF" style="font-size: 9pt"><strong>在 线 视 听</strong></font></div></td>
      </tr>
    </table>
  </center>
</div>
<table width="99%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#000000"
bordercolordark="#FFFFFF">
  <tr> 
    <td height="15" colspan="2" align="center" bgcolor="#0058C0"> 
      <div align=center class=p9>
        <OBJECT id=MusicPlayer style="LEFT: 0px; TOP: 0px" height=180 width=140 
classid=clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95>
          <PARAM NAME="AudioStream" VALUE="-1"><PARAM NAME="AutoSize" VALUE="0"><PARAM NAME="AutoStart" VALUE="-1"><PARAM NAME="AnimationAtStart" VALUE="-1"><PARAM NAME="AllowScan" VALUE="-1"><PARAM NAME="AllowChangeDisplaySize" VALUE="-1"><PARAM NAME="AutoRewind" VALUE="-1"><PARAM NAME="Balance" VALUE="0"><PARAM NAME="BaseURL" VALUE=""><PARAM NAME="BufferingTime" VALUE="5"><PARAM NAME="CaptioningID" VALUE=""><PARAM NAME="ClickToPlay" VALUE="-1"><PARAM NAME="CursorType" VALUE="-1"><PARAM NAME="CurrentPosition" VALUE="-1"><PARAM NAME="CurrentMarker" VALUE="0"><PARAM NAME="DefaultFrame" VALUE=""><PARAM NAME="DisplayBackColor" VALUE="0"><PARAM NAME="DisplayForeColor" VALUE="16777215"><PARAM NAME="DisplayMode" VALUE="0"><PARAM NAME="DisplaySize" VALUE="2"><PARAM NAME="Enabled" VALUE="-1"><PARAM NAME="EnableContextMenu" VALUE="-1"><PARAM NAME="EnablePositionControls" VALUE="-1"><PARAM NAME="EnableFullScreenControls" VALUE="0"><PARAM NAME="EnableTracker" VALUE="-1"><PARAM NAME="Filename" VALUE=""><PARAM NAME="InvokeURLs" VALUE="-1"><PARAM NAME="Language" VALUE="-1"><PARAM NAME="Mute" VALUE="0"><PARAM NAME="PlayCount" VALUE="1"><PARAM NAME="PreviewMode" VALUE="0"><PARAM NAME="Rate" VALUE="1"><PARAM NAME="SAMILang" VALUE=""><PARAM NAME="SAMIStyle" VALUE=""><PARAM NAME="SAMIFileName" VALUE=""><PARAM NAME="SelectionStart" VALUE="-1"><PARAM NAME="SelectionEnd" VALUE="-1"><PARAM NAME="SendOpenStateChangeEvents" VALUE="-1"><PARAM NAME="SendWarningEvents" VALUE="-1"><PARAM NAME="SendErrorEvents" VALUE="-1"><PARAM NAME="SendKeyboardEvents" VALUE="0"><PARAM NAME="SendMouseClickEvents" VALUE="-1"><PARAM NAME="SendMouseMoveEvents" VALUE="0"><PARAM NAME="SendPlayStateChangeEvents" VALUE="-1"><PARAM NAME="ShowCaptioning" VALUE="0"><PARAM NAME="ShowControls" VALUE="-1"><PARAM NAME="ShowAudioControls" VALUE="-1"><PARAM NAME="ShowDisplay" VALUE="0"><PARAM NAME="ShowGotoBar" VALUE="0"><PARAM NAME="ShowPositionControls" VALUE="0"><PARAM NAME="ShowStatusBar" VALUE="-1"><PARAM NAME="ShowTracker" VALUE="-1"><PARAM NAME="TransparentAtStart" VALUE="0"><PARAM NAME="VideoBorderWidth" VALUE="0"><PARAM NAME="VideoBorderColor" VALUE="0"><PARAM NAME="VideoBorder3D" VALUE="0"><PARAM NAME="Volume" VALUE="-80"><PARAM NAME="WindowlessVideo" VALUE="0"></OBJECT></div></td>
  </tr>
<tr>
    <td width="136" align="center" bgcolor="#0058C0">
<SELECT 
style="FONT-SIZE: 9pt; COLOR: blue; BACKGROUND-COLOR: #ffcc33" 
onchange=LoadVideo(this.value) size=1 name=Sel> <OPTION value="" 
  selected>请选择网络电台<OPTION
  value=mms://202.102.249.117/wytonline>郑州文艺台<OPTION
  value=mms://202.130.232.145:9999>北京文艺台<OPTION
  value=mms://movie.jxgdw.com/jxdt>江西电台<OPTION
  value=mms://202.175.80.10/LiveAudio>澳门电台<OPTION
  value=http://metromedia.997metroshowbiz.com/metro/hit997.asx>新城PLUS<OPTION
  value=mms://live.sys.hinet.net/heartradio-100>真心之音<OPTION
  value=http://203.187.31.160/FM917>台北之音</OPTION></SELECT>
	</td>
  </tr>
  <tr>
    <td width="136" align="center" bgcolor="#0058C0">
<SELECT 
style="FONT-SIZE: 9pt; COLOR: blue; BACKGROUND-COLOR: #ffcc33" 
onchange=LoadVideo(this.value) size=1 name=Sel> <OPTION value="" 
  selected>请选择网络电视<OPTION value=mms://61.155.107.164/jsws>江苏卫视<OPTION 
  value=mms://61.155.107.164/city>江苏城市频道<OPTION 
  value=mms://61.155.107.164/video1>江苏综艺频道<OPTION 
  value=mms://61.138.179.3/tv>长春卫视<OPTION 
  value=mms://movie.top86.com/qtv>青岛卫视<OPTION 
  value=mms://tv.sdtv.com.cn/sdtv>山东电视台<OPTION 
  value=mms://61.136.113.41/nytv_1>南阳电视台<OPTION 
  value=mms://202.96.114.251/lstv>浙江丽水电视台<OPTION 
  value=mms://4.19.71.20/cctv4b>中央电视台四套<OPTION 
  value=mms://4.19.71.20/cctv9b>中央电视台九套<OPTION 
  value=mms://211.216.53.154/KTV>韩国KTV电视台<OPTION 
  value=mms://211.233.25.61/mv2>韩国音乐电视<OPTION 
  value=http://211.4.214.246:80/avexnettv-cm>日本MTV频道<OPTION 
  value=mms://impresstv-wmt.stream.co.jp/impresstv-im>日本NBS电视台<OPTION 
  value=mms://avexnettv.stream.co.jp/avexnettv-channel-a>日本A频道1<OPTION 
  value=mms://avexnettv.stream.co.jp/avexnettv-pv>日本A频道2<OPTION 
  value=mms://avexnettv.stream.co.jp/avexnettv-cm>日本A频道3<OPTION 
  value=mms://mp.video.t-online.de/fashiontv_lod300k>法国时尚电视台<OPTION 
  value=mms://61.150.31.148/fhws>凤凰卫视中文1<OPTION 
  value=mms://61.139.48.218/tv11>凤凰卫视中文2<OPTION 
  value=mms://media.phoenixtv.com.cn/liveylk>凤凰卫视资讯1<OPTION 
  value=mms://61.150.31.148/fhzx>凤凰卫视资讯2<OPTION 
  value=mms://61.150.31.148/fhdyt>凤凰卫视电影1<OPTION 
  value=mms://61.128.101.18/phoenixmovie>凤凰卫视电影2</OPTION></SELECT>
	</td>
  </tr>
  <tr>
    <td width="136" align="center" bgcolor="#0058C0"> 
      <INPUT 
style="FONT-SIZE: 9pt; COLOR: blue; BACKGROUND-COLOR: #ffcc33" type=file 
onchange=LoadVideo(this.value) size=6 name=textfield2></td>
  </tr>
      <tr>
    <td width="136" align="center" valign="middle" bgcolor="#0058C0"><br>
      [<a href="http://yx8.net/tv.asp" target="_blank" title="更多电视台"><font color="#00FF00"><strong>观看更多电视</strong></font></a>]
	        　</td>
  </tr>
    <tr>
    <td width="136" align="center" bgcolor="#0058C0"> 宽带观看效果最佳</td>
  </tr>
  </table>

</body></html>