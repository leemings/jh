<%
response.buffer=true
dd=trim(cstr(session("players")))
application("com"&dd)=""
response.cookies("endtime")("oneplay")="no"
response.cookies("endtiom").Expires =Date( )+2
onep=request.cookies("endtime")("oneplay")
'response.write onep
%>

<html>
<head>
<title>E线江湖--游戏在线</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="forum.css">
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</head>

<body bgcolor="#FFFFFF" background="gif/msgbg.gif">
<a href=endclock.asp>clock</a> 
<table width="100%" border="0" cellspacing="0" bgcolor="#AB8EFF" align="center">
  <tr> 
    <td>
      <table width="100%" border="0" cellspacing="0" background="gif/msgbg.gif" align="center">
        <tr>
          <td>　</td>
        </tr>
        <tr> 
          <td>　</td>
        </tr>
        <tr> 
          <td> 
            <div align="left">谢谢对本游戏的支持</div>
          </td>
        </tr>
        <tr> 
          <td>游戏好玩人人都有机会</td>
        </tr>
        <tr> 
          <td>你有什么好的建议<a href="mailto:Eline_Email@etang.com">请告诉我</a></td>
        </tr>
        <tr> 
          <td>　</td>
        </tr>
        <tr> 
          <td> 
            <div align="left"></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <div align="left"></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>