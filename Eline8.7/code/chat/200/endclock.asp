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
<title>���ֽ���--��Ϸ����</title>
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
          <td>��</td>
        </tr>
        <tr> 
          <td>��</td>
        </tr>
        <tr> 
          <td> 
            <div align="left">лл�Ա���Ϸ��֧��</div>
          </td>
        </tr>
        <tr> 
          <td>��Ϸ�������˶��л���</td>
        </tr>
        <tr> 
          <td>����ʲô�õĽ���<a href="mailto:Eline_Email@etang.com">�������</a></td>
        </tr>
        <tr> 
          <td>��</td>
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