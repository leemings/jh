<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
call nav()
stats="在线测试"
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>您没有在本社区在线测试的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
call dvbbs_error()
else
response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>在线测试</b></td></tr><tr><td valign=middle class=tablebody1 height=100><CENTER>"
'本程序版权属于晨阳在线网站所有
%>
<html>

<head>
<title>在线测试</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<style>
<!--
BODY {SCROLLBAR-FACE-COLOR: #D6D6D6; SCROLLBAR-HIGHLIGHT-COLOR: #3A6EA5; SCROLLBAR-SHADOW-COLOR: #ffffff; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR:  #8FA5B6; SCROLLBAR-TRACK-COLOR: #f3f3f3; SCROLLBAR-DARKSHADOW-COLOR: #3A6EA5; }
-->
</style>

<%dim idd
idd=request.querystring("id")%>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="2" topmargin="2">
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="600" height="23" background="ontest/bg1.gif">
    <table border="0" cellspacing="0" width="600" id="AutoNumber2" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="25">
        <img border="0" src="ontest/err1.gif" width="24" height="23"></td>
        <td width="536" align="left">
        <table border="0" cellspacing="0" width="40%" id="AutoNumber5" cellpadding="0" height="15">
          <tr>
            <td width="100%" valign="bottom" height="15">&nbsp;&nbsp;
            <font color="#FF0000">
            <%if cint(idd)<100 then%>
            &nbsp;爱 情 测 试：
            <%elseif cint(idd)>100 and cint(idd)<200 then%>
            &nbsp;综 合 测 试：
            <%elseif cint(idd)>200 then%>
            &nbsp;个 人 测 试：
            <%end if%>            
            </font></td>
          </tr>
        </table>
        </td>
        <td width="39" valign="bottom" style="border-right: 1px solid #000000">
        <a href="javascript:history.go(-1)">
        <img border="0" src="ontest/help.gif" start="fileopen" width="15" height="18" align="absbottom"></a>
        <img border="0" src="ontest/close.gif" align="absbottom"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td height="562" valign="top" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-bottom: 1px solid #000000">
    <div align="center">
      <center>
      <table border="1" cellspacing="1" height="560" style="border-collapse: collapse; border-left: 1px solid #000000; border-right: 1px solid #000000; border-bottom: 1px solid #000000" width="596" id="AutoNumber3" bgcolor="#ECEDEF">
        <tr>
          <td width="100%" background="ontest/bg2.gif"> 
          <div align="center">
            <center>
            <table border="0" cellspacing="0" width="90%" id="AutoNumber6" cellpadding="0" height="140">
              <tr>
                <td width="100%" height="513">
                <iframe name="I1" src="ontest/<%=idd%>.html" height="100%" width="100%" border="0" frameborder="0" style="font-size: 9pt; font-family: 宋体; color: #000000; border: 1px dashed #008080; background-color: #EFF0F1" align="left"></iframe></td>
              </tr>
              <tr>
                <td width="100%" height="35" valign="bottom">
                <p align="center"><a href="javascript:history.go(-1)">
                <img border="0" src="ontest/goback.gif"></a></td>
              </tr>
            </table>
            </center>
          </div>
          </td>
        </tr>
      </table>
      </center>
    </div>
    </td>
  </tr>
</table>
</body>

</html>
<%
response.write "</CENTER></td></tr></table>"
end if
call footer()
%>