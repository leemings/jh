<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 门派 FROM 用户 where 姓名='" & sjjh_name &"'",conn
pai=rs("门派")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>夺宝特技</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (liji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style>
body{
CURSOR: url('3.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
a{font-size:9pt; color:#ccccff;text-decoration:none;}
a:hover{color:#ff0000;text-decoration:none;CURSOR:url('1.cur');}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" leftmargin="1" topmargin="1">
<script language="JavaScript">
if (navigator.appName.indexOf("Internet Explorer") != -1)	
document.onmousedown = noright;
function noright()
{
if (event.button == 2 | event.button == 3)
{
	alert("不要这样嘛 ^_^");
	 location.replace("#")
}
}
</script>
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#99CCCC" height="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#FF0000" size="2"><strong>夺宝特技</strong></font></div></td>
      </tr>
    </table>  
<table cellpadding="2" cellspacing="0" border="1" bgcolor="#66CCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr valign="middle"> 
    <td width="43%"> <div align="center"><a href="STUNTdb.ASP" target="f3"><font color="#990000" size="2">连 续 技</font></a> </div></td>
  </tr>
  <tr valign="middle"> 
    <td width="43%"> <div align="center"><a href="dgjjdb.asp" target="f3"><font color="#990000" size="2">独孤九剑</font></a> </div></td>
  </tr>
  <tr valign="middle"> 
    <td width="43%"> <div align="center"><a href="dbwupin.asp" onClick="javascript:s()" onMouseOver="window.status='看看你现在有哪些毒药';return true" onMouseOut="window.status='';return true" target="f3" title="看看有什么毒药"><font color="#990000" size="2">夺宝下毒</font></a> </div></td>
  </tr>
  <tr valign="middle"> 
    <td width="43%"> <div align="center"><a href="dbwupin.asp" onClick="javascript:s()" onMouseOver="window.status='看看现在有哪些暗器';return true" onMouseOut="window.status='';return true" target="f3" title="看看有什么暗器"><font color="#990000" size="2">夺宝投掷</font></a> </div></td>
  </tr>
  <tr valign="middle"> 
    <td width="43%"> <div align="center"><a href="BOY/dbBoy.asp" onClick="javascript:s()" target="f3" title="小孩夺宝"><font color="#990000" size="2">小孩夺宝</font></a> </div></td>
  </tr>
  <tr valign="middle"> 
    <td width="43%"> <div align="center"><a href="cw/dbcw.asp" onClick="javascript:s()" target=f3 title="宠物夺宝"><font color="#990000" size="2">宠物夺宝</font></a> </div></td>
  </tr>
</table>
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#66CCCC" height="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"><div align="center"><img src="go.gif" width="138"></div></td>
      </tr>
    </table>
</div>
</body>
</html>
