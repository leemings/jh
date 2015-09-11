
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
%>
<html>
<head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../chat/f2.gif">

<title><%=Application("aqjh_chatroomname")%>会员充值系统</title>
</head>
<SCRIPT>document.write(unescape("%3Cscript%3Edocument.oncontextmenu%3Dnew%20Function%28%22event.returnValue%3Dfalse%22%29%0D%0Adocument.onselectstart%3Dnew%20Function%28%22event.returnValue%3Dfalse%22%29%3C/script%3E"))</SCRIPT>
<body bgcolor="#000000">
<table width="666" border="3" align="center" cellpadding="8" cellspacing="8" bordercolor="#6633CC">
  <tr>
    <td align="center" valign="middle"><script language="JavaScript">
function load()
{var name=navigator.appName
var vers=navigator.appVersion
if(name=="Netscape")
{window.location.reload()
}else
{history.go(0)
}}
</script>

<a href="javascript:load()">【刷新】</a> <img src="5.gif" width="120" height="185">会员转换系统 
      <img src="5-2.gif" width="120" height="185"><a href="javascript:history.go(-1)">【返回】</a> </td>
</tr>
<tr>
    <td align="center" valign="middle"><a href="czk.asp">购买充值卡</a> </td>
</tr>
</table>
<table width="666" border="3" align="center" cellspacing="3" cellpadding="3" bordercolor="#6633CC">
  <!--DWLayoutTable-->
  <tr> 
    <td width="317" height="214" align="center" valign="middle"><form method="POST" action="zhuan1.asp">
        <img src="1.gif"><img src="1-2.gif"> 
        <input type="submit" value="转换1级会员" name="B1" class="p9">
      </form></td>
    <td width="318" align="center" valign="middle"><form method="POST" action="zhuan2.asp">
        <img src="2.gif"><img src="2-2.gif"> 
        <input type="submit" value="转换2级会员" name="B1" class="p9">
      </form></td>
  </tr>
  <tr> 
    <td height="24" align="center" valign="middle"> <form method="POST" action="zhuan3.asp">
        <img src="3.gif"><img src="3-2.gif"> 
        <input type="submit" value="转换3级会员" name="B1" class="p9">
      </form></td>
    <td align="center" valign="middle"> <form method="POST" action="zhuan4.asp">
        <img src="4.gif"><img src="4-2.gif"> 
        <input type="submit" value="转换4级会员" name="B1" class="p9">
      </form></td>
  </tr>
  <tr> 
    <td height="96" colspan="2" align="center" valign="top"> <br>
      可以使用自己的金卡、金币、银两、充值卡换取会员（切忌要满足条件）<br>
      一级会员/月需要500元金卡、1000个金币、银两1亿。加存点31250到期后减掉31250 <br>
      二级会员/月需要1000元金卡、1500个金币、银两1亿。加存点90000到期后减掉90000<br>
      三级会员/月需要8000元金卡、20000个金币、银两1.5亿。加存点250000到期后减掉250000<br>
      四级会员/月需要10000元金卡、30000个金币、银两2亿。加存点490000到期后减掉490000<br>
      购买五级别会员，请与站长永不放弃联系，电话：13635716119 QQ：51726805<br> </td>
  </tr>
</table>
</body>
</html>
