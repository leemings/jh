<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if aqjh_grade<7 then
	Response.Write "<script Language=Javascript>alert('管理等级不够！需要九级以上!');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<html>
<head>
<title>官府放怪</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:#ccccff;text-decoration:none;}
a:hover{color:#ffffff;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed">
<div align="center"><br><table width="70%" border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="#" onClick="window.open('listbak.asp','listbak','scrollbars=yes,resizable=yes,width=400,height=250')">发钱</a></div>
      </td>
    </tr>    <tr > 
      <td id="sm"> 
        <div align="center"><a href="sendnl.asp" target="_blank">内力</a></div>
      </td>
    </tr><tr> 
      <td id="sm"> 
        <div align="center"><a href="sendfy.asp" target="_blank">防御</a></div>
      </td>
    </tr><tr> 
      <td id="sm"> 
        <div align="center"><a href="sendtl.asp" target="_blank">体力</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=老虎&value=10" target="d" title="放出老虎，让大家打!">老虎</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=衰哥&value=10000" target="d" title="发放衰哥，很恐怖，最大体力1万点!">衰哥</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=僵尸&value=10000" target="d" title="发放僵尸，很恐怖，最大体力1万点!">僵尸</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=手枪&value=10000" target="d" title="发放手枪，要小心!">手枪</a></font></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=狗狗&value=50000" target="d" title="发放狗狗，最大体力5万点!">狗狗</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=饿鼠&value=20000" target="d" title="放出饿鼠，让大家打!">饿鼠</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=飞鸟&value=50000" target="d" title="发放飞鸟，最大体力5万点!">飞鸟</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=武士&value=50000" target="d" title="发放武士，最大体力5万点!">武士</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=八国联军&value=50000" target="d" title="八国联军入侵中国，最大人数5万人!">联军</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=土匪&value=50000" target="d" title="发放土匪，最大人数5万人!">土匪</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=舞女&value=50000" target="d" title="发放舞女，最大体力5万点!">舞女</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=舞男&value=50000" target="d" title="发放舞男，最大体力5万点!">舞男</a></font></div></td>
    </tr>
	<tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=流氓&value=100000" target="d" title="发放猪哥，很恐怖，最大体力10万点!">流氓</a></font></div></td>
    </tr>
 <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=苹果&value=10000" target="d" title="苹果掉下来了!">苹果</a></font></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=烟花&value=50000" target="d" title="发放烟花，最大体力5万点!">烟花</a></font></div></td>
    </tr>
	<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=狮子&value=50000" target="d" title="发放狮子，最大体力5万点!">狮子</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=非典&value=50000" target="d" title="非典病毒5万点!">非典</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=人贩&value=50000" target="d" title="人贩子体力5万点!">人贩</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=小偷&value=50000" target="d" title="小偷体力5万点!">小偷</a></font></div></td>
    </tr>
  </table>
</div>
</body></html>