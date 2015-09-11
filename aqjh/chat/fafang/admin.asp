<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn
if rs("grade")<10 then
   Response.Write "<script Language=Javascript>alert('你不是站长不能够操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<html>
<head>
<title>官府发放操作。</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:#ccccff;text-decoration:none;}
a:hover{color:#ffffff;text-decoration:none;CURSOR:url('1.cur');}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed">
<div align="center"><br><table width="70%" border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=银两&value=1000000" target="d" title="发放银两，最大值为10万两!">银两</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=元宝&value=300000" target="d" title="发放元宝，最大银两1万!">元宝</a></div>
      </td>
    </tr>
<tr >
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=道德&value=100" target="d" title="发放道德，最大值为100点!">道德</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=魅力&value=100" target="d" title="发放魅力，最大值为100点!">魅力</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=法力&value=100" target="d" title="发放法力，最大值为100点!">法力</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=轻功&value=100" target="d" title="发放轻功，最大值为100点!">轻功</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=金&value=50" target="d" title="发放金属性，最大值为50点!">金属性</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=木&value=50" target="d" title="发放木属性，最大值为50点!">木属性</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=水&value=50" target="d" title="发放水属性，最大值为50点!">水属性</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=火&value=50" target="d" title="发放火属性，最大值为50点!">火属性</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=土&value=50" target="d" title="发放土属性，最大值为50点!">土属性</a></div>
      </td>
    </tr>
	<tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=蛋糕&value=100000" target="d" title="发放蛋糕，最大加10万体力!">生日</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=内力&value=5000" target="d" title="发放内力，最高发放5千点!">内力</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=体力&value=5000" target="d" title="发放体力，最高发放5千点!">体力</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=武功&value=5000" target="d" title="发放武功，最高发放5千点!">武功</a></div>
      </td>
    </tr>
	
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=老虎&value=10" target="d" title="放出老虎，让大家打!">老虎</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=衰哥&value=10000" target="d" title="发放衰哥，很恐怖，最大体力1万点!">衰哥</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=僵尸&value=10000" target="d" title="发放僵尸，很恐怖，最大体力1万点!">僵尸</a></div>
      </td>
    </tr>
<tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=手枪&value=10000" target="d" title="发放手枪，要小心!">手枪</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=泡豆点数&value=100" target="d" title="发放豆点，最大值为100点!">豆点</a></div>
      </td>
    <tr> 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=抢豆点&value=50" target="d" title="抢豆点，最大数量50!">抢豆点</a></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=狗狗&value=50000" target="d" title="发放狗狗，最大体力5万点!">狗狗</a></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=饿鼠&value=20000" target="d" title="放出饿鼠，让大家打!">饿鼠</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=飞鸟&value=50000" target="d" title="发放飞鸟，最大体力5万点!">飞鸟</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=武士&value=50000" target="d" title="发放武士，最大体力5万点!">武士</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=八国联军&value=50000" target="d" title="八国联军入侵中国，最大人数5万人!">联军</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=土匪&value=50000" target="d" title="发放土匪，最大人数5万人!">土匪</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=舞女&value=50000" target="d" title="发放舞女，最大体力5万点!">舞女</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=舞男&value=50000" target="d" title="发放舞男，最大体力5万点!">舞男</a></div></td>
    </tr>
	<tr > 
      <td id="sm" height="22"> <div align="center"><a href="fafang.asp?cz=流氓&value=100000" target="d" title="发放猪哥，很恐怖，最大体力10万点!">流氓</a></div></td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=卡片&value=1" target="d" title="放卡片!">卡片</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=补药&value=1" target="d" title="放补药!">补药</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=配药&value=1" target="d" title="放配药!">配药</a></div>
      </td>
    </tr>
 <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=苹果&value=10000" target="d" title="苹果掉下来了!">苹果</a></div>
      </td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=礼物&value=1" target="d" title="放礼物!">礼物</a></div>
      </td>
    </tr><tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=月饼&value=1" target="d" title="中秋月饼">中秋月饼</a></div>
      </td>
    </tr>
 <tr > 
      <td id="sm" height="22"> <div align="center"><a href="fafang.asp?cz=别墅&value=50000" target="d" title="发放别墅，最大加5万体力!">别墅</a></div></td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=站长&value=90000" target="d" title="发放红袖，最大加10万!">站长</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=抢存点&value=50" target="d" title="抢泡点，最大数量50!">抢存点</a></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=烟花&value=50000" target="d" title="发放烟花，最大体力5万点!">烟花</a></div></td>
    </tr>
	<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=狮子&value=50000" target="d" title="发放狮子，最大体力5万点!">狮子</a></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=非典&value=50000" target="d" title="非典病毒5万点!">非典</a></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=人贩&value=50000" target="d" title="人贩子体力5万点!">人贩</a></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=小偷&value=50000" target="d" title="小偷体力5万点!">小偷</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=金币&value=10" target="d" title="发放金币，最大值为10个!">金币</a></div>
     
      </td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=抢金卡&value=5" target="d" title="发放会员金卡，最大值为5个!">抢金卡</a></div>
      </td>
    </tr>
 <tr>
      <td id="sm"> 
        <div align="center"><a href="sen.asp" target="d" title="发放等级">等级</a></div>
      </td>
    </tr>
 <tr>
      <td id="sm"> 
        <div align="center"><a href="#" onClick="window.open('zmbk.asp','listbak','scrollbars=yes,resizable=yes,width=400,height=250')">拨款</a></div>
      </td>
    </tr>
  </table></div></body></html>