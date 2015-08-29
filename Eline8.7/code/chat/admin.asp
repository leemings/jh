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
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：想干什么？！只有站长才可以操作！');window.close();}</script>"
	Response.End 
end if
if session("advtime")<>"" then
if session("advtime")>now()-12 then  Response.Redirect "../error.asp?id=492"
end if
session("advtime")=now()
%>
<html>
<head>
<title>站长发放操作♀一线网络→wWw.51eline.com♀</title>
<style>
body{
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#993300" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#FFFF00" size="2">站长发放</font></div></td>
      </tr>
    </table>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#9999CC" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr > 
      <td id="sm"> 
        <div align="center"><font color="#FFFFFF" size="-1"><a href="fafang.asp?cz=银两&value=10000000" target="d" title="发放银两，最大值为1000万两!">银两</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=金币&value=3" target="d" title="发放金币，最大值为3个!">金币</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=银币&value=10" target="d" title="发放银币，最大值为10个!">银币</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=道德&value=1000" target="d" title="发放道德，最大值为1000点!">道德</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=魅力&value=1000" target="d" title="发放魅力，最大值为1000点!">魅力</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=法力&value=1000" target="d" title="发放法力，最大值为1000点!">法力</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=轻功&value=1000" target="d" title="发放轻功，最大值为1000点!">轻功</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=狗狗&value=8000" target="d" title="放出狗狗，让大家打!">狗狗</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=饿鼠&value=20000" target="d" title="放出饿鼠，让大家打!">饿鼠</a></font></div>
      </td>
    </tr>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang.asp?cz=怪鸟&value=10000" target="d" title="发放怪鸟，最大体力5万点!">怪鸟</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=武士&value=50000" target="d" title="放出武士，让大家打!">武士</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=八国联军&value=80000" target="d" title="放出八国联军，让大家打!">联军</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=舞女&value=10000" target="d" title="放出舞女，让大家打!">舞女</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=舞男&value=10000" target="d" title="放出舞男，让大家打!">舞男</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang.asp?cz=流氓&value=30000" target="d" title="发放猪哥，很恐怖，最大体力3万点!">流氓</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=补药&value=5" target="d" title="发放补药，最大值为5个!">补药</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang.asp?cz=配药&value=1" target="d" title="放配药!">配药</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang.asp?cz=苹果&value=9000" target="d" title="发放苹果，最大加9000体力!">苹果</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang.asp?cz=别墅&value=50000" target="d" title="发放别墅，最大加5万体力!">别墅</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang.asp?cz=礼物&value=1" target="d" title="放礼物!">礼物</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=allvalue&value=100" target="d" title="发放存点，最大值为100点!">存点</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=泡豆点数&value=100" target="d" title="发放豆点，最大值为100点!">豆点</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=老虎&value=8000" target="d" title="放出老虎，让大家打!">老虎</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
      <div align="center"><font size="-1"><a href="fafang.asp?cz=手枪&value=8000" target="d" title="发放手枪，很恐怖，最大体力1万点!">手枪</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=僵尸&value=8000" target="d" title="发放僵尸，很恐怖，最大体力1万点!">僵尸</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=衰哥&value=8000" target="d" title="发放衰哥，很恐怖，最大体力1万点!">衰哥</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=元宝&value=100000" target="d" title="发放元宝，最大银两10万!">元宝</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=体力&value=5000" target="d" title="发放体力，最高发放5千点!">体力</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=内力&value=5000" target="d" title="发放内力，最高发放5千点!">内力</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=武功&value=5000" target="d" title="发放武功，最高发放5千点!">武功</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=抢存点&value=100" target="d" title="发放抢存点，最大值为100点!">抢存点</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=抢豆点&value=100" target="d" title="发放抢豆点，最大值为100点!">抢豆点</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=抢金卡&value=2" target="d" title="发放抢金卡，最大值为2元!">抢金卡</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=金&value=200" target="d" title="发放金，最大值为500点!">金属性</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=木&value=200" target="d" title="发放木，最大值为200点!">木属性</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=水&value=200" target="d" title="发放水，最大值为200点!">水属性</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=火&value=200" target="d" title="发放火，最大值为200点!">火属性</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=土&value=200" target="d" title="发放土，最大值为200点!">土属性</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm">
        <div align="center"><font color="#FF0000" size="-1">点击便发放.别放太多</font></div>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
