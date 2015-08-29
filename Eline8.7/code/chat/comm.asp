<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
dim comm(30),commti(30),commsm(30)
comm(0)="/本门发钱$"
commti(0)="掌门，掌门夫人(老公)、长老给本门弟子发钱！"
commsm(0)="本门发钱"
comm(1)="/指导弟子$"
commti(1)="掌门，掌门夫人(老公)、长老给本门弟子加内力、体力！"
commsm(1)="指导弟子"
comm(2)="/信息$"
commti(2)="查看别人的信息，需要5万两银子的！"
commsm(2)="查看信息"
comm(3)="/赌博$ 消息"
commti(3)="查看赌场的消息、赌博情况！"
commsm(3)="赌场消息"
comm(4)="/幸运$ 幸运数字1-1000"
commti(4)="看看自己运气如何!"
commsm(4)="E线风采"
comm(4)="/幸运$ 幸运数字1-1000"
commti(4)="看看自己运气如何!"
commsm(4)="E线风采"
<option style="color:red" value="/求婚$ 真爱表白">示爱求婚
<option style="color:red" value="/基金$">本门基金
<% if Application("sjjh_baowu")>0 then%>
<option style="color:#FF0000" value="/修练$ ">宝物修练
<%end if%>
<%if chatinfo(5)=0 then%>
<option style="color:#FF0000" value="/暴豆$ ">暴威力豆
<%end if%>
<option style="color:green" value="/主题$ ">千里传音
<option style="color:green" value="/心动$ ">心动感觉
<option style="color:green" value="/怒吼$ ">狂狮怒吼
<option style="color:green" value="/心跳$ ">心跳心语
<option style="color:blue" value="/拜师$ ">拜师习武
<option style="color:blue" value="/收徒$ ">招收徒弟
<option style="color:blue" value="/打坐$ ">打坐练功
<option style="color:blue" value="/闭目$ ">闭目养神
<option style="color:blue" value="/练武$ ">武馆习武
<option style="color:orange" value="/传授$ 100">传授内力
<%if chatinfo(5)=0 then%>
<option style="color:orange" value="/吸星大法$ ">吸取内力
<%end if%>
<option style="color:orange" value="/经验$ 100">传授经验
<option style="color:black" value="/送钱$ 1000">赠送现金
<option style="color:black" value="/赠送$ 写出礼品名">赠送物品
<option style="color:black" value="/使用$ 卡片名">使用卡片
<%if chatinfo(5)=0 then%>
<option style="color:black" value="/偷钱$ ">偷取钱财
<option style="color:red" value="/下毒$ 毒药名">偷偷下毒
<option style="color:red" value="/投掷$ 暗器名">投掷暗器
<option style="color:red" value="/发招$ 招式名">发招攻击
<%end if%>
<option style="color:#B000D0" value="/存钱$ 0">银行存钱
<option style="color:#B000D0" value="/取钱$ 0">银行取钱
<option style="color:#B000D0" value="/转账$ 10000">银行转账
<option style="color:blue" value="/好友$ 加入">结义兄弟
<option style="color:blue" value="/好友$ 查看">查看兄弟
<option style="color:blue" value="/好友$ 删除">割袍段义
<option style="color:blue" value="/清除$ 要清除的好友姓名">清除兄弟
<%
if jhmp="" or jhmp="游侠" or jhmp="无" then%>
<option value="/加入门派$">加入门派
<%end if%>
<%if chatinfo(5)=0 then%>
	<%if left(jhzy,2)="小偷" then%>
		<option style="color:blue" value="/小偷$ ">偷取物品
	<%else%>
		<option style="color:blue" value="/抓小偷$ ">抓 小 偷
	<%end if %>
<%end if%>
<%if jhsf="掌门" then%>
<option value="/逐出弟子$">逐出弟子
<option style="color:blue" value="/掌门令$ 公布的消息">掌 门 令
<option value="/册封$ 身份名">册封身份
<option value="/册封$ 长老">册封长老
<option value="/册封$ 护法">册封护法
<option value="/册封$ 堂主">册封堂主
<option value="/册封$ 弟子">取消册封
<%end if
if (jhsf="长老" or jhsf="掌门") and sjjh_grade>=4 then%>
<option style="color:blue" value="/哑穴$ 请写明原因 ">教训弟子
<option style="color:blue" value="/奖励$ 100 ">奖励弟子
<%end if
if sjjh_grade>=6 then%>
<option style="color:blue" value="/哑穴$ 请写明原因 ">哑穴操作
<option style="color:blue" value="/点穴$ 请写明原因 ">点穴操作
<option style="color:blue" value="/解穴$ 用户名">解穴操作
<option style="color:blue" value="/警告$ 请写明原因 ">警告操作
<option style="color:blue" value="/逮捕$ 请写明原因 ">逮捕操作
<option style="color:blue" value="/坐牢$ 请写明原因 ">坐牢操作
<option style="color:blue" value="/查ip$ ">查看IP
<option style="color:blue" value="/踢人$ 请写明原因 ">让你捣乱
<option style="color:blue" value="/炸弹$">吃 炸 弹
<%
end if
if sjjh_grade>=8 then%>
<option style="color:blue" value="/监禁$ 请写明原因 ">监禁操作
<option style="color:blue" value="/释放$ 用户名">解除监禁
<%
end if
if sjjh_grade>=9 then%>
<option style="color:red" value="/查看物品$">查看物品
<option style="color:red" value="/分类物品$ 卡片">分类查看
<option style="color:red" value="/放大$ ">放大文字
<option style="color:red" value="/状态$ ">查看状态
<option style="color:red" value="/公告$ ">站长公告
<option style="color:red" value="/斩首$ 原因 ">斩首人犯
<%

%>
<html>
<head>
<title>常用命令操作</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style></style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center"><font color="#FFFFFF"><span style='font-size:9pt'><font size="-1">【</font></span><font size="-1">常用命令<span style='font-size:9pt'>】</span></font></font><font size="3"><br>
</font><br>
  <table cellpadding="3" cellspacing="0" border="1" bgcolor="#CCCCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr > 
      <td> 
        <div align="center"><font color="#330033" size="2">常用命令</font></div>
        </td>
    </tr>
    <tr > 
      <td > 
        <div align="center"><a href="javascript:s('/本门发钱$')" title="掌门、掌门夫人(老公)、长老可以操作！">本门发钱</a></div>
      </td>
    </tr>
  </table>
</div>
<p class=p1 align="center"><font color="#FFFFFF" size="2">杀伤力等于<br>
所用内力+攻击力<br>
<input type="checkbox" name="liji" value="on">
立即攻击 </font></p>
</body>
</html>
