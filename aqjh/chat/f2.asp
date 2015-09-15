<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
'是否可以泡点
if cstr(chatinfo(11))=1 then
	jhpd=0
else
	jhpd=1
end if
if chatbgcolor="" then chatbgcolor="008888"%>
<HTML><HEAD><TITLE>发言区</TITLE>
<STYLE type=text/css>
BODY {FONT-SIZE: 9pt}
INPUT {FONT-SIZE: 9pt}
A {FONT-SIZE: 9pt; COLOR: #ffffff; TEXT-DECORATION: none}
A:hover {COLOR: red; TEXT-DECORATION: underline}
SELECT {BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-BOTTOM-WIDTH: 1px; COLOR: black; FONT-FAMILY: 宋体; BACKGROUND-COLOR: #efefef; BORDER-RIGHT-WIDTH: 1px}
TD {FONT-SIZE: 9pt; COLOR: #ffffff; FONT-FAMILY: 宋体}
</STYLE>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}</script>
<script language="JavaScript"> 
//if(window.top==window.self){var i=1;while (i<=50){window.alert("你想作什么呀，黑我？这里是不行的，去别处玩去吧！哈！慢慢点50次！！");i=i+1;}top.location.href="../exit.asp"}
</script>
</HEAD>
<BODY bgcolor=<%=Application("aqjh_chatbgcolor")%> leftMargin=0 topMargin=0 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
<TABLE height=83 cellSpacing=0 cellPadding=0 width=100% align=right border=0 bordercolor=yellow style="border-collapse: collapse"><TR><TD>
<DIV align=center>
<form name=af method=POST action='say.asp'  target='d' onsubmit='javascript:if(document.af.ctz.checked==true){if(document.af.sytemp.value.indexOf("/")!=0){document.af.sytemp.value="/粗体字$"+document.af.sytemp.value;}}  return(parent.checksays());'>
<input type=hidden name='fs' value='10'>
<input type=hidden name='lh' value='125'>
<input type=hidden name='sy' value=''>
<input type=hidden name='oldsays' value>
<input type=hidden name='oldact' value>
<input type=hidden name='oldtowho' value>
<input type=hidden name='npc' value='<%=Application("aqjh_npc")%>'>
<input type=hidden name='username' readonly size=5 maxlength=10>

<!功能一行>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;background-color:#F7F7F7'>
<option selected>离席
<option style=background-color:"#FFEEFF" value="/离席$ 正常">正常
<option style=background-color:"#FFEEFF" value="/离席$ 睡觉">睡觉
<option style=background-color:"#FFEEEE" value="/离席$ 吃饭">吃饭
<option style=background-color:"#EEFFEE" value="/离席$ 看书">看书
<option style=background-color:"#FFFFEE" value="/离席$ 挂网">挂网
<option style=background-color:"#EEEEFF" value="/离席$ 想事">想事
<option style=background-color:"#EEFFFF" value="/离席$ 逛街">逛街
<option style=background-color:"#EEFFFF" value="/离席$ 约会">约会
<option style=background-color:"#EEFFFF" value="/离席$ 沉思">沉思
<option style=background-color:"#EEFFFF" value="/离席$ 洗澡">洗澡
<option style=background-color:"#EEFFFF" value="/离席$ 泡妞">泡妞
<option style=background-color:"#EEFFFF" value="/离席$ 打架">打架
</select>  
<font color=red>我</font> <select name='addsays' onchange="document.af.sytemp.focus();" style='font-size:12px;background-color:#F7F7F7'> 
<option value="对" selected style="color:red">这样</option> 
<option value="微微笑对">微笑</option> 
<option value="温柔地对">温柔</option> 
<option value="红着脸对">脸红</option> 
<option value="摇头晃脑得意地对">得意</option> 
<option value="哈！哈！哈！笑着对">大笑</option> 
<option value="神秘兮兮地对">神秘</option> 
<option value="战战兢兢地对">战兢</option> 
<option value="毛手毛脚地对">毛手</option> 
<option value="嘟着嘴地对">嘟嘴</option> 
<option value="慢条斯理地对">慢条</option> 
<option value="同情地对">同情</option> 
<option value="幸灾乐祸地对">乐祸</option> 
<option value="快要哭地对">快哭</option> 
<option value="哭着对">哭</option> 
<option value="拳打脚踢地对">拳打</option> 
<option value="不怀好意地对">坏意</option> 
<option value="遗憾地对">遗憾</option> 
<option value="瞪大了眼睛，很诧异地对">诧异</option> 
<option value="幸福地对">幸福</option> 
<option value="翻箱倒柜地对">翻箱</option> 
<option value="悲痛地对">悲痛</option> 
<option value="正义凛然地对">正义</option> 
<option value="严肃地对">严肃</option> 
<option value="生气地对">生气</option> 
<option value="大声地对">大声</option> 
<option value="傻乎乎地对">傻</option> 
<option value="很满足地对">满足</option> 
<option value="手足无措地对">无措</option> 
<option value="很无辜地对">无辜</option> 
<option value="喃喃自语地对">自语</option> 
<option value="恶狠狠地瞪着眼对">瞪眼</option> 
<option value="快要吐地对">想吐</option> 
<option value="无精打采地对">无采</option> 
<option value="依依不舍地对">不舍</option> 
<option value="口吐白沫对">白沫</option> 
<option value="有气无力地对">无力</option> 
<option value="无可奈何地对">无奈</option> 
<option value="可怜兮兮地对">可怜</option> 
</select><font color=red>对</font>                                                                       
<input type="text" name="towho" readonly size="8" onClick="dj()" value="大家" style="height: 18; text-align: center; border-style: groove; border-width:1;background-color:#F7F7F7">  
<font color=blue>说</font> <input type=text name='sytemp' size=38 maxlength=250 style="height: 18; border-style: groove; border-width: 1;background-color:#F7F7F7"> <a href="#" onClick="gop()"><img src="ico/qian.gif" alt="上一句" border="0" align="absmiddle"></a> <a href="#" onClick="gon()"><img src="ico/hou.gif" alt="下一句" border="0" align="absmiddle"></a>        
<input type=submit name='subsay' value='发' style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #0088ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='清' onClick="javascript:parent.qp()" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: dd11ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='魔幻'  onClick="window.open('qq/SelQQ.asp','f3')"style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #8800ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">
<input type=button value='组队'  onClick="window.open('zhu/index.asp','f3')"style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #008888; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='ffffff'"><br>
<!功能二行>
<select name='addwordcolor' style='font-size:12px;background-color:#F7F7F7'> 
<option style="color:0088FF" value="0088FF" selected>名</option> 
<option style="color:000000" value="000000">名</option> 
<option style="color:0000FF" value="0000FF">名</option> 
<option style="color:000088" value="000088">名</option> 
<option style="color:888800" value="888800">名</option> 
<option style="color:008888" value="008888">名</option> 
<option style="color:008800" value="008800">名</option> 
<option style="color:8888FF" value="8888FF">名</option> 
<option style="color:AA00CC" value="AA00CC">名</option> 
<option style="color:8800FF" value="8800FF">名</option> 
<option style="color:888888" value="888888">名</option> 
<option style="color:CCAA00" value="CCAA00">名</option> 
<option style="color:FF8800" value="FF8800">名</option> 
<option style="color:CC3366" value="CC3366">名</option> 
<option style="color:FF00FF" value="FF00FF">名</option> 
<option style="color:3366CC" value="3366CC">名</option>
<option style="color:red" value="red">名</option> 
</select><select name='saycolor'  style='font-size:12px;background-color:#F7F7F7'>  
<option style="color:008800" value="008800" selected>话</option> 
<option style="color:000000" value="000000">话</option> 
<option style="color:0088FF" value="0088FF">话</option> 
<option style="color:0000FF" value="0000FF">话</option> 
<option style="color:000088" value="000088">话</option> 
<option style="color:888800" value="888800">话</option> 
<option style="color:008888" value="008888">话</option> 
<option style="color:888888" value="888888">话</option> 
<option style="color:8888FF" value="8888FF">话</option> 
<option style="color:AA00CC" value="AA00CC">话</option> 
<option style="color:8800FF" value="8800FF">话</option> 
<option style="color:CCAA00" value="CCAA00">话</option> 
<option style="color:FF8800" value="FF8800">话</option> 
<option style="color:CC3366" value="CC3366">话</option> 
<option style="color:FF00FF" value="FF00FF">话</option> 
<option style="color:3366CC" value="3366CC">话</option>
<option style="color:red" value="red">话</option>
</select><%aqjh_name=aqjh_name%><select name='command' onchange="rc(this.value);document.af.command.options[0].selected=true;" style='font-size:12px;'><%                                                                    
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("aqjh_usermdb") 
rs.open "select 会员等级,门派,身份,职业 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2 
hydj=rs("会员等级") 
jhmp=rs("门派") 
jhzy=rs("职业") 
jhsf=rs("身份") 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
%> 
<option value="" selected style="background-color:#F7F7F7">功能一区
<option style="color:#FF6633" value="">==个人==
<option style="color:blue" value="/欢迎新人$">欢迎新人
<option style="color:#FF6633" value="/暂离$ 我出去一下，请等我回来…… ">暂时离开
<option style="color:#FF6633" value="/回来$ ">上线回来
<option style="color:blue" value="/新人费$ 100000">送新人费
<option style="color:red" value="/新人福神$">新人福神
<option style="color:#0808CF" value="/申领金币$">申领金币
<option style="color:#B000D0" value="/邀请赏花$ ">邀请赏花
<option style="color:#B000D0" value="/邀请赏猪$ ">邀请看猪
<%if chatinfo(5)=0 then%>
<option style="color:#FF6633" value="/邀请串门$">邀请串门
<option style="color:#FF6633" value="/观赏家园$ ">观赏家园
<option style="color:blue" value="/邀请加入组$ ">邀请加组
<option style="color:#FF6633" value="/暴豆$ ">暴威力豆 <%end if%>
<option style="color:#FF6633" value="/好友$ 加入">结义兄弟
<option style="color:#FF6633" value="/好友$ 查看">查看兄弟
<option style="color:#FF6633" value="/好友$ 删除">割袍段义
<option style="color:#FF6633" value="/清除$ 要清除的好友姓名">清除兄弟
<option style="color:#FF6633" value="/标题$ 拉人官府有重赏!">更改标题 <% if Application("aqjh_baowu")>0 then%>
<option style="color:#FF6633" value="/修练$ ">宝物修练 <%end if%>
<option style="color:blue" value="/屏蔽$ 加入">加入屏蔽
<option style="color:blue" value="/屏蔽$ 查看">查看屏蔽
<option style="color:blue" value="/屏蔽$ 删除">去掉屏蔽
<option style="color:blue" value="/信息$">查看信息
<option style="color:red" value="">==银行==
<option style="color:#B000D0" value="/存钱$ 1000">银行存钱
<option style="color:#B000D0" value="/取钱$ 1000">银行取钱
<option style="color:#B000D0" value="/转账$ 10000">银行转账
<option style="color:#B000D0" value="">==金库==
<option style="color:#B000D0" value="/金币存库$ 10">金币存库
<option style="color:#B000D0" value="/金库取币$ 10">金库取币
<option style="color:red" value="">==门派==
<option style="color:#0808CF" value="/加入门派$">加入门派
<option style="color:#0808CF" value="/本门发钱$">本门发钱
<option style="color:#0808CF" value="/指导弟子$">指导弟子
<option style="color:#0808CF" value="/基金$">本门基金
<option style="color:#0808CF" value="/收徒$ ">招收徒弟
<option style="color:#0808CF" value="/逐出$ 徒弟名">逐出师门
<option style="color:#0808CF" value="/拜师$ ">拜师习武
<option style="color:blue" value="/奖励徒弟$ 100">奖励徒弟
<option style="color:blue" value="/惩罚徒弟$ 100">惩罚徒弟
<option style="color:#0808CF" value="/学成下山$ 师傅名">学艺有成
<option style="color:#0808CF" value="/挑战掌门$ 请输入掌门人的名字">挑战掌门
<option style="color:red" value="">==婚姻==
<option style="color:#FF6699" value="/求婚$ 真爱表白">示爱求婚
<option style="color:#FF6699" value="/发喜糖$ 500">分发喜糖
<option style="color:#FF6699" value="/飞吻$">在线飞吻
<option style="color:#FF6699" value="/亲密接触$">亲密接触
<option style="color:#FF6699" value="/抢亲$">抢强民女
<option style="color:red" value="">==动作==
<option style="color:#0808CF" value="/主题$ ">千里传音
<option style="color:#0808CF" value="/心动$ ">心动感觉
<option style="color:#0808CF" value="/怒吼$ ">狂狮怒吼
<option style="color:#0808CF" value="/心跳$ ">心跳心语
<option style="color:red" value="">==偷袭==
<option style="color:#333333" value="/单挑$ 看你拽还是我拽">不服单挑
<%if chatinfo(5)=0 then%>
<option style="color:#333333" value="/吸星大法$ ">吸取内力
<option style="color:#333333" value="/偷钱$ ">偷取钱财
<option style="color:#333333" value="/偷金币$ ">偷取金币
<option style="color:#333333" value="/乾坤一掷$ 10000">乾坤一掷 
<option style="color:#333333" value="/下毒$ 毒药名">偷偷下毒
<option style="color:#333333" value="/投掷$ 暗器名">投掷暗器
<option style="color:#333333" value="/发招$ 招式名">发招攻击
<option style="color:#333333" value="/宠物搏斗$ 宠物名称">宠物搏斗
<option style="color:#333333" value="/拿钱砸人$ 10000">拿钱砸人 <%end if%>
<%if jhsf="掌门" or jhsf="元老" then%>
<option style="color:#993333" value="">==掌门==
<option value="/逐出弟子$">逐出弟子
<option value="/册封$ 身份名">册封身份
<option style="color:blue" value="/册封$ 元老">册封元老
<option value="/册封$ 长老">册封长老
<option value="/册封$ 护法">册封护法
<option value="/册封$ 堂主">册封堂主
<option value="/册封$ 弟子">取消册封
<%end if
if (jhsf="长老" or jhsf="掌门" or jhsf="元老") and aqjh_grade>=4 then%>
<option value="/招收弟子$">招收弟子
<option value="/哑穴$ 请写明原因 ">教训弟子
<option value="/奖励$ 100 ">奖励弟子
<option value="/掌门令$ 命令 ">掌门下令
<option value="/门派ip$">门 派ip
<option style="color:red" value="/追杀令$ 命令 ">追 杀 令
<option value="/本派罚款$ 1000 ">本派罚款
<option value="/本派刑法$ 1000 ">本派刑法
<option value="/掌门发放$">掌门发放 <%end if%>
</select><select name='addsign' onChange="rc(this.value);" style='font-size:12px;background-color:#F7F7F7'>  
        <option value="" selected>功能二区</option>
<option style="color:red" value="">=新功能=
        <option style="color:blue" value="/吃饭$">请你进餐
        <option style="color:blue" value="/生日$">祝你生日
        <option style="color:blue" value="/邀请跳舞$ ">邀请跳舞
        <option style="color:blue" value="/杀气锐减$ 1">杀气锐减
        <option style="color:blue" value="/神灵降世$ 100000">神灵降世
        <option style="color:blue" value="/照顾新人$ 5000">照顾新人
        <option style="color:red" value="/走运气$">江湖运气
        <option style="color:red" value="/隐身$">转生隐身  
        <option style="color:red" value="/邀请语聊$">邀请语聊 
        <option style="color:red" value="/语音聊天$">语音聊天
        <option style="color:red" value="/成为盟主$">成为盟主
        <option style="color:red" value="/盟主令$ 帮助">盟 主 令
<option style="color:red" value="">==其他==
        <option style="color:blue" value="/申领奖励$">申领奖励</option>
        <option style="color:blue" value="/新人奖励$ 1">新人奖励</option>
        <option style="color:blue" value="/高级奖励$ 1">高级奖励</option>
<option style="color:red" value="">==国家==
<option style="color:blue" value="/国库资金$">查看国库</option>
 <option style="color:blue" value="/国家册封$ 丞相/将军/侍卫/自定义">君主册封</option>
 <option style="color:blue" value="/国家令$  写上你要说的话">国家公告</option>
 <option style="color:blue" value="/国家赈灾$ 10000 ">国家赈灾</option>
<option style="color:blue" value="/国家招收$ ">国家招收</option>
<option style="color:blue" value="/挑战君王$ ">挑战君王</option>
<option style="color:red" value="">==国战==
<option style="color:blue" value="/国家大战$ ">国家大战</option>
<option style="color:blue" value="/国家挑战$  写上对方君王的名字,挑战失败,国家灭亡,注意!">国家挑衅</option>
<option style="color:blue" value="/国家求和$  写上对方君王的名字">国家求和</option>
<option style="color:red" value="">==杂项==
<option style="color:blue" value="/振臂一呼$ ">振臂一呼
<option style="color:blue" value="/嫁衣神功$ 10000">嫁衣神功
<option style="color:blue" value="/暗然销魂$ 10000">暗然销魂
<option style="color:blue" value="/自杀$">悬梁自禁 
<option style="color:red" value="">==种族==
<option style="color:red" value="/人族任务$">人族任务 
<option style="color:red" value="/神族任务$">神族任务
<option style="color:red" value="/魔族任务$">魔族任务
<option style="color:red" value="/修炼精神$ 人类的修炼">修炼精神 
<option style="color:red" value="/修炼魔力$ 魔族的修炼">修炼魔力
<option style="color:red" value="/修炼仙术$ 神仙的修炼">修炼仙术
<option style="color:red" value="">==存送==
        <option style="color:#B000D0" value="/赠送$ 类别|名字|数量">赠送物品
        <option style="color:#B000D0" value="/丢弃$ 100000">丢弃银两
        <option style="color:#B000D0" value="/用卡$ 卡片名">使用卡片
        <option style="color:#B000D0" value="/送钱$ 1000">赠送现金
        <option style="color:#B000D0" value="/送豆点$ 10">赠送豆点
        <option style="color:#B000D0" value="/送金币$ 1">赠送金币
        <option style="color:blue" value="/传授内力$ 100">传授内力
        <option style="color:blue" value="/传授体力$ 100">传授体力
        <option style="color:blue" value="/传授武功$ 100">传授武功
        <option style="color:blue" value="/传授经验$ 100">传授经验
        <option style="color:blue" value="/传授魅力$ 100">传授魅力
        <option style="color:blue" value="/传授道德$ 100">传授道德
        <option style="color:blue" value="/传送法力$ 100">传送法力</option>
        <option style="color:blue" value="/传授轻功$ 100">传授轻功</option>
<option style="color:red" value="">==修炼==
<option style="color:#339933" value="/打坐$ ">打坐练功
<option style="color:#339933" value="/闭目$ ">闭目养神
<option style="color:#339933" value="/练武$ ">武馆习武
<option style="color:#339933" value="/顿悟$ ">江湖顿悟
<option style="color:#339933" value="/提点智力$ ">提高智力
<option style="color:#339933" value="/修炼法力$ ">修炼法力
<option style="color:#339933" value="/修炼轻功$ ">修炼轻功
<option style="color:#996633" value="/修习佛法$">修习佛法
<option style="color:#339933" value="/练金属性$ ">练金属性
<option style="color:#339933" value="/练木属性$ ">练木属性
<option style="color:#339933" value="/练水属性$ ">练水属性
<option style="color:#339933" value="/练火属性$ ">练火属性
<option style="color:#339933" value="/练土属性$ ">练土属性
<option style="color:#339933" value="/休身养性$ ">休身养性
 <option style="color:red" value="">==魔法==</option>
        <option style="color:#B000D0" value="/新阴那山$">新阴那山</option>
        <option style="color:#B000D0" value="/新雁南飞$">新雁南飞</option>
        <option style="color:#B000D0" value="/点石成金$">点石成金</option>
        <option style="color:#B000D0" value="/养心大法$">养心大法</option>
        <option style="color:#B000D0" value="/招财进宝$">招财进宝</option>
        <option style="color:#B000D0" value="/妙手回春$">妙手回春</option>
 <option style="color:red" value="">==混合==</option> 
        <option style="color:green" value="/蓝玛瑙$ 100">蓝 玛 瑙</option>       
        <option style="color:green" value="/绿玛瑙$ 100">绿 玛 瑙</option>
        <option style="color:green" value="/红玛瑙$ 100">红 玛 瑙</option>
        <option style="color:green" value="/魔法石$ 1">魔 法 石</option>
        <option style="color:green" value="/孔雀翎$ 50">孔 雀 翎</option>
        <option style="color:green" value="/帅哥令$ ">帅 哥 令</option>
        <option style="color:green" value="/美人令$ ">美 人 令</option>
        <option style="color:green" value="/多情环$ ">多 情 环</option>
        <option style="color:green" value="/绝命钩$ ">绝 命 钩</option>
 <option style="color:red" value="">==攻击==</option>
        <option style="color:blue" value="/魔幻水晶$ ">魔幻水晶</option>
        <option style="color:blue" value="/发射子弹$ ">发射子弹</option>
        <option style="color:blue" value="/使出$ 九阳神功 ">九阳神功</option>
        <option style="color:blue" value="/使出$ 天堂令 ">天 堂 令 </option>
        <option style="color:blue" value="/使出$ 玄冥棒 ">玄 冥 棒 </option>
        <option style="color:blue" value="/使出$ 盗取令 ">盗 取 令 </option>
        <option style="color:blue" value="/使出$ 七伤拳 ">七 伤 拳 </option>
        <option style="color:blue" value="/使出$ 圣火令 ">圣 火 令 </option>
        <option style="color:blue" value="/绝情刀$ ">绝 情 刀 </option>
        <option style="color:blue" value="/抢劫令$ ">抢 劫 令 </option>
        <option style="color:blue" value="/使出$ 铁笔银勾 ">铁笔银勾 </option>
 <option style="color:red" value="">==字型==</option>
        <option style="color:green" value="/大吼吼$ ">大 吼 吼 
        <option style="color:green" value="/小吼吼$ ">小 吼 吼
        <option style="color:#FF6699" value="/粗体字$ ">粗 体 字</option>
        <option style="color:#FF6699" value="/飞舞字$ ">飞 舞 字</option>
        <option style="color:#FF6699" value="/按钮字$ ">按 钮 字</option>
        <option style="color:#FF6699" value="/滚动按钮$ ">滚动按钮</option>     
        <option style="color:#FF6699" value="/上下按钮$ ">上下按钮</option>
         <option style="color:#FF6699" value="/爱神$ ">爱神传音</option>
         <option style="color:#FF6699" value="/颠倒$ ">颠倒字体</option>
        <option style="color:#FF6699" value="/字体魔法$ 内容|颜色|大小|字体">字体魔法</option>
        <option style="color:#FF6699" value="/移动魔法$ 内容|颜色|背景|方式|速度|延时">滚动魔法</option>
        <option style="color:#FF6699" value="/按钮魔法$ 内容|颜色|弹出内容">按钮魔法</option>              
 <option style="color:red" value="">==NPC==</option>
        <option style="color:#FF6699" value="/召唤$ ">召 唤 术</option>
        <option style="color:#FF6699" value="/加血$ ">治 愈 术</option>
</select><select name='com1' onchange="rc(this.value);document.af.com1.options[0].selected=true;" style='font-size:12px;'>
<option value="" selected>职业技能</option>
 <option style="color:red" value="">==出家==</option> 
<option style="color:blue" value="/出家$ 出家后只能吃斋，不能吃肉和近女色哦">出家修行
<option style="color:blue" value="/还俗$">修成还俗
<option style="color:blue" value="/念经$">普渡众生
<option style="color:blue" value="/如沐春风$">如沐春风
 <option style="color:red" value="">==天网==</option>
<option style="color:blue" value="/加入天网$ 加入天网后不能保护">加入天网
<option style="color:blue" value="/离开天网$">脱离天网
 <option style="color:red" value="">==其他==</option>
<option style="color:#996633" value="/拐骗少女$ ">拐骗少女
<option style="color:#996633" value="/拐骗少男$ ">拐骗少男
<option style="color:#996633" value="/抓小偷$ ">抓 小 偷
<option style="color:#996633" value="/生育$ 小孩名称">生育小孩 <%if jhzy="大夫" then%>
<option style="color:#996633" value="/治病$ ">妙手回春 <%end if%> <%if jhzy="武师" then%>
<option style="color:#996633" value="/教武$ ">教导武功 <%end if%> <%if jhzy="算命师" then%>
<option style="color:#996633" value="/求签$ ">求神问卜 <%end if%> <%if jhzy="倒夜香" then%>
<option style="color:#996633" value="/倒夜香$ ">倒夜香师 <%end if%>
<option style="color:#996633" value="/扫黄行动$">扫黄行动
<option style="color:#996633" value="/同归于尽$">同归于尽
<option style="color:red" value="">=召唤师=</option>
<option style="color:#333333" value="/神兽攻击$">神兽攻击
<option style="color:#333333" value="/唤回神兽$">唤回神兽
 <option style="color:red" value="">=冒险家=</option>        
        <option style="color:green" value="/寻找鲜花$ ">寻找鲜花</option>
        <option style="color:green" value="/寻找金卡$ ">寻找金卡</option>
        <option style="color:green" value="/寻找金币$ ">寻找金币</option>
        <option style="color:green" value="/寻找卡片$ ">寻找卡片</option>
        <option style="color:green" value="/寻找秘笈$ ">寻找秘笈</option>
        <option style="color:green" value="/寻水晶球$ ">寻找水晶</option>
        <option style="color:green" value="/寻找法器$ ">寻找法器</option>
        <option style="color:green" value="/寻找魔器$ ">寻找魔器</option>
        <option style="color:green" value="/宝物术$ ">寻找宝物</option>
 <option style="color:red" value="">=魔法师=</option>
        <option style="color:blue" value="/乞讨神术$ ">乞讨神术</option>
        <option style="color:blue" value="/布施术$ ">布施天下</option>
        <option style="color:blue" value="/魅惑人间$ ">魅惑人间</option>
        <option style="color:blue" value="/魔界咒语$ ">魔界咒语</option>
        <option style="color:blue" value="/迷魂大法$ ">迷魂大法</option>
        <option style="color:blue" value="/配制令牌$ ">配制令牌</option>
        <option style="color:blue" value="/配制宝石$ ">配制宝石</option>
 <option style="color:red" value="">=炼金师=</option>
        <option style="color:blue" value="/暗器合成$ ">暗器合成</option>
        <option style="color:blue" value="/卡片合成$ ">卡片合成</option>
        <option style="color:blue" value="/百变神通$ ">百变神通</option>
        <option style="color:blue" value="/炼金术$ 100">炼 金 术</option>
        <option style="color:blue" value="/炼魔法石$">炼魔法石</option>
        <option style="color:blue" value="/解毒术$ 100">解 毒 术</option>
        <option style="color:blue" value="/平安术$ 100">平 安 术</option>
        <option style="color:blue" value="/偷窃术$ 100">偷 窃 术</option>
        <option style="color:blue" value="/摇钱术$ 100">摇 钱 术</option>
        <option style="color:blue" value="/魔力钻石$ ">魔力钻石</option>
        <option style="color:blue" value="/生日蛋糕$ ">生日蛋糕</option>
 <option style="color:red" value="">==官府==</option>
        <option style="color:blue" value="/没收法器$ ">没收法器</option>
        <option style="color:blue" value="/没收魔器$ ">没收魔器</option>
 <option style="color:red" value="">==各人==</option>
        <option style="color:#B000D0" value="/存金属性$ 10">存金属性
        <option style="color:#B000D0" value="/取金属性$ 10">取金属性
        <option style="color:#B000D0" value="/存木属性$ 10">存木属性
        <option style="color:#B000D0" value="/取木属性$ 10">取木属性
        <option style="color:#B000D0" value="/存水属性$ 10">存水属性
        <option style="color:#B000D0" value="/取水属性$ 10">取水属性
        <option style="color:#B000D0" value="/存火属性$ 10">存火属性
        <option style="color:#B000D0" value="/取火属性$ 10">取火属性
        <option style="color:#B000D0" value="/存土属性$ 10">存土属性
        <option style="color:#B000D0" value="/取土属性$ 10">取土属性
        <option style="color:#FF6699" value="/存法力$ 100">存放法力</option>
        <option style="color:#FF6699" value="/取法力$ 100">取回法力</option>
        <option style="color:#FF6699" value="/轻功暂存$ 100">轻功暂存</option>
        <option style="color:#FF6699" value="/提出轻功$ 100">提出轻功</option>
 <option style="color:red" value="">==抢劫==</option>
        <option style="color:FF8800" value="/抢劫银行$">抢劫银行</option>
        <option style="color:FF8800" value="/抢劫金库$">抢劫金库</option>
 <option style="color:red" value="">==家法==</option>
        <option style="color:#B000D0" value="/拧嘴$">拧 嘴 嘴</option>
        <option style="color:#B000D0" value="/鸡毛掸子$">鸡毛掸子</option>
        <option style="color:#B000D0" value="/跪搓板$">跪 搓 板</option>
</select><select name='dubo' onchange="rc(this.value);document.af.dubo.options[0].selected=true;" style='font-size:12px;background-color:#F7F7F7'>                                                       
<option value="" selected>=赌博=
        <option value="/打扑克$ 1000|银两[或：金币，体力，内力]" style=color:red>打扑克 
        <option value="/发牌$ 公证|负[或胜]" style=color:red>-公证-
        <option value="/打麻将$ 1000|银两[或：金币，体力，内力]" style=color:red>打麻将 
        <option value="/出牌$ 公证|负[或胜]" style=color:red>-公证-
<option style="color:blue" value="/江湖赛马">赛&nbsp;&nbsp;马
<option style="color:blue" value="/双人赌博$ 1">赌石头
<option style="color:blue" value="/双人赌博$ 2">赌剪子
<option style="color:blue" value="/双人赌博$ 3">赌博布
<option style="color:red" value="/赌法力$100">赌法力
<option style="color:red" value="/赌轻功$100">赌轻功
<option style="color:red" value="/赌金属性$100">赌金属
<option style="color:red" value="/赌木属性$100">赌木属
<option style="color:red" value="/赌水属性$100">赌水属
<option style="color:red" value="/赌火属性$100">赌火属
<option style="color:red" value="/赌土属性$100">赌土属
<option style="color:red" value="/赌豆$100">赌豆点
<option style="color:red" value="/对赌$2">赌金币
<option style="color:red" value="/幸运$ 幸运数字1-1000">风&nbsp;&nbsp;采
<option style="color:red" value="/赌博$ 消息">赌消息
<option style="color:blue" value="/赌博$ 银作庄">银作庄
<option style="color:blue" value="/下注$ 银&大&10000">银押大
<option style="color:blue" value="/下注$ 银&小&10000">银押小
<option style="color:blue" value="/下注$ 银&双&10000">银押双
<option style="color:blue" value="/下注$ 银&单&10000">银押单
<option style="color:green" value="/赌博$ 点作庄">点作庄
<option style="color:green" value="/下注$ 点&大&100">点押大
<option style="color:green" value="/下注$ 点&小&100">点押小
<option style="color:green" value="/下注$ 点&双&100">点押双
<option style="color:green" value="/下注$ 点&单&100">点押单
<option style="color:blue" value="/赌博$ 法作庄">法作庄
<option style="color:blue" value="/下注$ 法&大&100">法押大
<option style="color:blue" value="/下注$ 法&小&100">法押小
<option style="color:blue" value="/下注$ 法&双&100">法押双
<option style="color:blue" value="/下注$ 法&单&100">法押单
<option style="color:green" value="/赌博$ 轻作庄">轻作庄
<option style="color:green" value="/下注$ 轻&大&100">轻押大
<option style="color:green" value="/下注$ 轻&小&100">轻押小
<option style="color:green" value="/下注$ 轻&双&100">轻押双
<option style="color:green" value="/下注$ 轻&单&100">轻押单
</select><select name='tu' onChange="rc(this.value);document.af.tu.options[0].selected=true;" style='font-size:12px;background-color:#F7F7F7'>
<option value="" selected>图片区</option>
<option style="color:red" value="">A版图</option> 
<option value="[IMG]a1.gif[/IMG]">困 了</option>
<option value="[IMG]a2.gif[/IMG]">惊 讶</option>
<option value="[IMG]a3.gif[/IMG]">无 聊</option>
<option value="[IMG]a4.gif[/IMG]">天 真</option>
<option value="[IMG]a5.gif[/IMG]">发 狂</option>
<option value="[IMG]a6.gif[/IMG]">调 皮</option>
<option value="[IMG]a7.gif[/IMG]">生 气</option>
<option value="[IMG]a27.gif[/IMG]">好 色</option>
<option value="[IMG]a9.gif[/IMG]">呆 了</option>
<option value="[IMG]a10.gif[/IMG]">无 奈</option>
<option value="[IMG]a11.gif[/IMG]">高 傲</option>
<option value="[IMG]a12.gif[/IMG]">尴 尬</option>
<option value="[IMG]a13.gif[/IMG]">害 羞</option>
<option value="[IMG]a14.gif[/IMG]">头 晕</option>
<option value="[IMG]a15.gif[/IMG]">痛 哭</option>
<option value="[IMG]a16.gif[/IMG]">得 意</option>
<option value="[IMG]a17.gif[/IMG]">闭 嘴</option>
<option value="[IMG]a18.gif[/IMG]">装 酷</option>
<option value="[IMG]a19.gif[/IMG]">不 妙</option>
<option value="[IMG]a20.gif[/IMG]">装 傻</option>
<option value="[IMG]a21.gif[/IMG]">发 火</option>
<option value="[IMG]a22.gif[/IMG]">可 爱</option>
<option value="[IMG]a23.gif[/IMG]">偷 笑</option>
<option value="[IMG]a24.gif[/IMG]">疑 问</option>
<option value="[IMG]a25.gif[/IMG]">笑 容</option>
<option style="color:red" value="">B版图</option>
<option value="[IMG]b1.gif[/IMG]">傻 笑</option>
<option value="[IMG]b2.gif[/IMG]">装 酷</option>
<option value="[IMG]b3.gif[/IMG]">不 妙</option>
<option value="[IMG]b4.gif[/IMG]">嘿 嘿</option>
<option value="[IMG]b5.gif[/IMG]">口 水</option>
<option value="[IMG]b6.gif[/IMG]">诚 恳</option>
<option value="[IMG]b7.gif[/IMG]">冷 汗</option>
<option style="color:red" value="">C版图</option>
<option value="[IMG]c1.gif[/IMG]">美 貌</option>
<option value="[IMG]c2.gif[/IMG]">嘻 嘻</option>
<option value="[IMG]c3.gif[/IMG]">塞 车</option>
<option value="[IMG]c4.gif[/IMG]">接 吻</option>
<option value="[IMG]c5.gif[/IMG]">加 班</option>
<option value="[IMG]c6.gif[/IMG]">单 条</option>
<option value="[IMG]c7.gif[/IMG]">蜡 烛</option>
<option value="[IMG]c8.gif[/IMG]">奋 斗</option>
<option value="[IMG]c9.gif[/IMG]">打 架</option>
<option value="[IMG]c10.gif[/IMG]">睡 吧</option>
<option value="[IMG]c11.gif[/IMG]">我 晕</option>
<option value="[IMG]c12.gif[/IMG]">爱 你</option>
<option value="[IMG]c13.gif[/IMG]">别惹我</option>
<option value="[IMG]c14.gif[/IMG]">拜 托</option>
<option value="[IMG]c15.gif[/IMG]">傻 笑</option>
<option value="[IMG]c16.gif[/IMG]">被离开</option>
<option value="[IMG]c17.gif[/IMG]">鄙视你</option>
<option value="[IMG]c18.gif[/IMG]">熊猫眼</option>
<option value="[IMG]c19.gif[/IMG]">去死吧</option>
<option value="[IMG]c20.gif[/IMG]">同 仇</option>
<option value="[IMG]c21.gif[/IMG]">强 迫</option>
<option value="[IMG]c22.gif[/IMG]">不理你</option>
<option style="color:red" value="">D版图</option>
<option value="[IMG]d1.gif[/IMG]">狂 哭</option>
<option value="[IMG]d2.gif[/IMG]">亲一个</option>
<option value="[IMG]d3.gif[/IMG]">讨 厌</option>
<option value="[IMG]d4.gif[/IMG]">白 牙</option>
<option value="[IMG]d5.gif[/IMG]">眼 泪</option>
<option value="[IMG]d6.gif[/IMG]">恳 求</option>
<option value="[IMG]d7.gif[/IMG]">舒 服</option>
<option value="[IMG]d8.gif[/IMG]">烧 香</option>
<option value="[IMG]d9.gif[/IMG]">掉 汗</option>
<option value="[IMG]d10.gif[/IMG]">看着你</option>
<option value="[IMG]d11.gif[/IMG]">自 信</option>
<option value="[IMG]d12.gif[/IMG]">老 大</option>
<option style="color:red" value="">E版图</option>
<option value="[IMG]e1.gif[/IMG]">高 兴</option>
<option value="[IMG]e2.gif[/IMG]">拜 拜</option>
<option value="[IMG]e8.gif[/IMG]">撞 头</option>
<option value="[IMG]e22.gif[/IMG]">快 逃</option>
<option value="[IMG]e23.gif[/IMG]">唱 歌</option>
<option value="[IMG]e24.gif[/IMG]">相 亲</option>
<option value="[IMG]e25.gif[/IMG]">浪 漫</option>
<option value="[IMG]e26.gif[/IMG]">拳 击</option>
<option value="[IMG]e27.gif[/IMG]">敲 晕</option>
<option value="[IMG]e28.gif[/IMG]">打篮球</option>
<option value="[IMG]e29.gif[/IMG]">沉 默</option>
<option value="[IMG]e30.gif[/IMG]">美女啊</option>
<option value="[IMG]e31.gif[/IMG]">逗 你</option>
<option value="[IMG]e32.gif[/IMG]">狂 哭</option>
<option value="[IMG]e33.gif[/IMG]">高 兴</option>
<option value="[IMG]e34.gif[/IMG]">干 杯</option>
<option value="[IMG]e35.gif[/IMG]">发 怒</option>
<option value="[IMG]e36.gif[/IMG]">不 满</option>
<option value="[IMG]e37.gif[/IMG]">牛 仔</option>
<option value="[IMG]e38.gif[/IMG]">新 婚</option>
<option value="[IMG]e39.gif[/IMG]">流 氓</option>
<option value="[IMG]e40.gif[/IMG]">困死了</option>
<option value="[IMG]e41.gif[/IMG]">吃西瓜</option>
<option value="[IMG]e42.gif[/IMG]">送情书</option>
<option value="[IMG]e43.gif[/IMG]">你 好</option>
<option value="[IMG]e44.gif[/IMG]">害 羞</option>
<option value="[IMG]e45.gif[/IMG]">罚出场</option>
<option value="[IMG]e46.gif[/IMG]">大 侠</option>
<option value="[IMG]e3.gif[/IMG]">无影掌</option>
<option value="[IMG]e4.gif[/IMG]">太极拳</option>
<option value="[IMG]e5.gif[/IMG]">灭 火</option>
<option value="[IMG]e6.gif[/IMG]">走 了</option>
<option value="[IMG]e7.gif[/IMG]">猪流汗</option>
<option value="[IMG]e9.gif[/IMG]">拍屁股</option>
<option value="[IMG]e10.gif[/IMG]">我是鱼</option>
<option value="[IMG]e11.gif[/IMG]">流氓兔</option>
<option value="[IMG]e12.gif[/IMG]">阴 暗</option>
<option value="[IMG]e13.gif[/IMG]">吻 你</option>
<option value="[IMG]e14.gif[/IMG]">被 打</option>
<option value="[IMG]e15.gif[/IMG]">擦 汗</option>
<option value="[IMG]e16.jpg[/IMG]">滴 汗</option>
<option value="[IMG]e17.gif[/IMG]">犹 豫</option>
<option value="[IMG]e18.jpg[/IMG]">无 奈</option>
<option value="[IMG]e19.jpg[/IMG]">这 个</option>
<option value="[IMG]e20.jpg[/IMG]">嘴 搀</option>
<option value="[IMG]e21.jpg[/IMG]">傻 笑</option>
<option style="color:red" value="">F版图</option>
<option value="[IMG]f1.jpg[/IMG]">没问题</option> 
<option value="[IMG]f2.gif[/IMG]">射杀你</option> 
<option value="[IMG]f3.gif[/IMG]">闹别扭</option> 
<option value="[IMG]f4.gif[/IMG]">眼 泪</option> 
<option value="[IMG]f5.gif[/IMG]">睡甜甜</option> 
<option value="[IMG]f6.gif[/IMG]">活 泼</option> 
<option value="[IMG]f7.gif[/IMG]">混 迹</option> 
<option value="[IMG]f8.gif[/IMG]">笑肚子</option> 
<option value="[IMG]f9.gif[/IMG]">我扁你</option> 
<option value="[IMG]f10.gif[/IMG]">天鹅舞</option> 
<option value="[IMG]f11.gif[/IMG]">发 怒</option> 
<option value="[IMG]f12.gif[/IMG]">左右摆</option> 
<option value="[IMG]f13.gif[/IMG]">冒冷汗</option> 
<option value="[IMG]f14.gif[/IMG]">亲 你</option> 
<option value="[IMG]f15.gif[/IMG]">猪哭了</option> 
<option value="[IMG]f16.gif[/IMG]">猪捎痒</option>
<option value="[IMG]f17.gif[/IMG]">鄙 视</option>
<option value="[IMG]f18.gif[/IMG]">吐舌头</option>
<option value="[IMG]f19.JPG[/IMG]">笑 死</option>
<option value="[IMG]f20.gif[/IMG]">受 伤</option>
<option value="[IMG]f21.gif[/IMG]">色 色</option>
<option value="[IMG]f22.gif[/IMG]">变 脸</option>
<option value="[IMG]f23.gif[/IMG]">胜 利</option>
<option value="[IMG]f24.gif[/IMG]">舌 头</option>
<option value="[IMG]f25.gif[/IMG]">没 种</option>
<option value="[IMG]f26.gif[/IMG]">逗 你</option>
<option value="[IMG]f27.gif[/IMG]">童 真</option>
<option style="color:red" value="">G版图</option>
<option value="[IMG]g1.gif[/IMG]">打死你</option>
<option value="[IMG]g2.gif[/IMG]">我打你</option>
<option value="[IMG]g3.gif[/IMG]">群 殴</option>
<option value="[IMG]g4.gif[/IMG]">大胜利</option>
<option value="[IMG]g5.gif[/IMG]">照 相</option>
<option value="[IMG]g6.gif[/IMG]">猪生气</option>
<option value="[IMG]g7.gif[/IMG]">布娃娃</option>
<option value="[IMG]g8.gif[/IMG]">哈哈哈</option>
<option value="[IMG]g9.gif[/IMG]">受委屈</option>
<option value="[IMG]g10.gif[/IMG]">独自玩</option>
<option value="[IMG]g11.gif[/IMG]">笑你呆</option>
<option value="[IMG]g12.gif[/IMG]">到永远</option>
<option value="[IMG]g13.gif[/IMG]">甜蜜蜜</option>
<option value="[IMG]g14.gif[/IMG]">命 苦</option>
<option value="[IMG]g15.jpg[/IMG]">大坏蛋</option>
<option value="[IMG]g16.gif[/IMG]">在一起</option>
<option value="[IMG]g17.gif[/IMG]">打开窗</option>
<option value="[IMG]g18.gif[/IMG]">一起呆</option>
<option value="[IMG]g19.gif[/IMG]">幸 福</option>
<option value="[IMG]g20.gif[/IMG]">甜死人</option>
<option value="[IMG]g21.gif[/IMG]">我爱你</option>
<option value="[IMG]g22.gif[/IMG]">潇 洒</option>
<option value="[IMG]g23.gif[/IMG]">擦眼泪</option>
<option value="[IMG]g24.gif[/IMG]">在跳舞</option>
<option value="[IMG]g25.gif[/IMG]">我错了</option>
<option value="[IMG]g26.gif[/IMG]">多靓仔</option>
<option value="[IMG]g27.gif[/IMG]">想追你</option>
<option value="[IMG]g28.gif[/IMG]">嫁给我</option>
<option value="[IMG]g29.gif[/IMG]">来打针</option>
<option value="[IMG]g30.gif[/IMG]">凶女人</option>
<option value="[IMG]g31.gif[/IMG]">帅呆了</option>
<option value="[IMG]g32.gif[/IMG]">我服了</option>
<option value="[IMG]g33.gif[/IMG]">请求你</option>
<option value="[IMG]g34.gif[/IMG]">砍死你</option>
<option value="[IMG]g35.gif[/IMG]">敲死你</option>
<option style="color:red" value="">H版图</option>
<option value="[IMG]h1.gif[/IMG]">色迷迷</option>
<option value="[IMG]h2.gif[/IMG]">嘟 嘴</option>
<option value="[IMG]h3.gif[/IMG]">叹 气</option>
<option value="[IMG]h4.gif[/IMG]">你厉害</option>
<option value="[IMG]h5.gif[/IMG]">醉 汉</option>
<option value="[IMG]h6.gif[/IMG]">不是吧</option>
<option value="[IMG]h7.gif[/IMG]">倒 霉</option>
<option value="[IMG]h8.gif[/IMG]">武 装</option>
<option value="[IMG]h9.gif[/IMG]">钱 眼</option>
<option value="[IMG]h10.gif[/IMG]">汗 水</option>
<option value="[IMG]h11.gif[/IMG]">调 皮</option>
<option value="[IMG]h12.gif[/IMG]">鄙 视</option>
<option value="[IMG]h13.gif[/IMG]">恭 喜</option>
<option value="[IMG]h14.gif[/IMG]">发 怒</option>
<option value="[IMG]h15.gif[/IMG]">亲 亲</option>
<option value="[IMG]h16.gif[/IMG]">偷 窃</option>
<option value="[IMG]h17.gif[/IMG]">偷 哭</option>
<option value="[IMG]h18.gif[/IMG]">偷 笑</option>
<option value="[IMG]h19.gif[/IMG]">困 了</option>
<option value="[IMG]h20.gif[/IMG]">骷 髅</option>
<option value="[IMG]h21.gif[/IMG]">流 泪</option>
<option value="[IMG]h22.gif[/IMG]">疑 问</option>
<option value="[IMG]h23.gif[/IMG]">傻 笑</option>
<option value="[IMG]h24.gif[/IMG]">非 典</option>
<option value="[IMG]h25.gif[/IMG]">哈 哈</option>
<option value="[IMG]h26.gif[/IMG]">流 汗</option>
<option value="[IMG]h27.gif[/IMG]">来打架</option>
<option value="[IMG]h28.gif[/IMG]">倒 霉</option>
<option value="[IMG]h29.gif[/IMG]">奚 落</option>
<option value="[IMG]h30.gif[/IMG]">头 晕</option>
<option value="[IMG]h31.gif[/IMG]">不 依</option>
<option value="[IMG]h32.gif[/IMG]">天 真</option>
<option value="[IMG]h33.gif[/IMG]">可 爱</option>
<option value="[IMG]h34.gif[/IMG]">无 话</option>
<option value="[IMG]h35.gif[/IMG]">大 声</option>
<option value="[IMG]h36.gif[/IMG]">流 汗</option>
<option value="[IMG]h37.gif[/IMG]">大 哭</option>
<option value="[IMG]h38.gif[/IMG]">投 降</option>
<option value="[IMG]h39.gif[/IMG]">吃批萨</option>
<option value="[IMG]h40.gif[/IMG]">被 打</option>
<option value="[IMG]h41.gif[/IMG]">小声点</option>
<option value="[IMG]h42.gif[/IMG]">咬 花</option>
<option value="[IMG]h43.gif[/IMG]">得 意</option>
<option value="[IMG]h44.gif[/IMG]">很 拽</option>
<option value="[IMG]h45.gif[/IMG]">发 火</option>
</select><select name="bgc" onchange="parent.f1.document.bgColor=parent.f0.document.bgColor=this.options[selectedIndex].value;this.value='#eeeeee';"  style="font-size:12px">
<option value="#eeeeee" selected>背景
<option value="FBE7DB" STYLE="background-color:#FBE7DB">淡橙
<option value="ffeaea" STYLE="background-color:#ffeaea">粉红
<option value="FCF8E2" STYLE="background-color:#FCF8E2">乳黄
<option value="eaffea" STYLE="background-color:#eaffea">浅绿
<option value="effaff" STYLE="background-color:#effaff">淡青
<option value="f2f2ff" STYLE="background-color:#f2f2ff">浅紫
<option value="eaeaff" STYLE="background-color:#eaeaff">蓝紫
<option value="000000" STYLE="background-color:#eaeaff">黑色
<option value="ffffff" STYLE="background-color:#ffffff">白色
<option value="f7f7f7" STYLE="background-color:#f7f7f7">默认
</select>
<select name="bgc1" onchange="parent.gg.document.bgColor=parent.f3.document.bgColor=parent.f2.document.bgColor=parent.CW_MENU.document.bgColor=this.options[selectedIndex].value;this.value='#<%=Application("aqjh_chatbgcolor")%>';"  style="font-size:12px">
<option value="#<%=Application("aqjh_chatbgcolor")%>" selected>右侧栏
<option value="<%=Application("aqjh_chatbgcolor")%>" STYLE="background-color:#<%=Application("aqjh_chatbgcolor")%>">默认</option>
<option value="ffeaea" STYLE="background-color:#ffeaea">粉红</option>
<option style="color:000000" value="000000">颜色1</option> 
<option style="color:0000FF" value="0000FF">颜色2</option> 
<option style="color:000088" value="000088">颜色3</option> 
<option style="color:888800" value="888800">颜色4</option> 
<option style="color:008888" value="008888">颜色5</option> 
<option style="color:008800" value="008800">颜色6</option> 
<option style="color:8888FF" value="8888FF">颜色7</option> 
<option style="color:AA00CC" value="AA00CC">颜色8</option> 
<option style="color:8800FF" value="8800FF">颜色9</option> 
<option style="color:888888" value="888888">颜色10</option> 
<option style="color:CCAA00" value="CCAA00">颜色11</option> 
<option style="color:FF8800" value="FF8800">颜色12</option> 
<option style="color:CC3366" value="CC3366">颜色13</option> 
<option style="color:FF00FF" value="FF00FF">颜色14</option> 
<option style="color:3366CC" value="3366CC">颜色15</option>
</select>
<input type=button name="tbclutch" value="全屏" onClick="javascript:parent.tbclutch();" title="合屏/分屏/垂直切换" onMouseOver="window.status='合屏/分屏/垂直切换。';return true" onMouseOut="window.status='';return true" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #008888; BORDER-RIGHT-WIDTH: 1px">  <input type=button value='僵' onClick="window.open('fafang/cins.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #8800ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='NPC' onClick="window.open('npc_list.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #cc0000; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">  <input type=button value='养猪' onClick="window.open('../hcjs/pig/zhu.asp','aqjh_win','scrollbars=0,resizable=0,width=670,height=400')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #cc0000; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">  <br>
<!功能三行>
<select name='command1' onchange="rc(this.value);document.af.command1.options[0].selected=true;" style='font-size:12px;'>
<option value="" selected>管理命令</option>
<%if aqjh_grade>=6 then%>
<option style="color:red" value="/公告$ ">官府公告
<option style="color:blue" value="/哑穴$ 请写明原因 ">哑穴操作
<option style="color:blue" value="/点穴$ 请写明原因 ">点穴操作
<option style="color:blue" value="/解穴$ 用户名">解穴操作
<option style="color:blue" value="/警告$ 请写明原因 ">警告操作
<option style="color:blue" value="/逮捕$ 请写明原因 ">逮捕操作
<option style="color:blue" value="/坐牢$ 请写明原因 ">坐牢操作
<option style="color:blue" value="/查ip$ ">查IP地址
<option style="color:blue" value="/查相同ip$ ">查相同IP
<option style="color:red" value="/解除$ 通缉犯名字">取消通辑
<option style="color:blue" value="/踢人$ 请写明原因 ">捣乱踢人
<option style="color:blue" value="/炸弹$ 找死啊你？">霹雳炸弹 <%
end if
if aqjh_grade>=8 then%>
<option style="color:red" value="/状态$ ">查看状态
<option style="color:blue" value="/监禁$ 请写明原因 ">监禁操作
<option style="color:blue" value="/释放$ 用户名">解除监禁
<option style="color:red" value="/清屏$ 操作">江湖清屏 <%
end if
if aqjh_grade>=9 then%>
<option style="color:red" value="/今日话题$ 祝大家玩得开心...">今日话题
<option style="color:red" value="/放大$ ">放大文字
<option style="color:red" value="/监禁ip$ 某一人捣乱，监禁所有ip相同用户！">监IP地址
<option style="color:red" value="/封锁ip$ 原因：在聊天室中多次捣乱！">封IP地址
<option style="color:red" value="/通缉$ 通缉犯名字">通缉人犯
<option style="color:red" value="/官府放怪$">官府发放
<option style="color:red" value="/斩首$ 原因 ">斩首人犯 <%
end if
if aqjh_grade>=10 then%>
<option style="color:red" value="/送金卡$ 1">赠送金卡</option>
<option style="color:red" value="/介绍奖励$ 1">介绍奖励</option>
<option style="color:red" value="/答题奖励$">答题奖励</option>
<option style="color:blue" value="/官府奖励$ 500000">官府奖励
<option style="color:blue" value="/册封花魁$ 20">册封花魁
<option style="color:blue" value="/册封状元$ 20">册封状元
<option style="color:blue" value="/册封盟主$">册封盟主
<option style="color:red" value="/查看物品$">查看物品
<option style="color:red" value="<script>parent.f3.location.href='savevalue.asp';</script>所有玩家全部存点!">全部存点
<option style="color:red" value="/站长发放$">站长发放
<option style="color:red" value="/原子弹$ 太不像话了">超级炸弹 <%end if%>
</select><input type=text name='clock' style="FONT-SIZE: 9pt; COLOR: #000000; BACKGROUND-COLOR: #c0c0c0; TEXT-ALIGN: right" value="" size=3 readonly> <input type=button value='巧' onClick="window.open('qiaozui.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #8800ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='牌' onClick="window.open('f2/xp.asp?name=<%=aqjh_name%>','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: ff5522; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"><%if zhuan>=10 or aqjh_grade>=9 then%> <input type=button value='爪' onClick="window.open('sjfunc/ckys.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #cc0000; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"><%end if%>
 <input type="checkbox" name="addvalues" checked=true onClick='if (<%=jhpd%>>=1){alert("此房间禁止泡点！");document.af.addvalues.checked=false;document.af.sytemp.focus();return false;}'><a href="#" onClick='if (<%=jhpd%>>=1){alert("此房间禁止泡点！");document.af.addvalues.checked=false;document.af.sytemp.focus();return false;};document.af.addvalues.checked=!(document.af.addvalues.checked);document.af.sytemp.focus()' style="color:red">泡</a> <input type=checkbox name='Zshow' value='1' accesskey='z' <%if aqjh_jhdj<90 then%>disabled<%end if%> onClick="document.af.Zshow.focus();"><a href='#' onClick='<%if aqjh_jhdj<90 then%>{alert("文字秀功能90级以下不可使用！")}return false;<%else%>document.af.Zshow.checked=!(document.af.Zshow.checked);<%end if%>document.af.Zshow.focus();' title='文字秀功能，90级以下不可用' style="color:red">秀</a> <input type='checkbox' name='py' accesskey='v'><a href=# onClick='document.af.py.checked=!(document.af.py.checked);' title="你的好友(配偶)上线、离开时是否自动通知您！" style="color:red">友</a> <input type='checkbox' name='mdsx' accesskey='b' checked><a href=# onClick='document.af.mdsx.checked=!(document.af.mdsx.checked);parent.m.location.reload();' title="当有人进入退出时是否自动刷新名单区" style="color:red">名</a>  <input onClick='document.af.dwtx.focus();' type=checkbox name=dwtx checked title='是否使用系统特效功能(如：震动、音乐等)！' value="ON"><a href='#' onClick='document.af.dwtx.checked=!(document.af.dwtx.checked);document.af.dwtx.focus();' title='是否使用系统特效功能(如：震动、音乐等)！' style="color:red">音</a>  <input type=checkbox name='towhoway' value='1' accesskey='s' <%if aqjh_jhdj<60 or (aqjh_grade>5 and aqjh_grade<9) then%>disabled<%end if%> onClick="document.af.sytemp.focus();"><a href='#' onClick='<%if aqjh_jhdj<60 or (aqjh_grade>5 and aqjh_grade<9) then%>{alert("很抱歉！此功能60级以下和官府不可用！")}return false;<%else%>document.af.towhoway.checked=!(document.af.towhoway.checked);<%end if%>document.af.sytemp.focus();' title='悄悄话儿悄悄说，60级以下和官府不可用' style="color:red">私</a> 
            <input type='checkbox' name=as accesskey='a' checked onClick='parent.f1.scrollit();document.af.sytemp.focus();' value="ON"><a href='#' onClick='document.af.as.checked=!(document.af.as.checked);document.af.as.focus();parent.f1.scrollit();document.af.sytemp.focus();' title='是否滚屏显示; ' style="color:red">滚</a>  
            <input type='checkbox' name='ctz' accesskey='v' value="OFF"><a href=# onMouseOver="window.status='在发言时，是否显示粗体字与个人头像？';return true" onMouseOut="window.status='快乐江湖首创，您是掌门就可以使用！';return true"  onClick='document.af.ctz.checked=!(document.af.ctz.checked);' title="在发言时，是否显示粗体字与个人头像？" style="color:red">粗</a>  
           <%if aqjh_grade>=6 then%>  
            <a href="#" onClick="window.open('../twt/dt/ask.asp','ask','scrollbars=no,resizable=no,width=400,height=300')" style="color:red">题</a>    <a class=blue href="#" onClick="window.open('../twt/dalie/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要官府放猎物了，你可以打猎物得钱！" style="color:red">猎</a>                                                                     
            <a class=blue href="#" onClick="window.open('../twt/jiaofei/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要官府放匪了，你可以打匪得钱！" style="color:red">匪</a>                                                                     
     <%else%>                                                                    
            <a href="hy.asp" target="_blank" style="color:red">会</a>                                                                    
            <a class=blue href="#" onClick="window.open('../twt/dt/reply.asp','daliwu','scrollbars=no,resizable=no,width=490,height=600')" title="来答字花吧答对有银两可赚喔！" style="color:red">答</a>                                                                     
            <a class=blue href="#" onClick="window.open('../twt/dalie/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要官府放猎物了，你可以打猎物得钱！" style="color:red">猎</a>                                                                    
            <a class=blue href="#" onClick="window.open('../twt/jiaofei/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要官府放匪了，你可以打匪得钱！" style="color:red">匪</a><%end if%> <%if chatinfo(7)=0 then%>                                                                    
 <a href=lg.asp title="在练功的时间，你与别人是不能打架的！" target="d" style="color:blue">护</a>
<%end if%> 
            <%if instr("," & Application("aqjh_slbox") & ",","," & aqjh_name & ",")<>0 then%>
            <input type='checkbox' name='slbox' accesskey='s' onClick="document.af.sytemp.focus();if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1}" value="ON"><a href='#' onMouseOver="window.status='选中本项后，可以查看到聊天室中的私聊！。';return true" onMouseOut="window.status='';return true" onClick="document.af.slbox.checked=!(document.af.slbox.checked);if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1};document.af.sytemp.focus();" title="看一看大家在说一些什么悄悄话" style="color:blue">天</a>                                                                     
            <%end if%>
<script language="JavaScript">function startnosay(){var nosay=<%=Application("aqjh_maxtimeout")*60%>;document.af.clock.value=nosay;}</script>                                                                    
            <script src="readonly/f2.js"></script>                                                                    
            <script>parent.fc();parent.m.location.href="f3.asp";</script>                               
</DIV></BODY></HTML>