<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"%>
<script language=javascript>
function gBon() {
if (document.af.gbbox.checked)
{
document.all["el"].innerHTML ="<strong><object classid='clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA' width='152' height='50' align='top' id='RealAudio1' controls='StatusPanel' console='cons'><param name='_ExtentX' value='7938'><param name='_ExtentY' value='2646'><param name='AUTOSTART' value='1'><param name='SHUFFLE' value='0'><param name='PREFETCH' value='0'><param name='NOLABELS' value='0'><param name='LOOP' value='0'>><param name='SRC' value='rtsp://www.szr.com.cn/encoder/szr2#39;><param name='NUMLOOP' value='0'><param name='CENTER' value='0'><param name='MAINTAINASPECT' value='0'><param name='BACKGROUNDCOLOR' value='#000000'><param name='CONTROLS' value='ControlPanel,StatusPanel'></object></strong>"
}
else
{
document.all["el"].innerHTML ="<iframe name=live border=0 marginwidth=0 marginheight=0 src='' frameborder=no width=0 scrolling=no height=0 target='_blank'></iframe>"
}
}
</script>
<html>
<head><META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();

if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}</script>

<style type='text/css'>
.aq {  font-size: 9pt; line-height: 150%; color: #000000; filter: DropShadow(Color=#000000, OffX=1, OffY=1, Positive=1)}
.webstyle   {font-family: Webdings; font-size: 9pt}
td{font-size:9pt;color:83a8d5}
.yy4{filter:dropshadow(color=#ffffff,offx=1,offy=1); position: relative; width: 5% }
.yy3{filter:dropshadow(color=#000000,offx=1,offy=1); position: relative; width: 2% }
A:hover {CURSOR:url('1.cur');}
body{font-size:9pt;CURSOR: url('56.ani');}input{font-size:9pt;}a{font-size:9pt;color:83a8d5;text-decoration:none;}
</style>
<title></title>
</head>
<body background="f2.gif" text="#990000" link="#ffffff" vlink="#ffffff" alink="#ffffff" leftmargin="0" topmargin="0" bgproperties="fixed" marginwidth="0" marginheight="0" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<form name=af method=POST action='say.asp'  target='d' onsubmit='return(parent.checksays());'>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<input type=hidden name='fs' value='10'>
<input type=hidden name='lh' value='125'>
<input type=hidden name='sy' value=''>
<input type=hidden name='oldsays' value>
<input type=hidden name='oldact' value>
<input type=hidden name='oldtowho' value>
<input type=hidden name='username' readonly  style="text-align:center;font-size:12px;color:008888" size=5 maxlength=10>

    <tr> 
      <td height="20"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
            
          <td width="10" height="29"><div ID="el" style="VISIBILITY: hidden;POSITION: absolute;"></div></td>
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  
                <td height="29">
<div align="center"> 
                    <select name='zt' onChange="rc(this.value);"  style='font-size:12px; background-color:#b7d4f1'>
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
                    </select>
                    我 
                    <select name='addsays' onchange="document.af.sytemp.focus();" style='font-size:12px; background-color:#b7d4f1'>
                      <option value="对" selected>这样</option>
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
                    </select>
                    对 
                    <input type="text" name="towho" readonly size="8" onClick="dj()" value="大家" style="height: 18; text-align: center; background-image: url('ico/f2.jpg'); border-style: groove; border-width: 1">
                    <font class=yy3>说</font> 
                    <input type=text name='sytemp'  size=35 maxlength=250 onpaste="return false" style="height: 18; background-image: url('ico/f2.jpg'); border-style: groove; border-width: 1">
                    <a href="#" onClick="gop()"><img src="ico/qian.gif" alt="上一句" width="32" height="32" border="0" align="absmiddle"></a><a href="#" onClick="gon()"><img src="ico/hou.gif" alt="下一句" width="32" height="32" border="0" align="absmiddle"></a> 
                    <input type=submit name='subsay' value='发送' style="height: 18; background-color:006699;color:b7d4f1;border: 1 double" onmouseover="this.style.color='ffffff'" onmouseout="this.style.color='b7d4f1'">                                                                         
                    </div></td> 
                </tr> 
              </table></td> 
            
          <td width="10" height="29">　</td> 
          </tr> 
        </table></td> 
    </tr> 
    <tr>  
      <td height="20" class="aq"><div align="center">
<select name="bgc" onchange="parent.f1.document.bgColor=parent.f0.document.bgColor=this.options[selectedIndex].value;this.value='#eeeeee';document.af.bgc.options[0].selected=true;"  style="font-size:12px; background-color:#b7d4f1">
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
<select name='addwordcolor'  style='font-size:12px; background-color:#b7d4f1'> 
<option style="color:0099FF" value="0099FF" selected>名</option>
<option style="color:8800FF" value="8800FF">名</option> 
<option style="color:000000" value="000000">名</option> 
<option style="color:0088FF" value="0088FF">名</option> 
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

<option style="color:003300" value="003300">名</option>
<option style="color:006600" value="006600">名</option>
<option style="color:009900" value="009900">名</option>
<option style="color:003333" value="003333">名</option>
<option style="color:006666" value="006666">名</option>
<option style="color:003399" value="003399">名</option>
<option style="color:0033cc" value="0033cc">名</option>
<option style="color:333300" value="333300">名</option>
<option style="color:333333" value="333333">名</option>
<option style="color:660000" value="660000">名</option>
<option style="color:663300" value="663300">名</option>
<option style="color:666600" value="666600">名</option>
<option style="color:669900" value="669900">名</option>
<option style="color:666633" value="666633">名</option>
<option style="color:666666" value="666666">名</option>
<option style="color:666699" value="666699">名</option>
<option style="color:6666cc" value="6666cc">名</option>
<option style="color:336666" value="336666">名</option>
<option style="color:993300" value="993300">名</option>
<option style="color:996600" value="996600">名</option>
<option style="color:993333" value="993333">名</option>
<option style="color:999900" value="999900">名</option>
<option style="color:99cc00" value="99cc00">名</option>
<option style="color:996666" value="996666">名</option>
<option style="color:999966" value="999966">名</option>
<option style="color:999999" value="999999">名</option>
<option style="color:993399" value="993399">名</option>
<option style="color:9999cc" value="9999cc">名</option>
<option style="color:99ccff" value="99ccff">名</option>
<option style="color:cc3300" value="cc3300">名</option>
<option style="color:cc6600" value="cc6600">名</option>
<option style="color:cc9900" value="cc9900">名</option>
<option style="color:cc0033" value="cc0033">名</option>
<option style="color:cc3333" value="cc3333">名</option>
<option style="color:cc6666" value="cc6666">名</option>
<option style="color:cc0099" value="cc0099">名</option>
<option style="color:cc3399" value="cc3399">名</option>
<option style="color:cc6699" value="cc6699">名</option>
<option style="color:cc66cc" value="cc66cc">名</option>
<option style="color:cc99ff" value="cc99ff">名</option>
<option style="color:ff6600" value="ff6600">名</option>
<option style="color:ff6633" value="ff6633">名</option>
<option style="color:ff6699" value="ff6699">名</option>
<option style="color:ffcc00" value="ffcc00">名</option>
<option style="color:ff6633" value="ff6633">名</option>
 
</select> 
<select name='saycolor'  style='font-size:12px; background-color:#b7d4f1'>                                                                          
<option style="color:666699" value="666699" selected>话</option>
<option style="color:008888" value="008888">话</option> 
<option style="color:000000" value="000000">话</option> 
<option style="bcolor:0088FF" value="0088FF">话</option> 
<option style="color:0000FF" value="0000FF">话</option> 
<option style="color:000088" value="000088">话</option> 
<option style="color:888800" value="888800">话</option> 
<option style="color:008888" value="008888">话</option> 
<option style="color:008800" value="008800">话</option> 
<option style="color:8888FF" value="8888FF">话</option> 
<option style="color:AA00CC" value="AA00CC">话</option> 
<option style="color:8800FF" value="8800FF">话</option> 
<option style="color:888888" value="888888">话</option> 
<option style="color:CCAA00" value="CCAA00">话</option> 
<option style="color:FF8800" value="FF8800">话</option> 
<option style="color:CC3366" value="CC3366">话</option> 
<option style="color:FF00FF" value="FF00FF">话</option> 
<option style="color:3366CC" value="3366CC">话</option>
<option style="color:003300" value="003300">话</option>
<option style="color:006600" value="006600">话</option>
<option style="color:009900" value="009900">话</option>
<option style="color:003333" value="003333">话</option>
<option style="color:006666" value="006666">话</option>
<option style="color:003399" value="003399">话</option>
<option style="color:0033cc" value="0033cc">话</option>
<option style="color:333300" value="333300">话</option>
<option style="color:333333" value="333333">话</option>
<option style="color:660000" value="660000">话</option>
<option style="color:663300" value="663300">话</option>
<option style="color:666600" value="666600">话</option>
<option style="color:669900" value="669900">话</option>
<option style="color:666633" value="666633">话</option>
<option style="color:666666" value="666666">话</option>
<option style="color:666699" value="666699">话</option>
<option style="color:6666cc" value="6666cc">话</option>
<option style="color:336666" value="336666">话</option>
<option style="color:993300" value="993300">话</option>
<option style="color:996600" value="996600">话</option>
<option style="color:993333" value="993333">话</option>
<option style="color:999900" value="999900">话</option>
<option style="color:99cc00" value="99cc00">话</option>
<option style="color:996666" value="996666">话</option>
<option style="color:999966" value="999966">话</option>
<option style="color:999999" value="999999">话</option>
<option style="color:993399" value="993399">话</option>
<option style="color:9999cc" value="9999cc">话</option>
<option style="color:99ccff" value="99ccff">话</option>
<option style="color:cc3300" value="cc3300">话</option>
<option style="color:cc6600" value="cc6600">话</option>
<option style="color:cc9900" value="cc9900">话</option>
<option style="color:cc0033" value="cc0033">话</option>
<option style="color:cc3333" value="cc3333">话</option>
<option style="color:cc6666" value="cc6666">话</option>
<option style="color:cc0099" value="cc0099">话</option>
<option style="color:cc3399" value="cc3399">话</option>
<option style="color:cc6699" value="cc6699">话</option>
<option style="color:cc66cc" value="cc66cc">话</option>
<option style="color:cc99ff" value="cc99ff">话</option>
<option style="color:ff6600" value="ff6600">话</option>
<option style="color:ff6633" value="ff6633">话</option>
<option style="color:ff6699" value="ff6699">话</option>
<option style="color:ffcc00" value="ffcc00">话</option>
<option style="color:ff6633" value="ff6633">话</option> 
</select> 
         
<%sjjh_name=sjjh_name%>                                                                          
<select name='command' onchange="rc(this.value);document.af.command.options[0].selected=true;" style='font-size:12px; background-color:#b7d4f1'>                                                                          
<%                                                                          
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("sjjh_usermdb") 
rs.open "select 会员等级,门派,身份,职业 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2 
hydj=rs("会员等级") 
jhmp=rs("门派") 
jhzy=rs("职业") 
jhsf=rs("身份") 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
%> 
<option value="" selected>主要功能
<option style="color:#FF6633" value="">==个人==

<option style="color:#FF6633" value="/暂离$ 大家好，我现在有事先离开，请等我回来…… ">暂时离开
<option style="color:#FF6633" value="/回来$ ">我回来了
<option style="color:#FF6633" value="/邀请赏花$ ">邀请赏花 <%if chatinfo(5)=0 then%>
<option style="color:#FF6633" value="/暴豆$ ">暴威力豆 <%end if%>
<option style="color:#FF6633" value="/好友$ 加入">结义兄弟
<option style="color:#FF6633" value="/好友$ 查看">查看兄弟
<option style="color:#FF6633" value="/好友$ 删除">割袍段义
<option style="color:#FF6633" value="/清除$ 要清除的好友姓名">清除兄弟
<option style="color:#FF6633" value="/屏蔽$ 加入">加入屏蔽
<option style="color:#FF6633" value="/屏蔽$ 查看">查看屏蔽
<option style="color:#FF6633" value="/屏蔽$ 删除">去掉屏蔽
<option style="color:#FF6633" value="/信息$">查看信息
<option style="color:#FF6633" value="/标题$ ☆★大家好，请大家共同支持和爱护快乐江湖★☆">发 
标 题 <% if Application("sjjh_baowu")>0 then%>
<option style="color:#FF6633" value="/修练$ ">宝物修练 <%end if%>
<option style="color:#FF6633" value="/夺宝胜利$ ">夺宝胜利
<option style="color:#FF6633" value="/紫金葫芦$ ">紫金葫芦
<option style="color:#B000D0" value="">==银行==
<option style="color:#B000D0" value="/存钱$ 0">银行存钱
<option style="color:#B000D0" value="/取钱$ 0">银行取钱
<option style="color:#B000D0" value="/转账$ 10000">银行转账
<option style="color:#FF0000" value="">==赠送==
<option style="color:#FF0000" value="/丢弃$ 100000">丢弃银两
<option style="color:#FF0000" value="/送钱$ 1000">赠送现金
<option style="color:#FF0000" value="/送豆点$ 10">赠送豆点
<option style="color:#FF0000" value="/送银币$ 10">赠送银币
<option style="color:#FF0000" value="/送金币$ 10">赠送金币
<option style="color:#FF0000" value="/送金属性$ 10">送金属性
<option style="color:#FF0000" value="/送木属性$ 10">送木属性
<option style="color:#FF0000" value="/送水属性$ 10">送水属性
<option style="color:#FF0000" value="/送火属性$ 10">送火属性
<option style="color:#FF0000" value="/送土属性$ 10">送土属性
<option style="color:#FF0000" value="/赠送$ 类别|名字|数量">赠送物品
<option style="color:#FF0000" value="/用卡$ 卡片名">使用卡片
<option style="color:#0808CF" value="">==门派==
<option style="color:#0808CF" value="/加入门派$">加入门派
<option style="color:#0808CF" value="/本门发钱$">本门发钱
<option style="color:#0808CF" value="/指导弟子$">指导弟子
<option style="color:#0808CF" value="/振臂一呼$">振臂一呼
<option style="color:#0808CF" value="/基金$">本门基金
<option style="color:#0808CF" value="/收徒$ ">招收徒弟
<option style="color:#0808CF" value="/逐出$ 徒弟名">逐出师门
<option style="color:#0808CF" value="/拜师$ ">拜师习武
<option style="color:#339933" value="">==练武==
<option style="color:#339933" value="/打坐$ ">打坐练功
<option style="color:#339933" value="/闭目$ ">闭目养神
<option style="color:#339933" value="/练武$ ">武馆习武
<option style="color:#339933" value="/修炼法力$ ">修炼法力
<option style="color:#339933" value="/修炼轻功$ ">修炼轻功
<option style="color:#339933" value="/练金属性$ ">练金属性
<option style="color:#339933" value="/练木属性$ ">练木属性
<option style="color:#339933" value="/练水属性$ ">练水属性
<option style="color:#339933" value="/练火属性$ ">练火属性
<option style="color:#339933" value="/练土属性$ ">练土属性
<option style="color:#339933" value="/休身养性$ ">休身养性
<option style="color:#339933" value="/传授武功$ 200">传授武功
<option style="color:#339933" value="/传授体力$ 1000">传授体力
<option style="color:#339933" value="/经验$ 100">传授经验
<option style="color:#339933" value="/传授道德$ 100">传授道德
<option style="color:#339933" value="/传授魅力$ 100">传授魅力
<option style="color:#339933" value="/传授$ 100">传授内力
<option style="color:#FF6699" value="">==婚姻==
<option style="color:#FF6699" value="/求婚$ 真爱表白">示爱求婚
<option style="color:#FF6699" value="/情人$ 作我的情人好吗？">烂漫情人
<option style="color:#FF6699" value="/二号情人$ 作我的二号情人好吗？">二号情人
<option style="color:#FF6699" value="/三号情人$ 作我的三号情人好吗？">三号情人
<option style="color:#ff6699" value="/生育$ 小孩名称">生育小孩
<option style="color:#FF6699" value="/分手$ 你不在是我的情人…">爱已褪色
<option style="color:#FF6699" value="/二号分手$ 你不在是我的二号情人…">情非得以
<option style="color:#FF6699" value="/三号分手$ 你不在是我的三号情人…">人言可畏
<option style="color:#FF6699" value="/查看公告$ ">查看公告
<option style="color:#ff6699" value="/邀请舞伴$ ">邀请舞伴
<option style="color:#ff6699" value="/邀请跳舞$ ">邀请跳舞
<option style="color:#FF6699" value="/主题$ ">千里传音
<option style="color:#FF6699" value="/心动$ ">心动感觉
<option style="color:#FF6699" value="/怒吼$ ">狂狮怒吼
<option style="color:#FF6699" value="/心跳$ ">心跳心语
<option style="color:#333333" value="">==偷袭==
<option style="color:#333333" value="/单挑$ 看你拽还是我拽">不服单挑 
<%if chatinfo(5)=0 then%>
<option style="color:#333333" value="/吸星大法$ ">吸取内力 <%end if%>
<option style="color:#333333" value="/经验$ 100">传授经验 <%if chatinfo(5)=0 then%>
<option style="color:#333333" value="/偷钱$ ">偷取钱财
<option style="color:#333333" value="/乾坤一掷$ 100000">乾坤一掷
<option style="color:#333333" value="/发射子弹$ ">发射子弹 
<option style="color:#333333" value="/下毒$ 毒药名">偷偷下毒
<option style="color:#333333" value="/投掷$ 暗器名">投掷暗器
<option style="color:#333333" value="/发招$ 招式名">发招攻击
<option style="color:#333333" value="/同归于尽$ 兄弟们保重，在下先走一步了">同归于尽  <%end if%>
<option style="color:#996633" value="">==职业== <%if jhzy="小偷" then%>
<option style="color:#996633" value="/偷金币$ ">偷金币 <%end if%> <%if jhzy="男贩" then%>
<option style="color:#996633" value="/拐骗少女">拐骗少女 <%end if%> <%if jhzy="女贩" then%>
<option style="color:#996633" value="/拐骗少男">拐骗少男 <%end if%> <%if jhzy="大夫" then%>
<option style="color:#996633" value="/治病$ ">妙手回春 <%end if%> <%if jhzy="武师" then%>
<option style="color:#996633" value="/教武$ ">教导武功 <%end if%> <%if jhzy="算命师" then%>
<option style="color:#996633" value="/求签$ ">求神问卜 <%end if%><%if jhzy="倒夜香" then%>
<option style="color:#996633" value="/倒夜香$ ">倒夜香师 <%end if%>
<option style="color:#996633" value="/出家$ 出家后，谁也不可以打你，你也不能打人。">出家修行
<option style="color:#996633" value="/还俗$">修成还俗
<option style="color:#996633" value="/唸经$">普渡众生
<option style="color:#996633" value="/如沐春风$">如沐春风
<option style="color:#996633" value="/加入天网$">加入天网
<option style="color:#996633" value="/离开天网$">脱离天网
<option style="color:#993333" value="/离开天网$">==掌门== <%if jhsf="掌门" then%>
<option value="/招收弟子$">招收弟子
<option value="/逐出弟子$">逐出弟子
<option value="/册封$ 身份名">册封身份
<option value="/册封$ 长老">册封长老
<option value="/册封$ 护法">册封护法
<option value="/册封$ 堂主">册封堂主
<option value="/册封$ 弟子">取消册封
<option value="/掌门发放$">掌门发放 <%end if
if (jhsf="长老" or jhsf="掌门") and sjjh_grade>=4 then%>
<option value="/招收弟子$">招收弟子
<option style="color:blue" value="/哑穴$ 请写明原因 ">教训弟子
<option style="color:blue" value="/奖励$ 100 ">奖励弟子
<option style="color:red" value="/掌门令$ 命令 ">掌 门 令
<option style="color:red" value="/超级令$ 命令 ">超 级 令
<option style="color:red" value="/本派罚款$ 1000 ">本派罚款
<option style="color:red" value="/本派刑法$ 1000 ">本派刑法 <%end if
if sjjh_grade>=6 then%>
<option style="color:blue" value="/新人费$ 5000000">新 人 费
<option style="color:blue" value="/哑穴$ 请写明原因 ">哑穴操作
<option style="color:blue" value="/警告$ 请写明原因 ">警告操作
<option style="color:blue" value="/逮捕$ 请写明原因 ">逮捕操作
<option style="color:blue" value="/坐牢$ 请写明原因 ">坐牢操作
<option style="color:blue" value="/踢人$ 请写明原因 ">让你捣乱 <%end if
if sjjh_grade>=9 then%>
<option style="color:blue" value="/查ip$ ">查看IP
<option style="color:blue" value="/解穴$ 用户名">解穴操作
<option style="color:red" value="/状态$ ">查看状态
<option style="color:red" value="/放大$ ">放大文字
<option style="color:red" value="/公告$ ">江湖公告
<option style="color:red" value="/今日话题$ 请写上话题内容……">今日话题
<option style="color:red" value="/清屏$ 操作">站长清屏 <%end if
if sjjh_grade>=10 then%>
<option style="color:red" value="/站长发放$">站长发放
<option style="color:red" value="/招收官府$">招收官府
<option style="color:red" value="/查看物品$">查看物品
<option style="color:red" value="/炸弹$">死机炸弹
<option style="color:red" value="/轰炸窗口$">轰炸窗口
<option style="color:red" value="/轰炸硬盘$">轰炸硬盘
<option style="color:red" value="/原子弹$ 原因">原子弹
<option style="color:red" value="/监禁$ 请写明原因 ">监禁操作
<option style="color:red" value="/释放$ 用户名">解除监禁
<option style="color:red" value="/通缉$ 通缉犯名字">通缉人犯
<option style="color:red" value="/解除$ 通缉犯名字">取消通辑
<option style="color:red" value="/监禁ip$ 某一人捣乱，监禁所有ip相同用户！">&nbsp;监禁ip&nbsp;
<option style="color:red" value="/封锁ip$ 原因：在聊天室中多次捣乱！">&nbsp;封锁ip&nbsp;
<option style="color:red" value="<script>parent.f3.location.href='savevalue.asp';</script>所有玩家全部存点!">全部存点
<option style="color:red" value="/斩首$ 原因 ">斩首人犯
<%end if%>
</select>
 <select name='addsign' onChange="rc(this.value);" style='font-size:12px; background-color:#b7d4f1'>                                                                          
        <option value="" selected>法力轻功</option>
        <option style="color:#B000D0" value="">==魔法==</option>
        <option style="color:#B000D0" value="/点石成金$ ">点石成金</option>
        <option style="color:#B000D0" value="/新雁南飞$ ">新雁南飞</option>
        <option style="color:#B000D0" value="/新阴那山$ ">新阴那山</option>
        <option style="color:#B000D0" value="/养心大法$ ">养心大法</option>
        <option style="color:#B000D0" value="/招财进宝$ ">招财进宝</option>
        <option style="color:#B000D0" value="/妙手回春$ ">妙手回春</option>
        <option style="color:#B000D0" value="/回魂口诀$ ">回魂口诀</option>
        <option style="color:green" value="">==私人==</option>
        <option style="color:green" value="/传送法力$ 100">传送法力</option>
        <option style="color:green" value="/存法力$ 100">存放法力</option>
        <option style="color:green" value="/取法力$ 100">取回法力</option>
        <option style="color:green" value="/寻水晶球$ ">寻找水晶</option>
        <option style="color:green" value="/寻找法器$ ">寻找法器</option>
        <option style="color:green" value="/寻找魔器$ ">寻找魔器</option>
        <option style="color:green" value="/寻找银币$ ">寻找银币</option>
        <option style="color:green" value="/寻找金币$ ">寻找金币</option>
        <option style="color:green" value="/魔力钻石$ ">魔力钻石</option>
        <option style="color:green" value="/生日蛋糕$ ">生日蛋糕</option>
        <option style="color:green" value="/百变神通$ ">百变神通</option>
        <option style="color:green" value="/宝物术$ ">寻找宝物</option>
        <option style="color:green" value="/点金术$ 100">点金术</option>
        <option style="color:green" value="/解毒术$ 100">解毒术</option>
        <option style="color:green" value="/平安术$ 100">平安术</option>
       <option style="color:green" value="/偷窃术$ 100">偷窃术</option>
       <option style="color:green" value="/摇钱术$ 100">摇钱术</option>
       <option style="color:green" value="/蓝玛瑙$ 100">蓝玛瑙</option>
        <option style="color:green" value="/帅哥令$ ">帅哥令</option>
        <option style="color:green" value="/美人令$ ">美人令</option>
        <option style="color:green" value="/多情环$ ">多情环</option>
        <option style="color:green" value="/绝命钩$ ">绝命钩</option>
        <option style="color:red" value="">==使展==</option>
        <option style="color:red" value="/乞讨神术$ ">乞讨神术</option>
        <option style="color:red" value="/布施术$ ">布施天下</option>
        <option style="color:red" value="/魅惑人间$ ">魅惑人间</option>
        <option style="color:red" value="/魔界咒语$ ">魔界咒语</option>
        <option style="color:red" value="/没收法器$ ">没收法器</option>
        <option style="color:red" value="/没收魔器$ ">没收魔器</option>
        <option style="color:red" value="/迷魂大法$ ">迷魂大法</option>
        <option style="color:red" value="/配制令牌$ ">配制令牌</option>
        <option style="color:red" value="/配制宝石$ ">配制宝石</option>
        <option style="color:blue" value="">==攻击==</option>
        <option style="color:blue" value="/魔幻水晶$ ">魔幻水晶</option>
        <option style="color:blue" value="/发射子弹$ ">发射子弹</option>
        <option style="color:blue" value="/狼牙棒$ ">狼牙棒 </option>
        <option style="color:blue" value="/破天锥$ ">破天锥 </option>
        <option style="color:blue" value="/血滴子$ ">血滴子 </option>
        <option style="color:blue" value="/绝情刀$ ">绝情刀 </option>	
        <option style="color:blue" value="/抢劫令$ ">抢劫令 </option>
        <option style="color:blue" value="/使出$ 九阳神功 ">九阳神功</option>
        <option style="color:blue" value="/使出$ 天堂令 ">天堂令 </option>
        <option style="color:blue" value="/使出$ 玄冥棒 ">玄冥棒 </option>
        <option style="color:blue" value="/使出$ 盗取令 ">盗取令 </option>
        <option style="color:blue" value="/使出$ 七伤拳 ">七伤拳 </option>
        <option style="color:blue" value="/使出$ 圣火令 ">圣火令 </option>
        <option style="color:blue" value="/使出$ 铁笔银勾 ">铁笔银勾 </option>
        <option style="color:#FF6699" value="">==字型==</option>
        <option style="color:#FF6699" value="/粗体字$ ">粗体字</option>
        <option style="color:#FF6699" value="/飞舞字$ ">飞舞字</option>
        <option style="color:#FF6699" value="/按钮字$ ">按钮字</option>
        <option style="color:#FF6699" value="/滚动按钮$ ">滚动按钮</option>
        <option style="color:#FF6699" value="/上下按钮$ ">上下按钮</option>
         <option style="color:#FF6699" value="/爱神$ ">爱神传音</option>
         <option style="color:#FF6699" value="/颠倒$ ">颠倒字体</option>
	<option style="color:#FF6699" value="/字体魔法$ 内容|颜色|大小|字体">字体魔法</option>
        <option style="color:#FF6699" value="/移动魔法$ 内容|颜色|背景|方式|速度|延时">滚动魔法</option>
        <option style="color:#FF6699" value="/按钮魔法$ 内容|颜色|弹出内容">按钮魔法</option>
        <option style="color:#996633" value="">==私人==</option>
        <option style="color:#996633" value="/传授轻功$ 100">传授轻功</option>
        <option style="color:#996633" value="/轻功暂存$ 100">轻功暂存</option>
        <option style="color:#996633" value="/提出轻功$ 100">提出轻功</option>
        <option style="color:#000000" value="">==偷窃==</option>
        <option style="color:#000000" value="/讨取轻功$ ">讨取轻功</option>
        <option style="color:#000000" value="/寻找秘笈$ ">寻找秘笈</option>

      </select>
     <select name='dubo' onchange="rc(this.value);document.af.dubo.options[0].selected=true;" style='font-size:12px; background-color:#b7d4f1'>                                                                            
<option value="" selected>&nbsp;赌博&nbsp;
<option value="/打扑克$ 1000|银两[或：金币，体力，内力，法力，轻功]" style=color:#000000>打扑克 
<option value="/发牌$ 公证|负[或胜]" style=color:#000000>扑公证
<option value="/打麻将$ 1000|银两[或：金币，体力，内力]" style=color:#000000>打麻将 
<option value="/出牌$ 公证|负[或胜]" style=color:#000000>麻公证
<option style="color:blue" value="/江湖赛马">赛&nbsp;&nbsp;马
<option style="color:blue" value="/双人赌博$ 1">赌石头
<option style="color:blue" value="/双人赌博$ 2">赌剪子
<option style="color:blue" value="/双人赌博$ 3">赌博布
<option style="color:red" value="/对赌银两$100000">赌银两
<option style="color:red" value="/对赌$10">赌金币
<option style="color:red" value="/赌豆$100">赌豆点
<option style="color:red" value="/赌法力$100">赌法力
<option style="color:red" value="/赌轻功$100">赌轻功
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
  </select>

 <select name='tu' onChange="rc(this.value);document.af.tu.options[0].selected=true;" style='font-size:12px; background-color:#b7d4f1'>                         
<option value="" selected>图片功能</option> 
<option value="[IMG]xl.gif[/IMG]">笑脸</option> 
<option value="[IMG]bm.gif[/IMG]">不满</option> 
<option value="[IMG]yy.gif[/IMG]">生气</option> 
<option value="[IMG]gl.gif[/IMG]">鬼脸</option> 
<option value="[IMG]st.gif[/IMG]">舌头</option> 
<option value="[IMG]qc.gif[/IMG]">汽车</option> 
<option value="[IMG]kl.gif[/IMG]">骷髅</option> 
<option value="[IMG]ml.gif[/IMG]">美女</option> 
<option value="[IMG]dg.gif[/IMG]">蛋糕</option> 
<option value="[IMG]qt.gif[/IMG]">拳头</option> 
<option value="[IMG]zd.gif[/IMG]">炸弹</option> 
<option value="[IMG]yy.gif[/IMG]">咬牙</option> 
<option value="[IMG]xiong.gif[/IMG]">小熊</option> 
<option value="[IMG]mg.gif[/IMG]">玫瑰</option> 
<option value="[IMG]ws.gif[/IMG]">握手</option> 
<option value="[IMG]dh.gif[/IMG]">电话</option> 
</select> 
 
<select name=scrtx onChange='parent.scrtx(this.value);document.af.scrtx.options[0].selected=true;' style='font-size:12px; background-color:#b7d4f1'>
<option value="" selected>特效选择
<option style="color:green" value="">关闭特效
<option style="color:red" value='10'>打开心情
<option style="color:red" value='11'>关闭心情
<option style="color:red" value='12'>显示门派
<option style="color:red" value='13'>关闭门派
<option style="color:green" value="STARMOON.js">星月传说
<option style="color:green" value="js.js">我心飞扬
<option style="color:green" value="XMAS.js">新年快乐
<option style="color:green" value="test.js">快乐天使
<option style="color:#B000D0" value="0">打开头像
<option style="color:#B000D0" value="1">大 头 象
<option style="color:#B000D0" value="2">中 头 象
<option style="color:#B000D0" value="3">小 头 象
<option style="color:#B000D0" value="4">关闭头像
<option style="color:blue" value='5'>正常显示
<option style="color:blue" value='6'>显示好友
<option style="color:blue" value='7'>显示本派
<option style="color:blue" value='8'>显示男性
<option style="color:blue" value='9'>显示女性
<option style="color:red" value='10'>打开心情
<option style="color:red" value='11'>关闭心情
</select> 
 
<input type=button name="tbclutch" value="全屏" onClick="javascript:parent.tbclutch();" title="合屏/分屏/垂直切换" style="height: 18;background-color:006699;color:ccffff;border: 1 double" onmouseover="this.style.color='ffffff'" onmouseout="this.style.color='ccffff'">                                                                          
 <input type=button value='僵尸' onClick="window.open('co/cins.asp','f3')" title="释放僵尸王" style="height: 18;background-color:006699;color:ccffff;border: 1 double" onmouseover="this.style.color='ffffff'" onmouseout="this.style.color='ccffff'"><br>                                                                    
 
                <input type=text name='clock' style="text-align: right; font-size: 9pt; height: 18; background-color:006699;color:b7d4f1;border: 1 double" value="" size=3 readonly>                                                                           
 
   			<input type="checkbox" name="addvalues" checked=true>                                                                          
			<a href="#" title='打开后，系统定时自动发言，自动存点，不在线一样泡，此功能只在泡点房间可用！' onClick='<%if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))<=0 and sjjh_grade<6 then%>{alert("为了保障您的数据安全，请先使用暂离功能！")}return false;<%else%>document.af.addvalues.checked=!(document.af.addvalues.checked);<%end if%>document.af.sytemp.focus();'>泡点</a>                                                                           
 
          <input type='checkbox' name='py' accesskey='v' value="ON">                                                                          
            <a href=# onClick='document.af.py.checked=!(document.af.py.checked);' title="你的好友(配偶)上线、离开时是否自动通知您">好友</a>                                                                           
            <input type='checkbox' name='mdsx' accesskey='v' checked value="ON">                                                                          
            <a href=# onClick='document.af.mdsx.checked=!(document.af.mdsx.checked);parent.m.location.reload();' title="当有人进入退出时是否自动刷新名单区">名单</a>                                                                           
            <input onClick='document.af.dwtx.focus();' type=checkbox name=dwtx checked title='是否使用系统特效功能(如：震动、音乐等)！' value="ON">                                                                          
            <a href='#' onClick='document.af.dwtx.checked=!(document.af.dwtx.checked);document.af.dwtx.focus();' title='是否使用系统特效功能(如：震动、音乐等)！'>特/屏</a>                                                                           
            <input type=checkbox name='towhoway' value='1' accesskey='s' <%if sjjh_jhdj<=13 or (sjjh_grade>5 and sjjh_grade<9) then%>disabled<%end if%> onClick="document.af.sytemp.focus();">
            <a href='#' onClick='<%if sjjh_jhdj<=13 or (sjjh_grade>5 and sjjh_grade<9) then%>{alert("很抱歉！此功能13级以下和官府不可用！")}return false;<%else%>document.af.towhoway.checked=!(document.af.towhoway.checked);<%end if%>document.af.sytemp.focus();' title='悄悄话儿悄悄说，１３级以下和官府不可用'>私聊</a>                       
            <input type='checkbox' name=as accesskey='a' checked onClick='parent.f1.scrollit();document.af.sytemp.focus();' value="ON">                                                                          
            <a href='#' onClick='document.af.as.checked=!(document.af.as.checked);                                                                          
           document.af.as.focus();parent.f1.scrollit();document.af.sytemp.focus();' title='是否滚屏显示'>滚屏</a>
        <input type='checkbox' name='gbbox' onClick="gBon()">                                                  
        <a href='#' onMouseOver="window.status='24小时广播';return true" onMouseOut="window.status='';return true" onClick="document.af.gbbox.click()" title="听24小时的广播，爽哦">广播</a>       
            <a href="f2/xp.asp?name=<%=nikename%>" target="f3" title="打扑克、玩麻将">[牌]</a> <a href='dushu.asp' onClick="javascript:s()" onMouseOver="window.status='小心点哦,不要倾家荡产啦^_^';return true" onMouseOut="window.status='';return true" target="f3")" title="小心点哦,不要倾家荡产啦^_^">[赌]</a>
            <%if sjjh_grade>=6 then%>                                                                          
             
            <a href="../dt/ask.asp" target="_blank" title="出题，让别人来抢答！">出题</a>                                                                           
            <a class=blue href="#" onClick="window.open('../dalie/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="放老虎，让别人来打！">虎</a>                                                                           
            <a class=blue href="#" onClick="window.open('../jiaofei/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="放土匪，让别人来打！">匪</a>
           
            <%else%>                                                                                                                                                  
            <a class=blue href="#" onClick="window.open('../dt/reply.asp','daliwu','scrollbars=no,resizable=no,width=490,height=600')" title="来答题吧，答对有银两赚喔！">答题</a>                                                                           
            <a class=blue href="#" onClick="window.open('../dalie/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要官府放老虎了，你可以打猎物得钱！">虎</a>                                                                          
            <a class=blue href="#" onClick="window.open('../jiaofei/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要官府放土匪了，你可以打匪得钱！">匪</a>                                                                           
  
            <%end if%>                                                                          
  
            <%if chatinfo(7)=0 then%>                                                                          
            <a href=lg.asp title="在练功的时间，你与别人是不能打架的！" target="d">保护</a>                                                                           
            <%end if%> 

            <%if sjjh_grade>=9 then%>     
            <input type='checkbox' name='slbox' accesskey='s' onClick="document.af.sytemp.focus();if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1}" value="ON">                                                                          
            <a href='#' onMouseOver="window.status='选中本项后，可以查看到聊天室中的私聊！。';return true" onMouseOut="window.status='';return true" onClick="document.af.slbox.checked=!(document.af.slbox.checked);if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1};document.af.sytemp.focus();" title="看一看大家在说一些什么悄悄话">查</a>                                                                          
            <%end if%>                                     
             
                                      
            <script language="JavaScript">function startnosay(){var nosay=<%=Application("sjjh_maxtimeout")*60%>;document.af.clock.value=nosay;}</script>                                                                          
            <script src="readonly/f2.js"></script>                                                                          
            <script>parent.fc();parent.m.location.href="f3.asp";</script>                                     
                                              
            <div id="dh" style="position:absolute; left:-80px; top:-80px; width:0px; height:0px; z-index:1">  
             
 
              <input type="button" value="<" name='gopp' onClick="Javascript:gop();" accesskey=","> 
              <input type="button" value=">" name='gonn' onClick="Javascript:gon();" accesskey="."> 
 
            </div> 
          </div> 
        </td> 
      </tr> 
    </table> 
    </form>
</body></html>