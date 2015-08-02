<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nosaytime=Application("Ba_jxqy_nosaytime")
username=session("Ba_jxqy_username")
chatroomsn=session("Ba_jxqy_userchatroomsn")
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
%>
<head>
<link rel="stylesheet" href="css.css">
<script language=javascript>
if(window==window.top){top.location.href="chatroom.asp";}
var maxlingual=19;
var pos=0;
var lingualnum=0;
var i;
var lingualarr=new Array(maxlingual);
var nosaytime=<%=nosaytime%>;
for(i=0;i<maxlingual;i++){lingualarr[i]='';}
function checklingual(){if(document.talkform.talkmsg.value==''){alert('请选择语言或动作！');document.talkform.talkmsg.focus();return false;}if (lingualnum<maxlingual+1){lingualarr[lingualnum]=document.talkform.talkmsg.value;lingualnum++;}else{for (i=0;i<maxlingual;i++)lingualarr[i]=lingualarr[i+1];lingualarr[i]=document.talkform.talkmsg.value;}pos=lingualnum;{document.talkform.talkmsgtmp.value=document.talkform.talkmsg.value;document.talkform.talkmsg.value='';document.talkform.allownosaytime.value=nosaytime;document.talkform.subm.disabled=true;setTimeout('document.talkform.subm.disabled=false;',3000);document.talkform.talkmsg.focus();document.talkform.filter.value='';document.talkform.action.value='';return true;}return false;}
function goback(){if(pos>0)pos--;document.talkform.talkmsg.value=lingualarr[pos];document.talkform.talkmsg.focus();}
function goforw(){if(pos<lingualnum-1)pos++;document.talkform.talkmsg.value=lingualarr[pos];document.talkform.talkmsg.focus();}
function nosay(){document.talkform.allownosaytime.value=document.talkform.allownosaytime.value-1;setTimeout('nosay()',1000);if(document.talkform.allownosaytime.value==300){window.open('goback.asp','goback','Status=yes,scrollbars=no,resizable=no');}if(document.talkform.allownosaytime.value<0){top.location.href='chaterror.asp?id=001';}}
function init(){var rndcolor=Math.floor(Math.random()*18);document.talkform.namecolor.value=document.talkform.namecolor.options[rndcolor].value;rndcolor=Math.floor(Math.random()*18);document.talkform.wordcolor.value=document.talkform.wordcolor.options[rndcolor].value;parent.msgfrm0.document.open();parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css.css'></head><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><body text=000000 oncontextmenu=self.event.returnValue=false><font color=FF0000>【浏览器涮新】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临<%=chatroomname%><font class=timsty><%=time()%></font><br>");parent.msgfrm1.document.open();parent.msgfrm1.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css.css'></head><body text=000000 ><font color=FF0000>【浏览器涮新】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临<%=chatroomname%><font class=timsty><%=time()%></font><br>");parent.getfrm.location.replace('getmsg.asp');parent.chgfrm.location.replace('selectchat.asp');nosay();}
function settalk(str1,str2){document.talkform.talkmsg.value=str1+' '+str2;document.talkform.talkmsg.focus();}
function addtalk(str1){document.talkform.talkmsg.value=document.talkform.talkmsg.value+str1+' ';}
</script>
</head>
<body oncontextmenu='self.event.returnValue=false' background="../images/bg.gif" bgcolor="#FFE4CA" onload='init();' leftmargin="5" topmargin="7" marginwidth="5" marginheight="5">
<form name=talkform action="sendmsg.asp" method=post target="sendfrm" onsubmit="javascript:return(checklingual());">
  <input type=button name="lrclutch" value="开功能菜单"  onclick="javascript:parent.lrclutch();" title="打开/关闭左边框" onmouseover="window.status='打开/关闭左边框。';return true" onMouseOut="window.status='';return true">
		<select name=namecolor onchange="javascript:document.talkform.talkmsg.focus();">
			<option style="background-color:000000;color:000000" value="000000">姓名</option>
			<option style="background-color:000088;color:000088" value="000088">姓名</option>
			<option style="background-color:0000ff;color:0000ff" value="0000ff">姓名</option>
			<option style="background-color:008800;color:008800" value="008800">姓名</option>
			<option style="background-color:008888;color:008888" value="008888">姓名</option>
			<option style="background-color:0088ff;color:0088ff" value="0088ff">姓名</option>
			<option style="background-color:004400;color:004400" value="004400">姓名</option>
			<option style="background-color:004488;color:004488" value="004488">姓名</option>
			<option style="background-color:0044ff;color:0044ff" value="0044ff">姓名</option>
			<option style="background-color:880000;color:880000" value="880000">姓名</option>
			<option style="background-color:880088;color:880088" value="880088">姓名</option>
			<option style="background-color:8800ff;color:8800ff" value="8800ff">姓名</option>
			<option style="background-color:888800;color:888800" value="888800">姓名</option>
			<option style="background-color:888888;color:888888" value="888888">姓名</option>
			<option style="background-color:8888ff;color:8888ff" value="8888ff">姓名</option>
			<option style="background-color:884400;color:884400" value="884400">姓名</option>
			<option style="background-color:884488;color:884488" value="884488">姓名</option>
			<option style="background-color:8844ff;color:8844ff" value="8844ff">姓名</option>		</select><input type="hidden" name="talkmsgtmp" value="">
			<input type=text name="talkmsg" value="" size=56 maxlength=100 class=norstyle>
		<a href="#" onclick="javascript:goback();" title="取得前一句话，共二十句" onmouseover="window.status='取得前一句话，共二十句。';return true" onMouseOut="window.status='';return true"><font face=webdings>7</font></a>
		<a href="#" onclick="javascript:goforw();" title="取得后一句话，共二十句" onmouseover="window.status='取得后一句话，共二十句。';return true" onMouseOut="window.status='';return true"><font face=webdings>8</font></a>
		<input name="subm"type=submit value="发言" class=norstyle>
		<input type="button" name=allownosaytime value="<%=nosaytime%>" readonly class=norstyle disabled><br>
		
  <input type=button name="tbclutch" value="关分屏窗口" onclick="javascript:parent.tbclutch();" title="合屏/分屏切换" onmouseover="window.status='合屏/分屏切换。';return true" onMouseOut="window.status='';return true">
		<select name=wordcolor onchange="javascript:document.talkform.talkmsg.focus();">
    <option style="background-color:000000;color:000000" value="000000" selected>说话</option>
    <option style="background-color:000088;color:000088" value="000088">说话</option>
    <option style="background-color:0000ff;color:0000ff" value="0000ff">说话</option>
    <option style="background-color:008800;color:008800" value="008800">说话</option>
    <option style="background-color:008888;color:008888" value="008888">说话</option>
    <option style="background-color:0088ff;color:0088ff" value="0088ff">说话</option>
    <option style="background-color:004400;color:004400" value="004400">说话</option>
    <option style="background-color:004488;color:004488" value="004488">说话</option>
    <option style="background-color:0044ff;color:0044ff" value="0044ff">说话</option>
    <option style="background-color:880000;color:880000" value="880000">说话</option>
    <option style="background-color:880088;color:880088" value="880088">说话</option>
    <option style="background-color:8800ff;color:8800ff" value="8800ff">说话</option>
    <option style="background-color:888800;color:888800" value="888800">说话</option>
    <option style="background-color:888888;color:888888" value="888888">说话</option>
    <option style="background-color:8888ff;color:8888ff" value="8888ff">说话</option>
    <option style="background-color:884400;color:884400" value="884400">说话</option>
    <option style="background-color:884488;color:884488" value="884488">说话</option>
    <option style="background-color:8844ff;color:8844ff" value="8844ff">说话</option>
  </select><select name=sendto class=norstyle onclick="javascript:parent.chgsendto('大家');">
			<option value='大家'>大家</option>
		</select><select name=action class=norstyle onchange="javascript:document.talkform.talkmsg.value=this.value;document.talkform.talkmsg.focus();">
		<option value="/#蹑手蹑脚地溜到%%背后，掏出一把大锤子，嘿嘿......">偷偷</option>
        <option value="/#对着%%狠狠的一锤当头敲下，微笑道：“%%，你给我发呆去吧！”">锤子</option>
        <option value="/#从天上召来一道闪电把%%化为一堆灰烬。">灰烬</option>
        <option value="/#结结巴巴地跟%%道歉：“实在是对不住，我下次再也不敢了。”">道歉</option>
        <option value="/#和%%正邪恶地笑着，八成同时想到什么坏事头上。">坏事</option>
        <option value="/#对%%小声嘀咕着：“君子报仇，十年不晚，五年后再来找你。”">报仇</option>
        <option value="/#看着%%偷偷直乐，幸灾乐祸地想：活该。">活该</option>
        <option value="/#笑呵呵地对%%抱拳打揖：“久仰阁下大名，如雷灌耳，今日相见，三生有幸！”">久仰</option>
        <option value="/#热情地地向%%叫了声“嗨，你好”">招呼
		<option value="/#依依不舍地对%%说道：“各位，我有事要先走了，咱们后会有期！”">离开
		<option value="/#对着%%拍手叫好。">叫好
		<option value="/#迎风傲立船首，望滚滚江水东逝，仰天长啸：“古往今来多英杰，中游击橹几轮回？”">长啸
		<option value="/#不由得十分佩服地叹道：“普天之下，茫茫苍穹，还有谁能与%%争锋啊？”">佩服
		<option value="/#觉得天下英雄舍我其谁？">英雄
		<option value="/#非常同意%%的讲法！">同意
		<option value="/#的眼光从在场每个人身上扫过，然后停在%%脸上，问道：是不是阁下在造谣？如果是请不要再瞎折腾了。">造谣
		<option value="/#对%%附耳道：好汉不吃眼前亏，还是赶紧溜吧。">逃跑
		<option value="/#对着%%大喝一声：小王八蛋，哪里逃！">别跑
		<option value="/#急得象热锅上的蚂蚁，不停地大叫：快点！快点！怎麽比猪还慢？！">太慢
		<option value="/#吓得躲在角落直发抖！">害怕
		<option value="/#躲在一旁偷偷地笑%%。">偷笑
		<option value="/#看着%%，很无奈地耸耸肩。">无奈
		<option value="" selected style="color:0000FF">动作</option>
		<option value="/#对着%%深深一躬，说道：小人有眼不识泰山，大人您宰相肚里能撑船，可别跟小人一般见识啊！">包函
		<option value="/#挤出人群，大声道：“我，我支持%%！！！”">支持
		<option value="/#对着%%大喝一声：好你个小王八蛋，想搞破坏！">害虫
		<option value="/#很疑惑地看着%%。">疑惑
		<option value="/#一声大喊：“此山是我开，此树是我栽，若要从此过，留下买路财！” ">山贼
		<option value="/#抱住%%的两腿，一把鼻涕一把泪地说：“这位大侠，您行行好，帮帮俺吧！”">乞讨
		<option value="/#哈哈大笑，对%%一拱手道：阁下过奖了，大家彼此彼此。">彼此
		<option value="/#高举右拳，咬牙切齿地高呼：“打倒%%！”">打倒
		<option value="/#举起好大好大的铁锤往%%头上用力一敲！***『锵！』*** %%表情呆滞！从他的眼神你彷佛看到。。。****5000000000Pt****">锤子
		<option value="/#很生气地看着%%。">生气
		<option value="/#很不好意思地向%%行礼致歉。">道歉
		<option value="/#对%%大叫：救命哪！救救我啊！">救命
		<option value="/#和%%正邪恶地笑着，八成同时想到什么坏事头上。">坏事
		<option value="/#对%%小声嘀咕着：“君子报仇，十年不晚，五年后再来找你。”">报仇
		<option value="/#看都不看%%一眼，哼了一声，高高的把头扬了起来………不理你………">不理
		<option value="/#拱手作揖。很有礼貌地对%%说：在下##，请多指教！">拱手
		<option value="/#痴痴地望着%%的身影，不禁呆了……">痴痴
		<option value="/#对%%一拱手道：原来小兄弟竟是志在天下。在下五体投地。">感叹
		<option value="/#将指节咔地一捏，伸手在%%的歪脑门上敲了一个双响脆毛栗子。">毛栗
		<option value="/#不怀好意地绕着%%打转，嘴巴里不知在嘀咕些什么。">打转
		<option value="/#大声地对%%喊道：“不许在这里做儿童不宜的动作！”">不许
		<option value="/#对%%赞道：阁下运筹于什么什么之中，决算于什么什么之外，那个那个。。。世人难及，世人难及呀！">佩服
		<option value="/#单腿下跪，脉脉含情的看着%%，虔诚说出了世界上最感人的三个字。。。">深情
		<option value="/#狠狠掴了%%几个大耳光，打得他眼冒金星。">耳光
		<option value="/#想到就能娶%%做老婆了，兴奋的几天睡不着觉。">想娶
		<option value="/#开始梦想什么时候能够嫁给%%。。。兴奋得满脸通红。">想嫁
		<option value="/#对着%%大声惊呼：“天呐，瞧瞧你都干了些什么？！”">惊叫
		<option value="/#对%%哼道：咱们骑驴看帐本，走着瞧。">威胁
		<option value="/#恶狠狠地威胁道：%%，你个小？？？！给我闭嘴！">闭嘴
		<option value="/#对%%说：我原以为只有螃蟹才这么走路，没想到小兄弟也喜欢横行，失敬失敬！">横行
		<option value="/#「唰」的一声抽出一柄宝剑握在手中，顿时全身凉飕飕地只感寒气逼人。">拨剑
		</select><select name=expression class=norstyle onchange="javascript:document.talkform.talkmsg.focus();">
			<option value="" style="color:0000FF" selected>表情</option>
			<option value="微微笑着">微笑</option>
			<option value="哈哈大笑着">大笑</option>
			<option value="嘿嘿地奸笑着">奸笑</option>
			<option value="黯然神伤地">神伤</option>
			<option value="大声地">大声</option>
			<option value="义正严辞地">正义</option>
			<option value="有气无力地">无力</option>
			<option value="拳打脚踢地">拳脚</option>
			<option value="幸福地">幸福</option>
			<option value="十二万分陶醉地">陶醉</option>
			<option value="嘻皮笑脸地">涎皮</option>
			<option value="深情地">深情</option>
		</select><select name="filter" class=norstyle onchange="javascript:document.talkform.talkmsg.value=this.value;document.talkform.talkmsg.focus();">
		<option value="" style="color:0000FF" selected>命令</option>
		<option value="//册封 ">册封</option>
		<option value="//查看 ">查看</option>
		<option value="//传功 ">传功</option>
		<option value="//传言 ">传言</option>
		<option value="//道具 ">道具</option>
		<option value="//发钱 ">发钱</option>
		<option value="//公告 ">公告</option>
		<option value="//攻击 ">攻击</option>
		<option value="//警告 ">警告</option>
		<option value="//送钱 ">送钱</option>
		<option value="//偷窃 ">偷窃</option>
		<option value="//投掷 ">投掷</option>
		<option value="//吸星 ">吸星</option>
		<option value="//赠送 ">赠送</option>
		<option value="//逐出 ">逐出</option>
		</select><input type=checkbox name="isprivacy" class=norstyle>
  <a href="#" onmouseover="window.status='悄悄话悄悄儿的说，别怕人听见！';return true;" onmouseout="window.status='';return true;" onclick="javascript:document.talkform.isprivacy.checked=!(document.talkform.isprivacy.checked);document.talkform.talkmsg.focus();">私聊</a> 
  <input type=checkbox name="autoscroll" class=norstyle checked onclick="javascript:document.talkform.talkmsg.focus();parent.msgfrm0.chgautoscroll();">
  <a href="#" onmouseover="window.status='自动/手动滚屏设置！';return true;" onmouseout="window.status='';return true;" onclick="javascript:document.talkform.autoscroll.checked=!(document.talkform.autoscroll.checked);document.talkform.talkmsg.focus();parent.msgfrm0.chgautoscroll();">滚屏</a> 
  <font color="#FF0000">[<b><a href=# onClick="javascript:window.open('../readme.htm','readme',' width=550,height=500,left=200,top=100,status=no,toolbars=yes,menubars=yes,scrollbars=yes,resize=no')" title=新手入门>帮助</a></b></font>] 
  <b>
 <%if session("Ba_jxqy_usercorp")="官府" then%> 
<a href=ask.asp target="_blank">[出题]</a>
<%else%>
<a href=reply.asp target="_blank">[答题]</a>
<%end if%>
</b>
</form>
</body> 