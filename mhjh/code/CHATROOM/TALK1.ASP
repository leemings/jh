<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=0
nosaytime=Application("yx8_mhjh_nosaytime")
username=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
chatroomname=Application("yx8_mhjh_systemname"&chatroomsn)
txxx=session("yx8_mhjh_tx")
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
function init(){var rndcolor=Math.floor(Math.random()*18);document.talkform.namecolor.value=document.talkform.namecolor.options[rndcolor].value;rndcolor=Math.floor(Math.random()*18);document.talkform.wordcolor.value=document.talkform.wordcolor.options[rndcolor].value;parent.msgfrm0.document.open();parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css3.css'></head><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><body text=000000 oncontextmenu=self.event.returnValue=false><font color=FF0000>【风起云涌-万马奔腾】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临<%=chatroomname%><font class=timsty><%=time()%></font><br>");parent.msgfrm1.document.open();parent.msgfrm1.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css3.css'></head><body text=000000 ><font color=FF0000>【风起云涌-万马奔腾】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临<%=chatroomname%><font class=timsty><%=time()%></font><br>");parent.getfrm.location.replace('getmsg.asp');parent.chgfrm.location.replace('selectchat.asp');nosay();}
function settalk(str1,str2){document.talkform.talkmsg.value=str1+' '+str2;document.talkform.talkmsg.focus();}
function addtalk(str1){document.talkform.talkmsg.value=document.talkform.talkmsg.value+str1+' ';}
</script>
<SCRIPT language=JavaScript>
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</SCRIPT>
<title></title>
</head>
<body oncontextmenu='self.event.returnValue=false'  onload='init();' leftmargin="5" topmargin="7" marginwidth="5" marginheight="5"  background="bg1.gif" style="CURSOR: url('../image/banana.ani')">
<form name=talkform action="sendmsg.asp" method=post target="sendfrm" onsubmit="javascript:return(checklingual());">
<input type=button name="lrclutch" value="功能开"  onclick="javascript:parent.lrclutch();" title="打开/关闭左边框" onmouseover="window.status='打开/关闭左边框。';return true" onMouseOut="window.status='';return true">
<select name=namecolor onchange="javascript:document.talkform.talkmsg.focus();">                                                  
<option style="background-color:000000;color:000000" value="000000">姓名</option>                                                  
<option style="background-color:000088;color:000088" value="000088">姓名</option>                                                  
<option style="background-color:000000;color:000000" value="000000">姓名</option>                                                  
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
<option style="background-color:8844ff;color:8844ff" value="8844ff">姓名</option>
</select><input type="hidden" name="talkmsgtmp" value="">                                                  
<input type=text name="talkmsg" value="" size=56 maxlength=100 class=norstyle>                                                  
<a href="#" onclick="javascript:goback();" title="取得前一句话，共二十句" onmouseover="window.status='取得前一句话，共二十句。';return true" onMouseOut="window.status='';return true"><font face=webdings>7</font></a>                                                  
<a href="#" onclick="javascript:goforw();" title="取得后一句话，共二十句" onmouseover="window.status='取得后一句话，共二十句。';return true" onMouseOut="window.status='';return true"><font face=webdings>8</font></a>                                                  
<input name="subm"type=submit value="发言" class=norstyle>                                                  
<input type="button" name=allownosaytime value="<%=nosaytime%>" readonly class=norstyle disabled><br>                                                  
                                                  
<input type=button name="tbclutch" value="关分屏" onclick="javascript:parent.tbclutch();" title="合屏/分屏切换" onmouseover="window.status='合屏/分屏切换。';return true" onMouseOut="window.status='';return true">                                                  
<select name=wordcolor onchange="javascript:document.talkform.talkmsg.focus();">                                                  
<option style="background-color:000000;color:000000" value="000000" selected>说话</option> 
<option style="background-color:000088;color:000088" value="000088">说话</option> 
<option style="background-color:000000;color:000000" value="000000">说话</option> 
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
</select> 
<select name=action class=norstyle onchange="javascript:document.talkform.talkmsg.value=this.value;document.talkform.talkmsg.focus();"> 
        <option value="" selected style="color:000000">动作</option> 
        <option value="/#蹑手蹑脚地溜到%%背后，掏出一把大锤子，嘿嘿......[img]16.GIF[/img]">我打</option>
        <option value="/#对着%%狠狠的一锤当头敲下[img]dz22.gif[/img]微笑道：“%%，你给我发呆去吧！”">锤子</option>
        <option value="/#从天上召来一道[img]dz33.GIF[/img]把%%化为一堆灰烬。">灰烬</option>
        <option value="/#结结巴巴地跟%%道歉“实在是对不住，我下次再也不敢了[img]11.GIF[/img]。”">道歉</option>
        <option value="/#和%%正邪恶地笑着[img]dz53.gif[/img]八成同时想到什么坏事头上。">坏事</option>
        <option value="/#对%%小声嘀咕[img]dz51.gif[/img]：“君子报仇，十年不晚，五年后再来找你。”">报仇</option>
        <option value="/#看着%%偷偷直乐，幸灾乐祸地想[img]dz40.gif[/img]活该。">活该</option>
        <option value="/#笑呵呵地对%%抱拳打揖“久仰阁下大名[img]10.GIF[/img]，如雷灌耳，今日相见，三生有幸！”">久仰</option>
        <option value="/#热情地地向%%叫了声[img]dz48.GIF[/img]“嗨，你好”">招呼 
        <option value="/#依依不舍地对%%说道：“各位，我有事要先走了，咱们后会有期[img]09.GIF[/img]！”">离开 
        <option value="/#对着%%[img]dz19.gif[/img]拍手叫好。">叫好 
        <option value="/#迎风傲立船首，望滚滚江水东逝，仰天长啸“古往今来多英杰，中游击橹几轮回？[img]12.GIF[/img]”">长啸 
        <option value="/#不由得十分佩服地叹道：“普天之下，茫茫苍穹[img]dz12.gif[/img]还有谁能与%%争锋啊？”">佩服 
        <option value="/#觉得天下英雄舍我其谁？[img]dz23.gif[/img]">英雄 
        <option value="/#非常同意%%的讲法！">同意 
        <option value="/#的[img]dz34.gif[/img]从在场每个人身上扫过，然后停在%%脸上问道：是不是阁下在造谣？如果是请不要再瞎折腾了。">造谣 
        <option value="/#对%%附耳道：好汉不吃眼前亏，还是赶紧溜吧。[img]dz39.gif[/img]">逃跑 
        <option value="/#对着%%大喝一声：小坏蛋，哪里逃！[img]dz27.gif[/img]">别跑 
        <option value="/#急得象热锅上的蚂蚁，不停地大叫：快点！快点！怎麽比[img]pic44.gif[/img]还慢？！">太慢 
        <option value="/#吓得躲在角落直发抖！别、别[img]dz21.gif[/img]别打我！">害怕 
        <option value="/#躲在一旁[img]dz49.gif[/img]偷偷地笑%%。">偷笑 
        <option value="/#看着%%，[img]dz46.gif[/img]很无奈地耸耸肩。">无奈 
        <option value="/#对着%%深深一躬，[img]dz37.gif[/img]说道：小人有眼不识泰山，大人您宰相肚里能撑船，可别跟小人一般见识啊！">包函 
        <option value="/#挤出人群[img]dz41.GIF[/img]大声道：“我，我支持%%！！！”">支持 
        <option value="/#对着%%大喝一声:好你个小坏蛋，想搞破坏！">害虫 
        <option value="/#很疑惑地[img]dz34.gif[/img]%%。">疑惑 
        <option value="/#一声大喊：“此山是我开，此树是我栽，若要从此过，留下买路财！” ">山贼 
        <option value="/#抱住%%的两腿[img]dz43.GIF[/img]一把鼻涕一把泪地说：“这位大侠，您行行好，帮帮俺吧！”">乞讨 
        <option value="/#哈哈[img]dz44.gif[/img]，对%%一拱手道：阁下过奖了，大家彼此彼此。">彼此 
        <option value="/#高举右拳[img]qt.gif[/img]咬牙切齿地高呼：“打倒%%！”">打倒 
        <option value="/#很看着%%[img]dz16.gif[/img]气死我了">生气 
        <option value="/#很不好意思地向%%致歉[img]dz37.gif[/img]">道歉 
        <option value="/#对%%大叫:救命哪！救救我啊！">救命 
        <option value="/#对%%小声嘀咕着[img]dz51.gif[/img]“君子报仇，十年不晚，五年后再来找你。”">报仇 
        <option value="/#看都不看%%一眼，哼了一声高高的把头扬了起来………不理你………[img]04.GIF[/img]">不理 
        <option value="/#拱手作揖。很有礼貌地对%%说:在下%%，请多指教！">拱手 
        <option value="/#痴痴地望着[img]dz35.gif[/img]%%的身影[img]dz52.gif[/img]不禁呆了……">痴痴 
        <option value="/#对%%一拱手道：原来小兄弟竟是志在天下在下五体投地。[img]06.GIF[/img]">感叹 
        <option value="/#将指节咔地一捏，伸手在%%的歪脑门上敲了一个双响脆毛栗子[img]24.GIF[/img]。">毛栗 
        <option value="/#不怀好意地绕着%%打转[img]dz36.gif[/img]不知在嘀咕些什么。">打转 
        <option value="/#大声地对%%喊道[img]dz45.gif[/img]“不许在这里做儿童不宜的动作！”">不许 
        <option value="/#对%%赞道：阁下[img]03.GIF[/img]运筹于什么什么之中，决算于什么什么之外，那个那个。。。世人难及，世人难及呀！">佩服 
        <option value="/#单腿下跪，脉脉含情的看着%%[img]dz15.gif[/img]虔诚说出了世界上最感人的三个字[img]01.GIF[/img]。。。">深情 
        <option value="/#狠狠掴了%%几个大耳光[img]dz02.gif[/img]打得他眼冒金星。">耳光 
        <option value="/#想到就能娶%%做老婆了[img]dz101.gif[/img]兴奋的几天睡不着觉。">想娶 
        <option value="/#开始梦想什么时候能够嫁给%%。。。[img]dz14.gif[/img]兴奋得满脸通红。">想嫁 
        <option value="/#对着%%大声惊呼：“天呐，瞧瞧你都干了些什么[img]18.GIF[/img]？！”">惊叫 
        <option value="/#对%%哼道咱们骑驴看帐本，走着瞧[img]08.GIF[/img]。">威胁 
        <option value="/#对%%：我原以为只有[img]dz32.gif[/img]才这么走路，没想到小兄弟也喜欢横行，失敬失敬！">横行 
        <option value="/#「唰」的一声抽出一柄宝剑握在手中[img]dz24.gif[/img]，顿时全身凉飕飕地只感寒气逼人。">拨剑 
        <option value="/#一看四下无人，猛地抱住%%[img]dz13.gif[/img]一阵狂啵儿~~！">狂吻 
        <option value="/#飞起一脚[img]dz04.gif[/img]把%%踢出了太阳系！">踢飞 
        <option value="/#恶狠狠地威胁道：%%，你再不老实！这[img]dz03.gif[/img]就是你的下场！">下场 
        <option value="/#眼前一花：我晕[img]dz17.gif[/img]！">我晕 
        <option value="/#苦练武功[img]dz06.gif[/img]心中默念：%%，你等着！我一定报仇雪恨！">报仇 
        <option value="/#恶狠狠地对%%[img]dz07.gif[/img]“踩死你！踩死你！”">踩人 
        <option value="/#挥舞着皮鞭冲着%%[img]dz08.gif[/img]叫你不听话！">鞭打 
        <option value="/#冲着%%大声叫到[img]dz31.gif[/img]不要呀！">大叫 
        <option value="/#兴奋的吹起[img]dz20.gif[/img]泰坦尼克！">演奏 
        <option value="/#冲着%%[img]dz09.gif[/img]气死你、气死你！">气人 
        <option value="/#眼睁睁地看着%%被[img]dz26.gif[/img]拉走！禁不住[img]dz30.gif[/img]">住院 
        <option value="/#觉得%%好恶心！我受不了了，我要~要~[img]dz10.gif[/img]">恶心 
        <option value="/#悄悄的挪动脚步，来到%%身旁，嘻嘻！[img]dz50.gif[/img]">吓人 
        <option value="/#向着%%大叫：“欢迎，欢迎，热烈欢迎！[IMG]dz63.gif[/IMG]”">欢迎 
        <option value="/#热情地向%%打招呼，并递上一杯香浓Coffee和一杯清香的茶，自己选哦。">饮料 
        <option value="/#拿出一张[IMG]dz01.gif[/IMG]摇头晃脑得意地对%%说，你可以陪我聊聊天……？">陪聊 
        <option value="/#对%%说，快些给我打电话[IMG]dz58.gif[/IMG]，否则……？">电话 
        <option value="/#对%%说，气死我了，气死我了，我一辈子也不再理你了！！！！！[IMG]dz16.gif[/IMG]">气死 
        <option value="/#用一根树枝捅了捅%%的鼻子：“醒醒呀，快醒醒。”">醒醒 
        <option value="/#看着镜中的%%，只见%%面色绯红，眼波流转，回眸一笑，百媚俱生令人不由得骨软筋酥，看着看着，##不由得痴了。">可爱 
        <option value="/#看着%%，心想：现在的女孩真厉害……佩服佩服[IMG]dz12.gif[/IMG">厉害 
        <option value="/#深情的对%%说，我们拉勾勾[IMG]dz11.gif[/IMG]不见不散哦~~~~~~！">拉勾 
        <option value="/#跪下来向%%磕头：我们结婚吧~!~~~~~[IMG]dz15.gif[/IMG]">飞吻 
        <option value="/#我晕~~~~~~~~~~[IMG]dz17.gif[/IMG]">晕倒 
        <option value="/#祝%%生日快乐[IMG]dz59.gif[/IMG]">生日 
        <option value="/#捧着%%的脸[IMG]dz13.gif[/IMG]说：嫁给我好吗？亲爱的，我等你10年了~~~~">鲜花 
        <option value="/#把头一甩，没有看大家一眼，[IMG]dz16.gif[/IMG]愤愤地说：气气气死我拉、、、、">气死 
        <option value="/#乐呵呵地对%%叫：我是流氓我怕谁，哈哈[IMG]33.gif[/IMG]">流氓</option>
        <option value="/#一张圆脸流下两行清泪哭着哀求%%，你就饶了俺吧，别打我[IMG]dz21.gif[/IMG]">求饶 
        <option value="/#见%%不怀好意，连忙抓起电话：114，119，120，122，96315 ... 心想：“警察叔叔怎么还不来呀？”">求助 
        <option value="/#看都不看%%一眼，哼了一声，高高的把头扬了起来……不理你……">不理 
        <option value="/#对着%%挥舞着静月江湖神拳[IMG]dz06.gif[/IMG]">拳打 
        <option value="/#把%%踩在地下。。[IMG]dz07.gif[/IMG]">踩扁 
        <option value="/#练成了“杀猪掌”，一掌把%%搞成这个造型了。[IMG]dz05.gif[/IMG]">棉掌 
        <option value="/#悄悄地把一包砒霜放在%%的碗里，躲在一旁，喜滋滋地看着她喝了下去">下毒
        <option value="/#向%%勇敢地跪了下来：“你愿意嫁给我吗”--真是勇气可嘉啊！音乐鼓励[IMG]8.gif[/IMG]">求婚 
        <option value="/#“啪啪啪”地扇了%%几个耳光，满脸通红说：强盗！[IMG]dz02.gif[/IMG]”">耳光 
        <option value="/#对着%%大声号啕：是大唐的兄弟姐妹救救我吧[IMG]dz03.gif[/IMG]”">救救 
        <option value="/#轻轻地摸着%%的脸，无限深情地说：“我好喜欢你哦......”">喜欢 
        <option value="/#抓起%%往烤箱里一放，大概过了五分钟，打开箱一看大吃一惊：怎么成了烤卤猪啦!!!!">卤猪 
        <option value="/#见%%快不行了，连忙召来[IMG]dz60.gif[/IMG]，把%%送进精神病院……">救护 
        <option value="/#轻轻的亲了%%一下，好深情呦……[IMG]dz57.GIF[/IMG]">亲热 
		<option value="/#举起好大好大的铁锤往%%头上用力一敲！***『锵！』*** %%表情呆滞！从他的眼神你彷佛看到。。。****5000000000Pt****">锤子 
		<option value="/#很生气地看着%%。">生气 
		<option value="/#很不好意思地向%%行礼致歉。">道歉 
		<option value="/#对%%大叫：救命哪！救救我啊！">救命 
		<option value="/#和%%正邪恶地笑着，八成同时想到什么坏事头上。">坏事 
		<option value="/#对%%小声嘀咕着：“君子报仇，十年不晚，五年后再来找你[img]07.GIF[/img]。”">报仇 
		<option value="/#看都不看%%一眼，哼了一声，高高的把头扬了起来………不理你………[img]05.GIF[/img]">不理 
		<option value="/#拱手作揖。很有礼貌地对%%说：在下##，请多指教[img]24.GIF[/img]！">拱手 
		<option value="/#痴痴地望着%%的身影，不禁呆了……">痴痴 
		<option value="/#对%%一拱手道：原来小兄弟竟是志在天下。在下五体投地。">感叹 
		<option value="/#将指节咔地一捏，伸手在%%的歪脑门上敲了一个双响脆毛栗子。">毛栗 
		<option value="/#不怀好意地绕着%%打转，嘴巴里不知在嘀咕些什么。">打转 
		<option value="/#大声地对%%喊道：“不许在这里做儿童不宜的动作[img]12.GIF[/img]！”">不许 
		<option value="/#对%%赞道：阁下运筹于什么什么之中，决算于什么什么之外，那个那个。。。世人难及，世人难及呀[img]10.GIF[/img]！">佩服 
		<option value="/#单腿下跪，脉脉含情的看着%%，虔诚说出了世界上最感人的三个字。。。">深情 
		<option value="/#狠狠掴了%%几个大耳光，打得他眼冒金星。">耳光 
		<option value="/#想到就能娶%%做老婆了，兴奋的几天睡不着觉。">想娶 
		<option value="/#开始梦想什么时候能够嫁给%%。。。兴奋得满脸通红。">想嫁 
		<option value="/#对着%%大声惊呼：“天呐，瞧瞧你都干了些什么？！”">惊叫 
		<option value="/#对%%哼道：咱们骑驴看帐本，走着瞧。">威胁 
		<option value="/#恶狠狠地威胁道：%%，你个小？？？！给我闭嘴！">闭嘴 
		<option value="/#对%%说：我原以为只有螃蟹才这么走路，没想到小兄弟也喜欢横行，失敬失敬！">横行 
		<option value="/#「唰」的一声抽出一柄宝剑握在手中，顿时全身凉飕飕地只感寒气逼人。">拨剑 
		</select> 
<select name=expression class=norstyle onchange="javascript:document.talkform.talkmsg.focus();">       
			<option value="" style="color:000000" selected>表情</option>       
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
<option value="//暂离 ">暂离</option> 
<option value="//回来 ">回来</option>      
<option value="//查看 ">查看</option>        
<option value="//放大 ">放大</option> 
<option value="//帮扶 ">帮扶</option>
<option value="//对赌 10000">对赌</option>
<option value="//传言 ">传言</option>
<option value="//传功 ">传功</option>
<option value="//抢亲 ">抢亲</option>    
<option value="//拜师 ">拜师</option>      
<option value="//收徒 ">收徒</option>      
<option value="//逐徒 ">逐徒</option>      
<option value="//离师 ">离师</option>      
<option value="//结拜 ">结拜</option>      
<option value="//接拜 ">接拜</option>  
<option value="//断义 ">断义</option>  
<option value="//送钱 ">送钱</option>      
<option value="//偷钱 ">偷钱</option>     
<option value="//偷窃 ">偷窃</option>      
<option value="//投掷 ">投掷</option>    
<option value="//吸星 ">吸星</option>       
<option value="//赠送 ">赠送</option> 
<option value="//求婚 " style=color:red>求婚</option>  
<option value="//逐出 ">逐出</option>  
<option value="//追杀 要杀的人的名字|天数" style=color:blue>追杀</option>  
<option value="//收尸 被杀人名字" style=color:blue>收尸</option> 
<option value="//结盟 门派结盟">结盟</option> 
<option value="//毁盟 ">毁盟</option> 
<option value="//拉人 ">拉人</option>
<option value="//传位 ">传位</option>
<option value="//册封 掌门册封弟子">册封</option>
<option value="//封位 君主册封臣子">封位</option>       
<option value="//逐出 ">逐出</option>
<option value="//基金 " style=color:red>基金</option>  
<option value="//加金 " style=color:red>加金</option>  
<option value="//奖励 这里输入数字" style=color:red>奖励</option>  
<option value="//发钱 " style=color:red>发钱</option>  
<option value="//掌令 这里输入号令内容" style=color:red>掌令</option>  
<option value="//君令 这里输入号令内容" style=color:red>君令</option>  
</select> 
<select name="filter" class=norstyle onchange="javascript:document.talkform.talkmsg.value=this.value;document.talkform.talkmsg.focus();">      
<option value="" style="color:000000" selected>会员</option>  
<option value="//冷饮 ">冷饮</option>
<option value="//心动 ">心动</option>
<option value="//怒吼 ">怒吼</option>
<option value="//心跳 ">心跳</option>
<option value="//道具 ">道具</option>
</select>  
<select name="filter" class=norstyle onchange="javascript:document.talkform.talkmsg.value=this.value;document.talkform.talkmsg.focus();">      
<option value="" style="color:000000" selected>官府</option>          
<%if session("yx8_mhjh_usercorp")="官府" then%>    
<option value="//公告 ">公告</option>  
<option value="//册封 ">册封</option> 
<option value="//警告 ">警告</option>                              
<option value="//封咒 ">封咒</option>          
<option value="//刑罚 ">刑罚</option>      
<option value="//罚款 ">罚款</option>      
<option value="//斩首 ">斩首</option>      
<%end if%>      
</select>
<input type=checkbox name="isprivacy" class=norstyle>                                                 
<a href="#" onmouseover="window.status='悄悄话悄悄儿的说，别怕人听见！';return true;" onmouseout="window.status='';return true;" onclick="javascript:document.talkform.isprivacy.checked=!(document.talkform.isprivacy.checked);document.talkform.talkmsg.focus();">私聊</a>                                                 
<input type=checkbox name="autoscroll" class=norstyle checked onclick="javascript:document.talkform.talkmsg.focus();parent.msgfrm0.chgautoscroll();">                                                 
<a href="#" onmouseover="window.status='自动/手动滚屏设置！';return true;" onmouseout="window.status='';return true;" onclick="javascript:document.talkform.autoscroll.checked=!(document.talkform.autoscroll.checked);document.talkform.talkmsg.focus();parent.msgfrm0.chgautoscroll();">滚屏</a>       
<input type=checkbox name="title" class=norstyle onclick="document.talkform.talkmsg.focus();">
<a href="javascript:;" onclick="document.talkform.title.checked=!(document.talkform.title.checked)" title='写标题字句'>标题</a> 
<input type=checkbox name="dwtx" checked class=norstyle title='是否使用系统特效功能(如：音乐等)！'>
<a href='#' onmouseover="window.status='是否使用系统特效功能(如：音乐等)！';return true;" onmouseout="window.status='';return true;" onClick='javascript:document.talkform.dwtx.checked=!(document.talkform.dwtx.checked);document.talkform.talkmsg.focus();' title='是否使用系统特效功能(如：音乐等)！'>声音</a>
<input type=checkbox name="tupian" checked class=norstyle title='是显示心情图片！'>
<a href='#' onmouseover="window.status='是否显示心情图片！选中的话屏幕显示会比较慢';return true;" onmouseout="window.status='';return true;" onClick='javascript:document.talkform.tupian.checked=!(document.talkform.tupian.checked);document.talkform.talkmsg.focus();' title='是显示心情图片！'>心情</a>
  <%if session("yx8_mhjh_usercorp")="官府" then%>
  <a href=ask.asp target="_blank">题</a> <a class=blue href="#" onClick="window.open('hand.asp','fangjs','scrollbars=yes,width=200,height=200,left=20,top=10')">僵</a> 
  <a class=blue href="#" onClick="window.open('dalie/zj.html','fangliwu','scrollbars=no,resizable=no,width=180,height=180')" title="点击后，请不要关闭小窗口！">猎</a> 
  <%else%>
  <a href=reply.asp target="_blank">答题</a> <a href="#" onclick="window.open('dalie/fl.asp','','width=444,height=278,location=no');">打猎</a> 
  <%end if%>
  <a href="lg1.asp" target="sendfrm">保</a>/<a href="lg2.asp" target="sendfrm">关</a> 
  <a href="setall.asp" target="resfrm">美屏</a> <a href="horse.asp" target="resfrm">赌马</a> 
  <a href="duzhuang.asp" target="resfrm">赌银</a> <a href="http://yx8.net/forum.asp" target="_blank">论坛</a>  <a class=red href="#" onClick="window.open('../js/factory.asp','fangliwu','scrollbars=no,resizable=no,width=400,height=300')" title="10级以下新手专用的免费补给店！"><b><font color=red>新人补给</font></b></a> 
  <input onmouseover="window.status='右下功能菜单控制';return true" title="右下功能菜单控制" onclick="javascript:parent.tbygn();" onmouseout="window.status='';return true" type="button" value="↓" name="tbygn">       
<input type=button name="youxia" value="→" onClick="javascript:parent.youxia();" title="用户名单控制" onMouseOver="window.status='用户名单控制';return true" onMouseOut="window.status='';return true" >   
</form> 
</body>
     
     
     
     
     
   
 
 





