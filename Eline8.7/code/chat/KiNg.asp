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
<html><head><META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; CURSOR: url('yinsu.ani');font-size: 9pt; color:'ffff00';
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff

	}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:D7EBff}
option	{font-size: 9pt; background-color:D7EBff}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'ffff77'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff4d4d;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: '291C03'; border: 1 solid '291C03'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-CENTER: 1px; PADDING-CENTER: 1px; 
	BORDER-CENTER-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-CENTER-WIDTH: 1px; 
	BORDER-CENTER-COLOR:'EBA82C'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: '333333'; 
	border: 333333px solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'F3CB81'; color: '83590C'; 
	border: 1px solid AB7510;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'EBA82C'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'EBA82C'}
</style>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}
function rc(list){if(list!="0"){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}}
</script>

</head>
<body bgcolor="#006699" topmargin="3" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<div align=left class=yy4>

</div>
<div align="center">
  <center>
    <table border="1" width="95%" bgcolor="#006699" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="2">
      <tr>
        <td width="100%" align="center"><font color="#CCCCFF"><strong>常用功能</strong></font></td>
  </tr>
  <tr>
    <td width="100%" align="center"><select name='addsign' onChange="rc(this.value);" style='font-size:12px'>
<option value="0" selected>动作命令</option>
<option value="//蹑手蹑脚地溜到%%背后，掏出一把大锤子，嘿嘿......">偷偷</option>
<option value="//对着%%狠狠的一锤当头敲下[img]dz22.gif[/img]微笑道：“%%，你给我发呆去吧！”">锤子</option>
<option value="//从天上召来一道[img]dz33.GIF[/img]把%%化为一堆灰烬。">灰烬</option>
<option value="//结结巴巴地跟%%道歉“实在是对不住，我下次再也不敢了。”">道歉</option>
<option value="//和%%正邪恶地笑着[img]dz53.gif[/img]八成同时想到什么坏事头上。">坏事</option>
<option value="//对%%小声嘀咕[img]dz51.gif[/img]：“君子报仇，十年不晚，五年后再来找你。”">报仇</option>
<option value="//看着%%偷偷直乐，幸灾乐祸地想[img]dz40.gif[/img]活该。">活该</option>
<option value="//笑呵呵地对%%抱拳打揖“久仰阁下大名，如雷灌耳，今日相见，三生有幸！”">久仰</option>
<option value="//热情地地向%%叫了声[img]dz48.GIF[/img]“嗨，你好”">招呼
<option value="//依依不舍地对%%说道：“各位，我有事要先走了，咱们后会有期！”">离开
<option value="//对着%%[img]dz19.gif[/img]拍手叫好。">叫好
<option value="//迎风傲立船首，望滚滚江水东逝，仰天长啸“古往今来多英杰，中游击橹几轮回？”">长啸
<option value="//不由得十分佩服地叹道：“普天之下，茫茫苍穹[img]dz12.gif[/img]还有谁能与%%争锋啊？”">佩服
<option value="//觉得天下英雄舍我其谁？[img]dz23.gif[/img]">英雄
<option value="//非常同意%%的讲法！">同意
<option value="//的[img]dz34.gif[/img]从在场每个人身上扫过，然后停在%%脸上问道：是不是阁下在造谣？如果是请不要再瞎折腾了。">造谣
<option value="//对%%附耳道：好汉不吃眼前亏，还是赶紧溜吧。[img]dz39.gif[/img]">逃跑
<option value="//对着%%大喝一声：小坏蛋，哪里逃！[img]dz27.gif[/img]">别跑
<option value="//急得像热锅上的蚂蚁，不停地大叫：快点！快点！怎么比[img]pic44.gif[/img]还慢？！">太慢
<option value="//吓得躲在角落直发抖！别、别[img]dz21.gif[/img]别打我！">害怕
<option value="//躲在一旁[img]dz49.gif[/img]偷偷地笑%%。">偷笑
<option value="//肚子好通[img]dz620.gif[/img]等等我大一下便%%。">大便
<option value="//哇忍不住了[img]dz621.gif[/img]终于拉出来了%%。">拉屎
<option value="//大家好阿[img]dz622.gif[/img]你好阿%%。">HI
<option value="//不行了快尿出来了[img]dz623.gif[/img]等等在聊好吗%%。">小便
<option value="//猪猪屁股扭来扭去[img]dz624.gif[/img]好笑吗%%。">猪猪
<option value="//扭死了哈[img]dz625.gif[/img]%%:说哇这模快就死了。">死猪
<option value="//会跳舞的猪喔[img]dz626.gif[/img]酷不酷阿%%。">舞猪
<option value="//猪生气了[img]dz627.gif[/img]不要笑他%%。">气猪
<option value="//我的魅力还不错吧[img]dz628.gif[/img]%%。">钩心







<option value="//看着%%，[img]dz46.gif[/img]很无奈地耸耸肩。">无奈
<option value="//对着%%深深一躬，[img]dz37.gif[/img]说道：小人有眼不识泰山，大人您宰相肚里能撑船，可别跟小人一般见识啊！">包涵
<option value="//挤出人群[img]dz41.GIF[/img]大声道：“我，我支持%%！！！”">支持
<option value="//对着%%大喝一声:好你个小坏蛋，想搞破坏！">害虫
<option value="//很疑惑地[img]dz34.gif[/img]%%。">疑惑
<option value="//一声大喊：“此山是我开，此树是我栽，若要从此过，留下买路财！” ">山贼
<option value="//抱住%%的两腿[img]dz43.GIF[/img]一把鼻涕一把泪地说：“这位大侠，您行行好，帮帮俺吧！”">乞讨
<option value="//哈哈[img]dz44.gif[/img]，对%%一拱手道：阁下过奖了，大家彼此彼此。">彼此
<option value="//高举右拳[img]qt.gif[/img]咬牙切齿地高呼：“打倒%%！”">打倒
<option value="//很看着%%[img]dz16.gif[/img]气死我了">生气
<option value="//很不好意思地向%%致歉[img]dz37.gif[/img]">道歉
<option value="//对%%大叫:救命哪！救救我啊！">救命
<option value="//对%%小声嘀咕着[img]dz51.gif[/img]“君子报仇，十年不晚，五年后再来找你。”">报仇
<option value="//看都不看%%一眼，哼了一声高高的把头扬了起来………不理你………">不理
<option value="//拱手作揖。很有礼貌地对%%说:在下%%，请多指教！">拱手
<option value="//痴痴地望着[img]dz35.gif[/img]%%的身影[img]dz52.gif[/img]不禁呆了……">痴痴
<option value="//对%%一拱手道：原来小兄弟竟是志在天下在下五体投地。">感叹
<option value="//将指节卡地一捏，伸手在%%的歪脑门上敲了一个双响脆毛栗子。">毛栗
<option value="//不怀好意地绕着%%打转[img]dz36.gif[/img]不知在嘀咕些什么。">打转
<option value="//大声地对%%喊道[img]dz45.gif[/img]“不许在这里做儿童不宜的动作！”">不许
<option value="//对%%赞道：阁下运筹于什么什么之中，决算于什么什么之外，那个那个。。。世人难及，世人难及呀！">佩服
<option value="//单腿下跪，脉脉含情的看着%%[img]dz15.gif[/img]虔诚说出了世界上最感人的三个字。。。">深情
<option value="//狠狠掴了%%几个大耳光[img]dz02.gif[/img]打得他眼冒金星。">耳光
<option value="//想到就能娶%%做老婆了[img]dz101.gif[/img]兴奋的几天睡不着觉。">想娶
<option value="//开始梦想什么时候能够嫁给%%。。。[img]dz14.gif[/img]兴奋得满脸通红。">想嫁
<option value="//对着%%大声惊呼：“天呐，瞧瞧你都干了些什么？！”">惊叫
<option value="//对%%哼道咱们骑驴看帐本，走着瞧。">威胁
<option value="//对%%：我原以为只有[img]dz32.gif[/img]才这么走路，没想到小兄弟也喜欢横行，失敬失敬！">横行
<option value="//“唰”的一声抽出一柄宝剑握在手中[img]dz24.gif[/img]，顿时全身凉飕飕地只感寒气逼人。">拨剑
<option value="//一看四下无人，猛地抱住%%[img]dz13.gif[/img]一阵狂啵儿~~！">狂吻
<option value="//飞起一脚[img]dz04.gif[/img]把%%踢出了太阳系！">踢飞
<option value="//恶狠狠地威胁道：%%，你再不老实！这[img]dz03.gif[/img]就是你的下场！">下场
<option value="//眼前一花：我晕[img]dz17.gif[/img]！">我晕
<option value="//苦练武功[img]dz06.gif[/img]心中默念：%%，你等着！我一定报仇雪恨！">报仇
<option value="//恶狠狠地对%%[img]dz07.gif[/img]“踩死你！踩死你！”">踩人
<option value="//挥舞着皮鞭冲着%%[img]dz08.gif[/img]叫你不听话！">鞭打
<option value="//冲着%%大声叫到[img]dz31.gif[/img]不要呀！">大叫
<option value="//兴奋的吹起[img]dz20.gif[/img]泰坦尼克！">演奏
<option value="//冲着%%[img]dz09.gif[/img]气死你、气死你！">气人
<option value="//眼睁睁地看着%%被[img]dz26.gif[/img]拉走！禁不住[img]dz30.gif[/img]">住院
<option value="//觉得%%好恶心！我受不了了，我要~要~[img]dz10.gif[/img]">恶心
<option value="//悄悄的挪动脚步，来到%%身旁，嘻嘻！[img]dz50.gif[/img]">吓人
<option value="//对%%痛苦地说道：“结婚才三天就把老公打成这样，我要离婚，这日子没法过了。”">离婚
 <option value="//呆呆地看着镜子里的自己，傻呼呼的要和镜子里的人握手，谁知道镜子里的人竟然变成了%%的模样，哪里有这样的事情啊？">自怜
 <option value="//在江湖猪圈之上，负手而立，凝望聊天室里乱跑的猪群、羊群和变压器群。只觉得天下拱猪英雄舍我与%%其谁啊~~">拱猪
 <option value="//对%%使用了一招中华楼的满汉全席，自己的内力下降965987458,道德下降5000，得到965987的经验值，杀伤对方生命96598745685，%%惨叫一声，口吐鲜血，慢慢的倒了下去......，在场的众位大侠都无不惊骇中华楼有这样狠毒的武功着数。同时也觉得##过于心狠手黑了">杀手
<option value="//弯腰拱手，淡淡一笑“久闻风云城人杰地灵，小弟初来乍到，还望诸位提点提点”。">拱手
<option value="//留恋地说道：“世上没有不散的宴席，我先走一步了，大家保重。”">先走 
<option value="//仰起脸，伸长鼻子对着空气猛嗅了几下， 对着%%大声赞道：好香，好香!">好香
<option value="//最近一直觉得自己头晕目眩,恶心想吐,不知是不是.....">头晕
<option value="//趁%%不注意时在%%的脸上偷偷亲了一下，没想到%%叫到：“你口好臭！多久没嗽口了？">偷吻
<option value="//紧紧靠在%%的怀里，口中呢喃，鼻翼歙动，面赛桃花">激动
<option value="//大叫一声：“风紧，扯呼”,撒开腿就跑，片刻就消失在一片茫茫白雾中"">扯呼
<option value="//一把眼泪一把鼻涕地喊：冤枉啊，%%青天大老爷!!"">喊冤
<option value="//飘啊飘的一段情，有雨也有风……蓦然回首你仍在，浪漫红尘中……我走了……"">道别
<option value="//象暴风雨一样扑了过来，嘴里说着，%%，不能没有你啊，……可惜，脚下一滑，一嘴吻在对方的鞋帮子上。"">狂吻
<option value="//象朵含羞草一样躲了起来，又突然跑回来，轻轻地、快快地吻了%%一下，转身就跑了。"">羞吻
<option value="//浑身酒气挣扎着踉跄走了两步，突然不留神，一个跟头直直插进了路边的水沟里。"">喝醉
<option value="//像只小猫似的依偎在%%的怀里!!"">小猫
<option value="//突然拔出一把手枪，指着自己的头，泪眼婆娑的看着大家。"">自杀
<option value="//紧紧咬住自己的嘴唇，眼睛象水滴一样生动地注视着对方，用可以摧毁一个国家，杀死一个国王，揉碎一颗心的声音说道，云想衣裳，花想容，我又怎么会不想呢。"">冲动
<option value="//强忍住烦躁和一刹那的烈火焚心，从房间里冲到花园，果断地纵身跳进冰凉的游泳池中，大家惊讶的发现，不一会儿，诺大一个游泳池的水居然全沸腾了！"">冲动
<option value="//突然大喊一声，这是我的第三千八百次初恋，求求你，%%，接受我的这一小片痴情吧！"">求爱
<option value="//象断线的风筝一样徒劳地挣扎了两下，然后无可救药的落进了%%的温柔陷阱……"">陷阱
<option value="//对自己的长像极度不满意, 决定向双亲提出抗议.">自卑
<option value="//丢几罐冰凉的啤酒给%%，然后说:“干啦 !”">敬酒
<option value="//用脸颊轻轻地磨擦着%%的粉脸, 悄声说道: 我好喜欢你哦....">脸颊
<option value="//用手托起%%的下巴深情的长吻 [IMG]dz57.GIF[/IMG] ~@~"">长吻
<option value="//自言自语道:“我,先天下人之忧而忧, 后天下人之乐而乐...这个这个好像不太妥”">自语
<option value="//心中默默念道：“由爱故生忧，由爱故生怖；若离于爱者，无忧亦无怖。”">默念
<option value="//歪着头憧憬地向往：“如果上班又有工资，又没有老板管，那不知道该有多好。”">憧憬<option value="//俏脸生春，妙目含情，只看得大家心慌意乱。">含情
<option value="//绞着手中的手绢，害羞地把头扭开“臭男人，看什么看？”">害羞
<option value="//闭上双目，柔情如流水般的唱道：谢谢你给我的爱，今生今世我不忘怀；谢谢你给我的温柔，伴我度过那个年代……">唱歌
<option value="//自言自语道：“今儿个不知该到谁家蹭饭去了......”">蹭饭
<option value="//说道：“乖乖不得了，在下的老板来了，没功夫聊了。回见！”">老板
<option value="//眉开眼笑地想：人都说打是疼，骂是爱，人家会不会，这个这个......">臭美
<option value="//嘴角一撇，狞笑道：“我是流氓，我怕谁？！”">流氓
<option value="//看着镜中日益消瘦的脸庞自怨自艾，“唉，##，你好命苦！”">命苦
<option value="//一脸的迷茫，不知道周围发生了什么事。">迷茫
<option value="//使劲地拍巴掌，啪啪啪啪啪.....拍得手都红起来了！[IMG]dz19.gif[/IMG]">巴掌
<option value="//往洞里一躺，嘟咙到：“熊熊我要冬眠了，不要打搅我！”">冬眠
<option value="//抱拳团团一拜，讨好地说道：“敝人对各位的景仰之情，有如涛涛江水连绵不绝。”">讨好
<option value="//的脸都红了，恨没有地洞，好钻进去躲起来。">脸红
<option value="//把胸脯拍得辟啪响：“武林中拳头大的说话，有种的上来比划比划！”">吹牛
<option value="//朗声说道：“今番良晤，豪兴不浅，他日江湖相逢，再当杯酒言欢。咱们就此别过。”说着袍袖一拂，飘然离去。">再见
<option value="//躲在墙脚瑟瑟发抖，嘴里喃喃自语“我不好，我检讨，我不对，我有罪”。">检讨
<option value="//快要饿死啦，走啦吃饭啦，“人是铁，饭是钢！”，再见了！">吃钣
<option value="//拔剑长吟道:“十年磨一剑，霜寒未曾试。今日把君问，谁有不平事？”">吟诗
<option value="//狂笑三声，从身后抽出一柄斧头，大喊着，天下没有公理了，让我来伸张正义。">砍人<option value="//刷的抽出一把闪闪发光的戒刀，一刀将自己的大好头颅切了下来，嘴里还不停的嘀咕，和我比酷，谁能比我酷？">比酷
<option value="//突然觉得胃一阵抽筋, 哇啦哇啦地吐了满地。">呕吐
<option value="//奏罢一曲６６江湖，一抚手中古松琴，长叹一声“知音难觅，罢罢罢，还进什么万梅！”手中琴弦寸断。">赠琴
<option value="//手握长刀，慨叹“天下之大，竟无一个对手，寂寞啊！！”">无敌
<option value="//打了自己一耳光，不好意思地说道：“真是对不起，在下一时失言，大家多多包涵。”">道歉
<option>----</option> 
<option value="//向着%%大叫：“欢迎，欢迎，热烈欢迎！[IMG]dz63.gif[/IMG]”">欢迎
<option value="//热情地向%%打招呼，并递上一杯香浓Coffee和一杯清香的茶，自己选哦。">饮料
<option value="//拿出一张[IMG]dz01.gif[/IMG]摇头晃脑得意地对%%说，你可以陪我聊聊天……？">陪聊<option value="//对%%说，快些给我打电话[IMG]dz58.gif[/IMG]，否则……？">电话
<option value="//对%%说，气死我了，气死我了，我一辈子也不再理你了！！！！！[IMG]dz16.gif[/IMG]">气死
<option value="//用一根树枝捅了捅%%的鼻子：“醒醒呀，快醒醒。”">醒醒
<option value="//看着镜中的%%，只见%%面色绯红，眼波流转，回眸一笑，百媚俱生令人不由得骨软筋酥，看着看着，##不由得痴了。">可爱
<option value="//看着%%，心想：现在的女孩真厉害……佩服佩服[IMG]dz12.gif[/IMG">厉害
<option value="//深情的对%%说，我们拉勾勾[IMG]dz11.gif[/IMG]不见不散哦~~~~~~！">拉勾
<option value="//跪下来向%%磕头：我们结婚吧~!~~~~~[IMG]dz15.gif[/IMG]">飞吻
<option value="//我晕~~~~~~~~~~[IMG]dz17.gif[/IMG]">晕倒
<option value="//祝%%生日快乐[IMG]dz59.gif[/IMG]">生日
<option value="//捧着%%的脸[IMG]dz13.gif[/IMG]说：嫁给我好吗？亲爱的，我等你10年了~~~~">鲜花
<option value="//把头一甩，没有看大家一眼，[IMG]dz16.gif[/IMG]愤愤地说：气气气死我拉、、、、">气死
<option value="//乐呵呵地对%%叫：我是流氓我怕谁，哈哈[IMG]33.gif[/IMG]">流氓</option>
<option value="//一张圆脸流下两行清泪哭着哀求%%，你就饶了俺吧，别打我[IMG]dz21.gif[/IMG]">求饶<option value="//见%%不怀好意，连忙抓起电话：114，119，120，122，96315 ... 心想：“警察叔叔怎么还不来呀？”">求助
<option value="//看都不看%%一眼，哼了一声，高高的把头扬了起来……不理你……">不理
<option value="//对着%%挥舞着静月江湖神拳[IMG]dz06.gif[/IMG]">拳打
<option value="//把%%踩在地下。。[IMG]dz07.gif[/IMG]">踩扁
<option value="//练成了“杀猪掌”，一掌把%%搞成这个造型了。[IMG]dz05.gif[/IMG]">棉掌
<option value="//悄悄地把一包砒霜放在%%的碗里，躲在一旁，喜滋滋地看着她喝了下去。>下毒
<option value="//向%%勇敢地跪了下来：“你愿意嫁给我吗”──真是勇气可嘉啊！音乐鼓励[IMG]8.gif[/IMG]">求婚
<option value="//“啪啪啪”地扇了%%几个耳光，满脸通红说：强盗！[IMG]dz02.gif[/IMG]”">耳光
<option value="//对着%%大声号啕：是大唐的兄弟姐妹救救我吧[IMG]dz03.gif[/IMG]”">救救
<option value="//轻轻地摸着%%的脸，无限深情地说：“我好喜欢你哦......”">喜欢
<option value="//抓起%%往烤箱里一放，大概过了五分钟，打开箱一看大吃一惊：怎么成了烤卤猪啦!!!!">卤猪
<option value="//见%%快不行了，连忙召来[IMG]dz60.gif[/IMG]，把%%送进精神病院……">救护
<option value="//轻轻的亲了%%一下，好深情呦……[IMG]dz57.GIF[/IMG]">亲热
<option value="//%%突然间被人蒙住了双眼，睁眼一看，##站在面前，手上放着好大一簇玫瑰[img]dz61.gif[/img]！(">玫瑰 
<option value="//高兴的吻了吻%%的小脸，满怀对下一代殷切的期望说道：“%%小鬼好聪明哦！看到你我就放心啦。不过要注意小小的年纪别太早熟喽”。%%听见夸奖，嘿嘿傻笑起来。">聪明
<option value="//轻轻的搂着%%，含情脉脉的低语：“今晚皓月当空，你看那嫦娥仙子也羡慕我们耶！”">轻搂
<option value="//对%%说道：“阁下听了这样的话，居然脸不红，心不跳，真是功力深厚！不知在下这等条件，可否学到您本领的万一？”">皮厚
<option value="//赶紧挡住腰包，对%%求饶道：“小兄弟您可千万别搞错了，在下穷得响叮噹呀！这腰包看起来鼓鼓的，其实里面可都是树叶子呀！”">很穷
<option value="//深情地对%%说道：“一年三百六十五天，一天二十四小时，每分每秒我都在这里等着你......”">等你
<option value="//望着%%离去的背影，泪水止不住簌簌而下，起歌和道：“天之涯，海之角，知交半零落。一瓢浊酒尽余欢，今宵别梦寒。”">离别
<option value="//长剑当胸，几绺长发飘荡风中，那神情气度皆已完美到了极点，不知不觉%%的一颗芳心已牢牢系在了##的身上。">气度
<option value="//怪眼圆翻，有气无力的对着%%说道：“连我皇帝都不急，你太监急有屁用……？”">着急<option value="//用力拍着%%的背，一副“好小子！干的好”的表情。">拍背
<option value="//以藐视的目光盯着%%说道：“就凭你.....也配？？！！”">轻蔑
<option value="//从耳朵上拿下一颗皱巴巴的烟卷儿递给%%：来一根儿？">敬烟
<option value="//朗声说道：“%%是老同志了，虽然脑子有点糊涂了，但是大家还是应该尊敬她嘛！”">尊敬
<option value="//眼中快喷出火来了，对着%%说道:“青山不改绿水长流，君子报仇十年不晚！！！”">报仇
<option value="//搂住%%的纤腰，指着初升的月亮说：“天上明月，是我们的证人。”">证人
<option value="//望着%%手中的刀，对她叹到：“牡丹树下死，做鬼也风流！”">风流
<option value="//对%%大声叫道：怕什么怕，掉个脑袋不过碗大的疤，你小子有胆就跟他拼一场！">英雄<option value="//看着%%默默远走的背影，看着他在风中飘散的长发和那孤寂的背影，泪水止不住滑落到胸口，心中呼喊着“%%，其实我真的爱你”。">离开
<option value="//忽然叹道：“人在江湖还是收敛一点好，否则下场难免要和%%一样，死了都还不知道是谁下的手！”">小心

</select>
</td>
  </tr>
  <tr>
    <td width="100%" align="center">
      <table border="0" width="12%">
        <tr>
              <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a  href="#" onClick="window.open('../hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="当要使用药品、装备物品时点这里！"><font size="-1"><img src="img/gn/wupin.gif" width="30" height="19" border="0"></font></a></td>
              <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../bbs/index.asp" target="_blank"><img src="img/gn/luntan.gif" width="30" height="19" border="0" alt="江湖信息|会友交流"></a></td>        </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#" onClick="window.open('../hcjs/jhjs/dan.asp')" ><img src="img/gn/dangpu.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../hcjs/jhjs/yaopu.asp')" title="江湖神医，悬壶济世"><img src="img/gn/yaodian.gif" width="30" height="19" border="0"></a></td>        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../hcjs/jhjs/wuqi.asp')" title="绝世武器店"><img src="img/gn/bingqi.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../hcjs/card/card.asp"  target="_blank" title="会员卡片屋"><img src="img/gn/kapin.gif" width="30" height="19" border="0"></a></td>        </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../yamen/laofan.asp')" title="少做点坏事哦...不然...嘿嘿..."><img src="img/gn/laofang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../yamen/minan.asp')" title="江湖命案"><img src="img/gn/mingan.gif" width="30" height="19" border="0"></a></td></tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" nowrap width="3%"><a href="#"onClick="window.open('../hcjs/yilao.asp')" title="无论多重的伤，医仙胡青牛会让你起死回生的"><img src="img/gn/liaoshang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a href="work.asp" target="f3"  title="我的职业-我的工作"><font size="-1"><img src="img/gn/gongzuo.gif" width="30" height="19" border="0"></font></a></td></tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a href="#" onClick="window.open('../yamen/regmodi.asp','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')" title="江湖档案修改"><font size="-1"><img src="img/gn/ziliao.gif" width="30" height="19" border="0"></font></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" height="2" width="1%"><a href="../Diary/Diary.asp" target="_blank" title="签写心情日记"><img src="img/gn/rji.gif" width="30" height="19" border="0"></a></td>      </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jl/jlmain.asp" target="_blank" title="江湖每天发生的大事"><img src="img/gn/dashi.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jhjd/jd.asp" target="_blank" title="E线大酒店"><img src="img/gn/jiudian.gif" width="30" height="19" border="0"></a></td>
        </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../hcjs/xuewu/xuetang.htm" target="_blank" title="江湖武馆"><img src="img/gn/wuguan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jhmp/wuguan.asp" target="_blank" title="修炼绝世神功...得拜师傅才能进入"><img src="img/gn/mishi.gif" width="30" height="19" border="0"></a></td>
          </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../top/top.htm" target="_blank" title="江湖侠客排行"><img src="img/gn/paihang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../gupiao/stock.asp" target="_blank" title="买卖自由，想发就来试试运气"><img src="img/gn/gupiao.gif" width="30" height="19" border="0"></a></td>
                  </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../dk/Daikuan.asp" target="_blank" title="到期不还钱会被杀手追杀哦"><img src="img/gn/daikuan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../bet/betindex.asp" target="_blank" title="江湖赌馆"><img src="img/gn/duguan.gif" width="30" height="19" border="0"></a></td>
        </tr>
        <tr>
            <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="12%"><a href="../jhmp/index.asp" target="_blank" title="闯荡江湖找个好老大照顾准没错"><img src="img/gn/menpai.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../dg/dg.asp" target="_blank" title="打工赚钱"><img src="img/gn/dagong.gif" width="30" height="19" border="0"></a></td>
        </tr>
        <tr>

    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="58%"><a href="../jhmp/gry.asp" target="_blank" title="积德行善的好事，会增加道德哦"><img src="img/gn/guer.gif" width="30" height="19" border="0"></a></td>
          <td></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
      </center>
    </div>
