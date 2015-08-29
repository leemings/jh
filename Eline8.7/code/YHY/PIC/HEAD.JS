// 上海热线ICP评比流量统计
 function OnlineICP()
 {
 	//var snect = "<iframe marginheight=0 marginwidth=0 frameborder=0  width=0 height=0 scrolling=yes src='http://61.129.163.253/cgi-bin/collect.cgi?icp_id=20'></iframe>";
	var snect = "<iframe marginheight=0 marginwidth=0 frameborder=0  width=0 height=0 scrolling=yes src=''></iframe>";
	document.writeln(snect);
 }


function GetCardPay()
{
	if (confirm("您想使用九城帐户里的钱求签吗？\n每求一张签您需要花费0.1金币"))
	{
		//window.open("/lot2/ctl_lot2_card.php?action=getcardpay",'card','left=150,top=50,width=600,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
		OpenCard("/lot2/ctl_lot2_card.php?action=getcardpay");
	}
}

function AddAction()
{
	if (confirm("您想使用九城帐户里的钱加行动力吗？\n你确定花费1金币增加40点行动力吗？"))
	{
		//window.open("/lot2/ctl_lot2_card.php?action=getcardpay",'card','left=150,top=50,width=600,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
		//OpenWin("");
		window.open("/lot2/ctl_lot2_set.php?action=addaction",'king','left=200,top=150,width=400,height=300,resizable=1,status=0,menubar=0,scrollbars=1');
	}
}

var n,ie;
if (document.all)    {n=0;ie=1;fShow="visible";fHide="hidden";}
if (document.layers) {n=1;ie=0;fShow="show";   fHide="hide";}

var lastMenu = null;
var leftX = 0, leftY = 0;
var rightX = 0, rightY = 0;
function displaySubMenu(idMainMenu)
{
	var submenu;
	if (n)
	{
		submenu = document.layers[idMainMenu];
		if (lastMenu != null && lastMenu != submenu) hideAll();
		submenu.visibility = fShow;

		leftX  = document.layers[idMainMenu].left;
		rightX = leftX + 100;
		leftY  = document.layers[idMainMenu].top+
				 document.layers[idMainMenu].clip.height;
		rightY = leftY;
	} else if (ie) {
		submenu = eval(idMainMenu+".style");
		submenu.visibility = fShow;
		if (lastMenu != null && lastMenu != submenu) hideAll();

		leftX  = document.all[idMainMenu].style.posLeft;
		rightX = leftX + 100;
		leftY  = document.all[idMainMenu].style.posTop+
				 document.all[idMainMenu].offsetHeight;
		rightY = leftY;
	}
	lastMenu = submenu;
}

function hideAll()
{
	if (lastMenu != null) {lastMenu.visibility = fHide; lastMenu = null;}
}

function updateIt(e)
{
	if (ie)
	{
		var x = window.event.clientX;
		var y = window.event.clientY;

		if (x > rightX || x < leftX) hideAll();
		else if (y > rightY) hideAll();
	}
	if (n)
	{
		var x = e.pageX;
		var y = e.pageY;

		if (x > rightX || x < leftX) hideAll();
		else if (y > rightY) hideAll();
	}
}

function ShowHeadi(){document.write("<table width=778 cellspacing=0 cellpadding=0 border=0><tr bgcolor=#000000><td><img src=http://img.the9.com/img/empty.gif width=778 height=2></td></tr><tr bgcolor=#65251C><td background=http://img.the9.com/img/lot2/topbg.gif height=59><table cellspacing=0 cellpadding=0 border=0><tr><td><img src=http://img.the9.com/img/empty.gif width=21 height=59></td><td><img src=http://img.the9.com/img/lot2/icon1.gif></td><td width=680 align=right><img src=http://img.the9.com/img/empty.gif width=680 height=12><br><img src=http://img.the9.com/img/empty.gif width=22 height=24><a href=/static/lot2/register/bgstory.htm><img src=http://img.the9.com/img/lot2/button/bgstory.gif border=0 onmouseover=\"this.className='gr'\" onmouseout=\"this.className='nogr'\"></a><img src=http://img.the9.com/img/empty.gif width=22 height=24><a href=/static/lot2/register/rule.htm><img src=http://img.the9.com/img/lot2/button/chanyue.gif border=0 onmouseover=\"this.className='gr'\" onmouseout=\"this.className='nogr'\"></a><img src=http://img.the9.com/img/empty.gif width=22 height=24><a href=http://member.the9.com/billing/index.php><img src=http://img.the9.com/img/lot2/button/paycenter.gif border=0 onmouseover=\"this.className='gr'\" onmouseout=\"this.className='nogr'\"></a><img src=http://img.the9.com/img/empty.gif width=22 height=24><a href=/static/lot2/register/pic.htm><img src=http://img.the9.com/img/lot2/button/picdown.gif border=0 onmouseover=\"this.className='gr'\" onmouseout=\"this.className='nogr'\"></a><img src=http://img.the9.com/img/empty.gif width=22 height=24><a href=/static/lot2/register/vh.htm onclick=OpenKing() target=king><img src=http://img.the9.com/img/lot2/button/intro2.gif border=0 onmouseover=\"this.className='gr'\" onmouseout=\"this.className='nogr'\"></a><img src=http://img.the9.com/img/empty.gif width=22 height=24><a href=/main.htm  target=_top><img src=http://img.the9.com/img/lot2/button/tothe9.gif border=0 onmouseover=\"this.className='gr'\" onmouseout=\"this.className='nogr'\"></a></td><td width=20><img src=http://img.the9.com/img/empty.gif width=1 height=59></td></tr></table></td></tr><tr bgcolor=#FFFDDF><td><img src=http://img.the9.com/img/empty.gif width=778 height=1></td></tr><tr bgcolor=#65251C><td><img src=http://img.the9.com/img/empty.gif width=778 height=3></td></tr></table>");
 OnlineICP();
 }

function ShowFoot(style)
{
	if (arguments.length == 0) style = 0;
	if (style == 0)
	{
		var snect = "<table width=100% border=0 cellspacing=0 cellpadding=0>"+
		"<tr><td background=http://img.the9.com/img/lot2/all/line.gif><img src=http://img.the9.com/img/empty.gif width=1 height=36></td></tr></table>"+
		"<table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor=#CCCCCC align=center>"+
		"<tr><td><img src=http://img.the9.com/img/empty.gif width=774 height=10></td></tr>"+
		"<tr><td align=center><br><a href=http://www.the9.com/ class=copy1 target=GameNow>第九城市计算机技术咨询（上海）有限公司提供全部技术支持</a><span class=copy1><br>本游戏的故事及情节纯属虚构，所有涉及的国家、地区名称均与现实无关</span><br><br></td></tr></table>";
	}
	else
	{
		var snect = "<table width=534 border=0 cellspacing=0 cellpadding=0 align=center>"+
		"<tr><td><img src=http://img.the9.com/img/empty.gif height=10></td></tr>"+
		"<tr><td align=center><br><a href=http://www.the9.com/ class=copy1 target=GameNow>第九城市计算机技术咨询（上海）有限公司提供全部技术支持</a><span class=copy1><br>本游戏的故事及情节纯属虚构，所有涉及的国家、地区名称均与现实无关</span><br><br></td></tr></table>";
	}
	document.write(snect);
}

function ShowHead(type)
{
	if (arguments.length == 0) type = 1;
	
	// 第一级分类
	var snect = "<table width=100% border=0 cellspacing=0 cellpadding=0 background=http://img.the9.com/img/lot2/main/head_bg.gif>"+
				"<tr>"+
				"<td width=555><img src=http://img.the9.com/img/empty.gif height=8><br>"+
				"<table width=555 border=0 cellspacing=0 cellpadding=0>"+
				"<tr align=right>"+
				"<td width=55><a href=/lot2/><img src=http://img.the9.com/img/lot2/main/bt-home.gif border=0 onmouseover=\"this.className='gr'\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-polity.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu1')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-cult.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu2')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-fun.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu3')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-move.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu4')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-special.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu5')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-news.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu6')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-lot.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu7')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-commu.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu8')\" onmouseout=this.className=''></a></td>"+
				"<td width=55><a href=#><img src=http://img.the9.com/img/lot2/main/bt-prop.gif border=0 onmouseover=\"this.className='gr';displaySubMenu('Menu9')\" onmouseout=this.className=''></a></td>"+
				"<td width=20></td>"+
				"</tr></table>"+
				"<img src=http://img.the9.com/img/empty.gif height=10><br>"+
				"</td>"+
				"<td width=12><img src=http://img.the9.com/img/lot2/main/head_bg_b.gif></td>"+
				"<td width=45>"+
				"<img src=http://img.the9.com/img/empty.gif height=8><br>"+
				"<table width=45 border=0 cellspacing=0 cellpadding=0>"+
				"<tr align=right>"+
				"<td><a href=http://www.the9.com/main.htm target=_top alt=退出游戏，返回九城中心广场><img src=http://img.the9.com/img/lot2/main/head_arrow.gif border=0></a></td>"+
				"<td><a href=http://www.the9.com/static/lot2/ alt=给与你足够的引导><img src=http://img.the9.com/img/lot2/main/head_hlp.gif border=0></a></td>"+
				"</tr></table>"+
				"<img src=http://img.the9.com/img/empty.gif height=10><br>"+
				"</td>"+
				"<td><img src=http://img.the9.com/img/empty.gif height=1 width=1></td></tr></table>";
	
	var snect2 = "<table width=100% border=0 cellspacing=0 cellpadding=0 background=http://img.the9.com/img/lot2/main/head_bg.gif>"+
				"<tr>"+
					"<td width=555>"+
						"<div id=\"marquee1\" style=\"position:absolute;left:20;top:8;z-index:2;\">"+
						"<marquee direction=left width=520 scrollamount=3 scrolldelay=60 onmouseover=this.stop() onmouseout=this.start() class=redtext>"+
						"<font color=white>欢迎来到纵横，我们十分欢迎您的到来！希望您在这里的生活能够愉快！</font></marquee></div>"+
					"</td>"+
					"<td width=12><img src=http://img.the9.com/img/lot2/main/head_bg_b.gif></td>"+
					"<td width=45>"+
						"<img src=http://img.the9.com/img/empty.gif height=8><br>"+
						"<table width=45 border=0 cellspacing=0 cellpadding=0>"+
						"<tr align=right>"+
						"<td><a href=/main.htm target=_top alt=退出游戏，返回九城中心广场><img src=http://img.the9.com/img/lot2/main/head_arrow.gif border=0></a></td>"+
						"<td><a href=javascript:ViewHelp() alt=给与你足够的引导><img src=http://img.the9.com/img/lot2/main/head_hlp.gif border=0></a></td>"+
						"</tr></table>"+
						"<img src=http://img.the9.com/img/empty.gif height=10><br>"+
					"</td>"+
					"<td><img src=http://img.the9.com/img/empty.gif height=1 width=1></td>"+
				"</tr></table>";
	
	// 政治
	var sMenu1=	"<div id=\"Menu1\" style=\"position:absolute; left:44; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_polity.php?action=list onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=投票选举新君主>选举君主</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_lord.php?action=office  class=ct-def1 title=国家官员上朝>上　　朝</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_polity.php?action=politymain class=ct-def1 title=更改国籍>国　　籍</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_polity.php?action=purse onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=捐献虚拟货币给国家>募　　捐</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_polity.php?action=applyrank class=ct-def1 title=申请成为贵族>爵位申请</a>&nbsp;&nbsp;</td></tr></table></div>";
	
	// 修行
	var sMenu2=	"<div id=\"Menu2\" style=\"position:absolute; left:100; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_work.php?action=index class=ct-def1 title=参加劳动获得回报>工　　作</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_study.php?action=main class=ct-def1 title=丰富知识提高技能>学　　习</a>&nbsp;&nbsp;</td></tr></table></div>";

	//娱乐
	var sMenu3= "<div id=\"Menu3\" style=\"position:absolute; left:154; top:30; visibility:hidden;z-index=:9000;\">"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_bet.php?action=index class=ct-def1 title=顷刻之间腰缠万贯>博　　彩</a>&nbsp;&nbsp;</td></tr></table></div>";
	
	// 移动
	var sMenu4=	"<div id=\"Menu4\" style=\"position:absolute; left:205; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_move.php?action=move&iReturn=1 class=ct-def1 title=回到自己的祖国>返回祖国</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_move.php?action=choice class=ct-def1 title=移动到任意国家>自由移动</a>&nbsp;&nbsp;</td></tr></table></div>";
	
	// 特殊
	var sMenu5=	"<div id=\"Menu5\" style=\"position:absolute; left:260; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_move.php?action=visitdesc onclick='javascript:OpenWin();' class=ct-def1 target='newwin'  title=游览本地名胜古迹>游览名胜</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=getcards&cCardPlace=23  class=ct-def1  title=营救所在国家监狱中的玩家>营　　救</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=getcards&cCardPlace=24 class=ct-def1  title=抢夺所在国家他人财物>抢　　劫</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=getcards&cCardPlace=19 class=ct-def1  title=暗杀所在国家别的玩家>刺　　杀</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=getcards&cCardPlace=21 class=ct-def1  title=破坏所在国家防御>破　　坏</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=getcards&cCardPlace=20 class=ct-def1  title=在所在国家造谣>谣　　言</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=getcards&cCardPlace=22 class=ct-def1  title=引起所在国家骚乱>骚　　乱</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=getcards&cCardPlace=18  class=ct-def1  title=探听所在国家情报>打　　探</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_spec.php?action=prisonerlist onclick='javascript:OpenWin();' target='newwin'  class=ct-def1  title=参观所在国家的监狱>参观监狱</a>&nbsp;&nbsp;</td></tr>"+
		"</table></div>";
		
	// 情报
	var sMenu6=	"<div id=\"Menu6\" style=\"position:absolute; left:300; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td>　<a href=/static/lot2/info/map.htm class=ct-def1  onclick='javascript:OpenMap();' target='map' title=查看81国方位>查看世界地图</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_info.php?action=grainprice  onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=查看各地粮价>查看粮食价格</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_info.php?action=weaponprice  onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=查看各地武器价格>查看兵器价格</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_info.php?action=competition  onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=查看当前国家排名>查看国家排名</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_list.php?action=hero  onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=纵横英雄榜>纵横英雄榜</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_list.php?action=loser  onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=纵横悲情榜>纵横悲情榜</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_list.php?action=credit  onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=纵横功臣榜>纵横功臣榜</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_info.php?action=getmessage onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=查看用户事件>查看英雄传</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/static/lot2/info/searchplayer.htm onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=查看指定玩家信息>查找用户信息</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/static/lot2/info/searchlocation.htm onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=查看指定玩家所在国家>查找用户位置</a>&nbsp;&nbsp;</td></tr></table></div>";
	
	// 求签
	var sMenu7=	"<div id=\"Menu7\" style=\"position:absolute; left:370; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td>　<a href=javascript:OpenCard(\"/lot2/ctl_lot2_card.php?action=getcardfree\") class=ct-def1 title=该求签需要花费虚拟货币>求　　签</a>&nbsp;&nbsp;</td></tr>"+
		//"<tr><td>　<a href=javascript:GetCardPay() class=ct-def1 title=该求签需要花费九城金币>求付费签</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_card.php?action=mycardlist class=ct-def1 title=查看自己拥有的签>查看签库</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/card/readme.htm onclick='javascript:OpenWin();' target='newwin' class=ct-def1 title=求签说明>求签说明</a>&nbsp;&nbsp;</td></tr></table></div>";
	
	// 交流
	var sMenu8=	"<div id=\"Menu8\" style=\"position:absolute; left:425; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		"<tr><td class=redtext>　<a href=http://www.the9.com/message/ctl_message.php?action=send onclick='javascript:OpenWin(80, 50, 700, 460);' class=ct-def1 target='newwin'>发&nbsp;消&nbsp;息</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td class=redtext>　<a href=http://bbs.the9.com/bbs/index.php?forum=746 class=ct-def1 target=GameNow>论　　坛</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td class=redtext>　<a href=http://www.the9.com/chat/ class=ct-def1>聊&nbsp;天&nbsp;室</a>&nbsp;&nbsp;</td></tr></table></div>";

	
	// 属性
	var sMenu9=	"<div id=\"Menu9\" style=\"position:absolute; left:464; top:30; visibility:hidden;z-index:9000;\" >"+
		"<table border=1 cellspacing=2 cellpadding=2 bordercolor=505960 bgcolor=c0cbcc>"+
		//"<tr><td>　<a href=/lot2/ctl_lot2_set.php?action=getability class=ct-def1 title=重新分配>个人能力评定</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_set.php?action=getphoto class=ct-def1 title=修改个人信息>修改照片列传</a>&nbsp;&nbsp;</td></tr>"+
		//"<tr><td>　<a href=/lot2/ctl_lot2_set.php?action=getmoney class=ct-def1 title=将九城虚拟货币转入游戏或转出>转换虚拟货币</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/ctl_lot2_info.php?action=self onclick='javascript:OpenWin();' class=ct-def1 target='newwin' title=查看自己的状态>查看个人信息</a>&nbsp;&nbsp;</td></tr>"+
		//"<tr><td>　<a href=javascript:AddAction() class=ct-def1 title=增加行动力>加&nbsp;行&nbsp;动&nbsp;力</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/info/readme.htm class=ct-def1 title=属性说明>个人属性说明</a>&nbsp;&nbsp;</td></tr>"+
		"<tr><td>　<a href=/lot2/info/readme_work.htm class=ct-def1 title=工作详细说明>工作详细说明</a>&nbsp;&nbsp;</td></tr></table></div>";

	var sEventIE="<script language=javascript>document.body.onclick=hideAll;document.body.onscroll=hideAll;document.body.onmousemove=updateIt;</script>";
	var sEventNS="<script language=javascript>document.onmousedown=hideAll;window.captureEvents(Event.MOUSEMOVE);window.onmousemove=updateIt;</script>";

	if(type != 2)
	{
		if (navigator.appName=="Netscape") {document.write(snect+sMenu1+sMenu2+sMenu3+sMenu4+sMenu5+sMenu6+sMenu7+sMenu8+sMenu9+sEventNS);}
		else {document.write(snect+sMenu1+sMenu2+sMenu3+sMenu4+sMenu5+sMenu6+sMenu7+sMenu8+sMenu9+sEventIE);}
	}
	else
	{
		document.write(snect2);
	}
	//OnlineICP();
}

function ShowHeadPop(poptext){
if (poptext =="") {poptext="纵 横";}
var snect1="<table width=505 border=0 cellspacing=0 cellpadding=0><tr><td><img src=http://img.the9.com/img/empty.gif width=3 height=21><img src=http://img.the9.com/img/lot2/icon/icont1.gif></td><td><table border=0 cellspacing=0 cellpadding=0><tr><td rowspan=3><img src=http://img.the9.com/img/empty.gif width=14></td><td rowspan=3><img src=http://img.the9.com/img/lot2/frame/wtdr1.gif width=11></td><td background=http://img.the9.com/img/lot2/frame/wtdrbg1.gif><img src=http://img.the9.com/img/empty.gif width=103 height=2></td><td rowspan=3><img src=http://img.the9.com/img/lot2/frame/wtdr2.gif width=9></td><td rowspan=3><img src=http://img.the9.com/img/empty.gif width=8></td></tr><tr><td class=redtext align=center height=15>::"
var snect2="::</td></tr><tr><td background=http://img.the9.com/img/lot2/frame/wtdrbg2.gif><img src=http://img.the9.com/img/empty.gif width=103 height=4></td></tr></table></td><td><img src=http://img.the9.com/img/empty.gif width=254 height=10></td><td><a href=\"javascript:window.close();\"><img src=http://img.the9.com/img/lot2/button/close.gif border=0></a><img src=http://img.the9.com/img/empty.gif width=30 height=21></td></tr></table>"
document.write(snect1+poptext+snect2);
OnlineICP();
}


function OpenWin(iLeft, iTop, iWidth, iHeight)
{
	if (arguments.length == 0){iLeft=150; iTop=50; iWidth=600; iHeight=450;}
	window.open("", "newwin", "left="+iLeft+", top="+iTop+", width="+iWidth+", height="+iHeight+", resizable=1,status=0,menubar=0,scrollbars=1");
}

function OpenKing(){
window.open("",'king','left=150,top=50,width=600,height=450,resizable=1,status=0,menubar=0,scrollbars=1');}

function OpenMap(){
window.open("",'map','left=0,top=0,width=800,height=420,resizable=1,status=0,menubar=0,scrollbars=1');}

function OpenCard(lcard){
if(navigator.appName=="Netscape"){
splashWin=window.open(lcard,'card','resizable=1,status=0,menubar=0,scrollbars=0');
splashWin.resizeTo(300,481);
splashWin.moveTo(120,10);}
else{
splashWin=window.open(lcard,'card','resizable=0,status=0,menubar=0,scrollbars=0');
splashWin.resizeTo(310,531);
splashWin.moveTo(120,10);}}

function ViewHelp(name)
{
        if (arguments.length == 0)
        {
                name = "";
        }
        if (n)
        {
                name = escape(name);
        }
        if (name == "")
        {
                var url = escape(window.location);
                var helpurl = "/help/index.php?url=" + url;
        }
        else
        {
                var helpurl = "/help/index.php?kind=" + name;
        }
        window.open(helpurl, 'helpwin', 'left=100, top=80, width=540,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}
 
