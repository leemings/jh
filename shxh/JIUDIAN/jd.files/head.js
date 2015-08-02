
// 显示浮层
function ShowMarquee(iPlaceId)
{
	if(GetBannerUrl(iPlaceId) != "") mode = 1;
	else mode = 2;
	if(iPlaceId == 16 || iPlaceId == 17 || iPlaceId == 18 || iPlaceId == 19) mode = 3;
	if(iPlaceId == 8) mode = 4;
	if(iPlaceId == 1) mode = 5;
	//alert(mode)
	switch(mode)
	{
		case 1: l = 430; t = 115; w = 316; h = 16; show = true; color='ffff00'; break;
		case 2: l = 400; t = 81; w = 270; h = 16; show = true; color='ffccff'; break;
		case 3: l = 10; t = 70; w = 264; h = 16; show = true; color='ff3366'; break;
		case 4: l = 400; t = 112; w = 270; h = 16; show = true; color='ffff00'; break;
		default: show = false; break;
	}

	if(show)
	{
		msg_1 = "九城付费全过程：1、到付费中心选择适合的充值方式；2、网上进行充值；3、购买九城各项付费服务，成为付费用户。";
		msg_2 = "九城充值便捷方式——九城卡充值：有20块、50块两种零售，除上海外全国开设昆明、南宁、南京、桂林、济南等销售点，并有两家配送公司全国范围内送卡。";
		msg_3 = "九城付费便捷方式——13种银行卡全国支持：招商银行一网通、中国工商银行牡丹信用卡、九城充值便捷方式：中行长城借记卡、中行长城信用卡、  九城充值便捷方式：上海浦东发展银行东方卡（仅限北京地区） 、建设银行网上银行（深圳）、 九城充值便捷方式：工行存折账户（仅限北京地区）、上海浦东发展银行存折账户（仅限北京地区）、 建设银行网上银行（北京）";
		msg_4 = "四种国际银行卡VISA Card、MASTER Card、American Express Card、JCB Card 支持九城付费，国际付费畅通无阻！";
		msg_5 = "九城套卡由侠客、圣骑士、精灵王、光明使者、大天使五张九城卡组成，其中侠客、精灵王、天使卡只有在套卡中才能获得，不会单独出售，全套210元。印制精美，值得收藏，九城忠实居民的至爱！";
		msg_6 = "九城付费中心服务热线：021-52984545；服务时间：9:00—21:00";
		if(iPlaceId == 15 || iPlaceId == 16 || iPlaceId == 8 || iPlaceId == 9) 
			textCont = 
			"<a href=http://billing.the9.com/billing/ style='font size:9pt;' target=GameNow><font color="+color+">"+msg_1+"</font></a>　　　"+
			"<a href=http://billing.the9.com/billing/ style='font size:9pt;' target=GameNow><font color="+color+">"+msg_2+"</font></a>　　　"+
			"<a href=http://billing.the9.com/billing/ style='font size:9pt;' target=GameNow><font color="+color+">"+msg_3+"</font></a>　　　";
		else 
			textCont = 
			"<a href=http://billing.the9.com/billing/ style='font size:9pt;' target=GameNow><font color="+color+">"+msg_4+"</font></a>　　　"+
			"<a href=http://billing.the9.com/billing/ style='font size:9pt;' target=GameNow><font color="+color+">"+msg_5+"</font></a>　　　"+
			"<a href=http://billing.the9.com/billing/ style='font size:9pt;' target=GameNow><font color="+color+">"+msg_6+"</font></a>　　　";
			showCont = "<div id=ShowMarquee style='position: absolute; left: "+l+"px; top: "+t+"px; width: "+w+"px; height: "+h+"px'>"+
			"<marquee width="+w+" scrolldelay=2 scrollamount=2>"+textCont+"</marquee></div>";
	}
	else showCont = "";
	return showCont;
	
}

 // 得到中心广场及各个小区的banner link
 function GetBannerUrl(iPlaceId)
 {
	switch(iPlaceId)
	{
		case 1: // 中心广场
			//var sBannerUrl = "pl-7-top_1";	
			var sBannerUrl = "CP_the9|CG_CenterSquare|C_TOP";		
			break;
		case 2: // 电脑网络区
			var sBannerUrl = "CP_the9|CG_Area|C_it";			
			break;
		case 3: // 体育健身区
			var sBannerUrl = "CP_the9|CG_Area|C_sport";			
			break;
		case 4: // 游戏动漫区
			var sBannerUrl = "CP_the9|CG_Area|C_game";			
			break;
		case 5: // 人文情感区
			var sBannerUrl = "CP_the9|CG_Area|C_feeling";			
			break;
		case 6: // 时尚生活区
			var sBannerUrl = "CP_the9|CG_Area|C_pop";		
			break;
		case 7: // 商业区
			var sBannerUrl = "";		
			break;
		case 8: // 职业介绍所
			//var sBannerUrl = "pl-11-work_1";	
			var sBannerUrl = "CP_the9|CG_Work|C_index";	
			break;
		case 9: // 宝物商店
			//var sBannerUrl = "pl-9-trea_1";
			var sBannerUrl = "CP_the9|CG_Treasure|C_TOP";		
			break;
		case 10: // 知识问答
			var sBannerUrl = "CP_the9|CG_comp|C_top";
			break;
		case 14: // bbs
			//var sBannerUrl = "pl-8-bbs_1";
			var sBannerUrl = "CP_the9|CG_BBS|C_TOP";		
			break;
		case 15: // chat
			var sBannerUrl = "";		
			break;
		case 16: // home_index
			var sBannerUrl = "CP_the9|CG_MyHome|C_livingroom";
			break;
		case 17: // bedroom
			var sBannerUrl = "CP_the9|CG_MyHome|C_bedroom";
			break;
		case 18: // livingroom
			var sBannerUrl = "CP_the9|CG_MyHome|C_livingroom";
			break;
		case 19: // kitchen
			var sBannerUrl = "CP_the9|CG_MyHome|C_livingroom";
			break;
		case 99://卓越商城
		    //var sBannerUrl = "pl-12-zhuoyue";	
			var sBannerUrl = "CP_the9|CG_mall|C_zhuoyue";
			break;
		default:
			var sBannerUrl = "";
			break;
	}
	return sBannerUrl;
 }
 
 // 得到重定向的连接 --modified by telyn shi: add redirect.php
 function Redirect(iPlaceId)
 {
	var sReUrl = "";
	switch(iPlaceId)
	{
		case 14: // bbs 
			sReUrl = "http://bbs.the9.com/bbs/redirect.php?url=";	
			break;
		case 20: // vsm 
			sReUrl = "http://vsm.the9.com/vsm/redirect.php?url=";	
			break;
		case 21: // sms
			sReUrl = "http://sms.the9.com:9001/sms/redirect.jsp?url=";	
			break;
		case 22: // billing
			sReUrl = "http://www.the9.com";	
			break;
		case 99: // 卓越
		    sReUrl = "http://www.the9.com";
			break;
		case 100: // 佳能
		    sReUrl = "http://www.the9.com";
			break;
		default:
			break;
	}
	return sReUrl;
 }
 
 // 上海热线ICP评比流量统计
 function OnlineICP()
 {
 	//var snect = "<iframe marginheight=0 marginwidth=0 frameborder=0  width=0 height=0 scrolling=yes src='http://61.129.163.253/cgi-bin/collect.cgi?icp_id=20'></iframe>";
	var snect = "<iframe marginheight=0 marginwidth=0 frameborder=0  width=0 height=0 scrolling=yes src=''></iframe>";
	document.writeln(snect);
 }
 
 
 // 得到中心广场及各个小区的TopMenu
 function GetTopMenuNew(sBannerUrl,iPlaceId,img_src)
 {
	/*if(iPlaceId == 1 || iPlaceId == 8 || iPlaceId == 14 || iPlaceId == 99)
	{
		var BannerCont = "<IFRAME SRC='http://61.151.253.46/adsunion/get/;pl="+sBannerUrl+";tp=if;sk=0;ck=0;/?'  MARGINHEIGHT=0 MARGINWIDTH=0 FRAMEBORDER=0 BORDER=0 VSPACE=0 WIDTH=468 NORESIZE HEIGHT=60 SCROLLING=NO >"+
		"<script language=JavaScript src='http://61.151.253.46/adsunion/get/;pl="+sBannerUrl+";tp=if;sk=0;ck=0;/?'></script></IFRAME>";
	}
	else
	{*/	
		var BannerCont = "<IFRAME MARGINHEIGHT=0 MARGINWIDTH=0 FRAMEBORDER=0 WIDTH=468 HEIGHT=60 SCROLLING=NO SRC=http://the9.allyes.com/main/adfshow?user="+sBannerUrl+"&db=the9&border=0&local=yes>"+
		"<SCRIPT LANGUAGE=JavaScript1.1 SRC='http://the9.allyes.com/main/adfshow?user="+sBannerUrl+"&db=the9&local=yes&js=on'>"+
		"</SCRIPT><NOSCRIPT>"+
		"<A HREF=http://the9.allyes.com/main/adfclick?user="+sBannerUrl+"&db=the9>"+
		"<IMG SRC=http://the9.allyes.com/main/adfshow?user="+sBannerUrl+"&db=the9 WIDTH=468 HEIGHT=60 BORDER=0></a>"+
		"</NOSCRIPT></IFRAME>";
 	//}

	 var TopMenu = "<table border=0 width=776 cellspacing=0 cellpadding=0 height=72>"+
	  	"<tr><td width=155>"+
		"<a href="+Redirect(iPlaceId)+"/main.php target=GameNow><img border=0 src=http://img.the9.com/img/9logo.gif width=155 height=72 alt=第九城市标志></a>"+
		"</td>"+
	    "<td width=498 align=center valign=middle>"+
		BannerCont+"</td>"+
		"<td width=63 valign=top><a href=javascript:ViewHelp()><img border=0 src=http://img.the9.com/img/help.gif width=63 height=27 alt=如何在第九城市生活></a></td>"+
		"<td width=60 valign=top><a href=javascript:logout()><img border=0 src=http://img.the9.com/img/logout.gif width=60 height=27 alt=离开第九城市></a></td>"+
		"</tr></table>";
		
	var TopMenu1= "<table border=0 width=776 cellspacing=0 cellpadding=0 height=72>"+
	  	"<tr><td width=155>"+
		"<a href="+Redirect(iPlaceId)+"/main.php target=GameNow><img border=0 src=http://img.the9.com/img/9logo.gif width=155 height=72 alt=第九城市标志></a>"+
		"</td>"+
	    "<td width=498 align=center valign=middle>"+
		BannerCont+"</td>"+
		"<td width=123 valign=top align=right><img src=" + img_src + " border=0 width=155 height=72></td>"+
		"</tr></table>";
		
	 if (iPlaceId == 99) return TopMenu1;//卓越
	 else	 return TopMenu;
}

 
// TopMenu's map
function GetTopMenu(iPlaceId)
{
	var TopMap = "<map name=TopMenu>"+
	"<area href="+Redirect(iPlaceId)+"/billing/index.php shape=rect coords=\"1, 0, 92, 21\" alt=进入付费中心，享用更多的服务项目 target=GameNow>"+
	"<area href=javascript:CheckMail() shape=rect coords=\"91, 0, 181, 21\" alt=查看我的消息>"+
	"<area href="+Redirect(iPlaceId)+"/service/ shape=rect coords=\"182, 0, 254, 21\"  alt=新手入门  target=GameNow>"+
	"<area href=javascript:ViewGuid() shape=rect coords=\"253, 0, 324, 21\" alt=浏览城市导航图>"+
	"<area href=javascript:ViewHelp() shape=rect coords=\"323, 0, 395, 21\" alt=在九城生活有疑问吗，看看帮助吧 target=GameNow>"+
	"<area href=javascript:logout() shape=rect coords=\"392, 0, 437, 21\" alt=离开九城></map>";
	
	return TopMap;
}

// the image of TopMenu's map
var TopMapImgSrc = "http://img.the9.com/img/topmenu.gif";
var TopMapImgWidth = 438, TopMapImgHeight = 22;


// 在"我的家"里使用的TopMenu
function showTopMenuInHome(id)
{
	if (arguments.length == 0)	id = 999;
	if (id != 999)
	{
		var snect =  "<table border=0 width=760 cellspacing=0 cellpadding=0 height=33>"+ 
		"<tr><td width=256 height=33>"+ 
		"<a href=/main.php target=GameNow><img src=http://img.the9.com/img/home/homelgo1.gif width=256 height=33 border=0 alt=回到中心广场></a><br>"+ 
		"<a href=/main.php target=GameNow><img src=http://img.the9.com/img/home/homelgo2.gif width=256 height=41 border=0 alt=回到中心广场></a>"+ 
		"</td><td class=ct-def1 align=right width=504 align=right>"+ 
		GetTopMenu(id) + "<img border=0 src="+TopMapImgSrc+" usemap=#TopMenu " +
		"width="+TopMapImgWidth+" height="+TopMapImgHeight+"><br>"+	
		"<IFRAME MARGINHEIGHT=0 MARGINWIDTH=0 FRAMEBORDER=0 WIDTH=468 HEIGHT=60 SCROLLING=NO SRC=http://the9.allyes.com/main/adfshow?user="+GetBannerUrl(id)+"&db=the9&border=0&local=yes>"+
		"<SCRIPT LANGUAGE=JavaScript1.1 SRC='http://the9.allyes.com/main/adfshow?user="+GetBannerUrl(id)+"&db=the9&local=yes&js=on'>"+
		"</SCRIPT><NOSCRIPT>"+
		"<A HREF=http://the9.allyes.com/main/adfclick?user="+GetBannerUrl(id)+"&db=the9>"+
		"<IMG SRC=http://the9.allyes.com/main/adfshow?user="+GetBannerUrl(id)+"&db=the9 WIDTH=468 HEIGHT=60 BORDER=0></a>"+
		"</NOSCRIPT></IFRAME>"+
		"</td></tr>"+
		"<tr><td colspan=2 height=5><img src=http://img.the9.com/img/empty.gif></td></tr></table>";
	}
	else
	{
		var snect =  GetTopMenu(1) + 
			"<img border=0 src="+TopMapImgSrc+" usemap=#TopMenu " +
			"width="+TopMapImgWidth+" height="+TopMapImgHeight+">";
	}
	document.writeln(snect+ShowMarquee(17));
	OnlineICP();
}

function SmallItem(iPlaceId)
{
	var con = "<td width=49><a href=javascript:ViewGuid()><img border=0 src=http://img.the9.com/img/ico1.gif width=63 height=39 alt=城市地图></a></td>"+
	"<td width=49><a href=javascript:CheckMail()><img border=0 src=http://img.the9.com/img/ico2.gif width=55 height=39 alt=我的消息></a></td>"+
	"<td width=49><a href="+Redirect(iPlaceId)+"/billing/index.php target=GameNow><img border=0 src=http://img.the9.com/img/ico3.gif width=68 height=39 alt=付费中心></a></td>";
	return con;
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
		rightX = leftX + document.layers[idMainMenu].clip.width;
		leftY  = document.layers[idMainMenu].top+
			document.layers[idMainMenu].clip.height;
		rightY = leftY;
	} else if (ie) {
		submenu = eval(idMainMenu+".style");
		submenu.visibility = fShow;
		if (lastMenu != null && lastMenu != submenu) hideAll();

		leftX  = document.all[idMainMenu].style.posLeft;
		rightX = leftX + document.all[idMainMenu].offsetWidth;
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
	/*
	if (n)
	{
		var x = e.pageX;
		var y = e.pageY;

		if (x > rightX || x < leftX) hideAll();
		else if (y > rightY) hideAll();
	}*/
}

var business_area_url = "/area/business/business.php";

// 学习
function FM1(left,top,iPlaceId)
{
	var sMenu1="<div id=\"Menu1\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/ad/the9/index.htm class=ct-def9 target=GameNow>九城指南</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/service/ class=ct-def9 target=GameNow>新手学堂</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href=javascript:ViewGuid() class=ct-def9>城市导航</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/comp/ class=ct-def9 target=GameNow>知识问答</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu1;
}

// 生活
function FM2(left,top,iPlaceId)
{
	var sMenu2="<div id=\"Menu2\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/feel/ctl_register.php?action=index class=ct-def9 target=GameNow>九城喜屋</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/feel/ctl_wedding.php?action=rest class=ct-def9 target=GameNow>九城大饭店</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/treasure/ class=ct-def9 target=GameNow>宝物商城</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/facility/ class=ct-def9 target=GameNow>城市事务所</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/court/ctl_court.php?action=index class=ct-def9>九城法院</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/gov/gov.htm class=ct-def9>九城市政府</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/news/index.php class=ct-def9 target=GameNow>九城通讯社</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu2;
}

// 闲逛
function FM3(left,top,iPlaceId)
{
	var sMenu3="<div id=\"Menu3\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/area/it/it.php class=ct-def9 target=GameNow>电脑网络区</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/area/sport/sport.php class=ct-def9 target=GameNow>体育健身区</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/area/game/game.php class=ct-def9 target=GameNow>游戏动漫区</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/area/pop/pop.php class=ct-def9 target=GameNow>时尚生活区</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/area/feeling/feeling.php class=ct-def9 target=GameNow>人文情感区</a></td></tr>"+
	//"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+business_area_url+" class=ct-def9 target=GameNow>商　业　区</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu3;
}

// 社交
function FM4(left,top,iPlaceId)
{
	var sMenu4="<div id=\"Menu4\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/gather/ class=ct-def9 target=GameNow>聚会中心</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/card/ctl_card.php?action=classify class=ct-def9 target=GameNow>贺卡礼品</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href=javascript:WriteMail() class=ct-def9>短&nbsp;消&nbsp;息</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href=http://bbs.the9.com/bbs/bbsindex.php class=ct-def9 target=GameNow>论　　坛</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/chat/ class=ct-def9 target=GameNow>聊　　天</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu4;
}

// 娱乐
function FM5(left,top,iPlaceId)
{
	var sMenu5="<div id=\"Menu5\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
    "<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/sms/index.php class=ct-def9 target=GameNow>SMS短信中心</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/agon/ctl_agon.php?action=index class=ct-def9 target=GameNow>九城竞猜中心</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/fmt/ class=ct-def9 target=GameNow>九城足球经理</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/my_way/my_way.htm class=ct-def9 target=GameNow>摘星工厂</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/gamecard/ class=ct-def9 target=GameNow>GameOnline</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu5;
}

// 探险
function FM6(left,top,iPlaceId)
{
	var sMenu6="<div id=\"Menu6\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/lot/ class=ct-def9 target=GameNow>穿越时空（纵横）</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/baker/ class=ct-def9 target=GameNow>侦破案情（贝克街）</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu6;
}

// 购物
function FM7(left,top,iPlaceId)
{
	var sMenu7="<div id=\"Menu7\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/mall/ class=ct-def9 target=GameNow>九城专卖</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href=http://joyo.the9.com/default.asp?source=the9 class=ct-def9 target=GameNow>卓越音像</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href=http://the9.ebid.com.cn/ class=ct-def9 target=GameNow>易必得精品商城</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu7;
}

// 投资
function FM8(left,top,iPlaceId)
{
	var sMenu8="<div id=\"Menu8\" style=\"position:absolute;left:"+left+";top:"+top+";visibility:hidden;z-index:9;\" >"+
	"<table border=0 cellspacing=0 cellpadding=0 bgcolor=#cccccc><tr><td><table border=0 cellspacing=1 cellpadding=0>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/vsm/vsm_login.php class=ct-def9 target=GameNow>九城股市</a></td></tr>"+
	//"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/ad/stock/vsm/main.htm class=ct-def9 target=GameNow>九城股市</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/ad/stock/03.htm class=ct-def9 onclick=OpenNews() target=newswin>九城银行</a></td></tr>"+
	"<tr><td bgcolor=#FEFDCE width=110 align=center><a href="+Redirect(iPlaceId)+"/static/ad/stock/02.htm class=ct-def9 onclick=OpenNews() target=newswin>九城保险</a></td></tr>"+
	"</table></td></tr></table></div>";
	return sMenu8;
}

var sEventIE="<script language=javascript>document.body.onclick=hideAll;document.body.onscroll=hideAll;document.body.onmousemove=updateIt;</script>";
var sEventNS="<script language=javascript>document.onmousedown=hideAll;window.captureEvents(Event.MOUSEMOVE);window.onmousemove=updateIt;</script>";

function GetSnect(iPlaceId,way)
{
	var snect = 
	"<table border=0 width=776 cellspacing=0 cellpadding=0 height=39><tr>"+
	"<td width=86><a href="+Redirect(iPlaceId)+"/main.php target=GameNow><img border=0 src=http://img.the9.com/img/c00.gif width=86 height=39 alt=回到中心广场></a></td>"+
	"<td width=49><a href="+Redirect(iPlaceId)+"/home/ target=GameNow><img border=0 src=http://img.the9.com/img/c01.gif width=49 height=39 alt=回家></a></td>"+
	"<td width=49><a href="+Redirect(iPlaceId)+"/work/ target=GameNow><img border=0 src=http://img.the9.com/img/c02.gif width=49 height=39 alt=工作></a></td>"+
	"<td width=49><a href=javascript:displaySubMenu('Menu1') onmouseover=\"displaySubMenu('Menu1')\" onclick=\"displaySubMenu('Menu1')\" target=GameNow><img border=0 src=http://img.the9.com/img/c03.gif width=49 height=39 ></a></td>"+ //alt=学习
	"<td width=49><a href=javascript:displaySubMenu('Menu2') onmouseover=\"displaySubMenu('Menu2')\" onclick=\"displaySubMenu('Menu2')\" target=GameNow><img border=0 src=http://img.the9.com/img/c04.gif width=49 height=39 ></a></td>"+ //alt=生活
	"<td width=49><a href=javascript:displaySubMenu('Menu3') onmouseover=\"displaySubMenu('Menu3')\" onclick=\"displaySubMenu('Menu3')\" target=GameNow><img border=0 src=http://img.the9.com/img/c05.gif width=49 height=39 ></a></td>"+ //alt=闲逛
	"<td width=49><a href=javascript:displaySubMenu('Menu4') onmouseover=\"displaySubMenu('Menu4')\" onclick=\"displaySubMenu('Menu4')\" target=GameNow><img border=0 src=http://img.the9.com/img/c06.gif width=49 height=39 ></a></td>"+ //alt=社交
	"<td width=49><a href=javascript:displaySubMenu('Menu5') onmouseover=\"displaySubMenu('Menu5')\" onclick=\"displaySubMenu('Menu5')\" target=GameNow><img border=0 src=http://img.the9.com/img/c07.gif width=49 height=39 ></a></td>"+ //alt=娱乐
	"<td width=49><a href=javascript:displaySubMenu('Menu8') onmouseover=\"displaySubMenu('Menu8')\" onclick=\"displaySubMenu('Menu8')\" target=GameNow><img border=0 src=http://img.the9.com/img/c08.gif width=49 height=39 ></a></td>"+ //alt=投资
	"<td width=49><a href=javascript:displaySubMenu('Menu6') onmouseover=\"displaySubMenu('Menu6')\" onclick=\"displaySubMenu('Menu6')\" target=GameNow><img border=0 src=http://img.the9.com/img/c09.gif width=49 height=39 ></a></td>"+ //alt=探险
	"<td width=49><a href="+Redirect(iPlaceId)+"/mall/shopping_index.htm onmouseover=\"displaySubMenu('Menu7')\" onclick=\"displaySubMenu('Menu7')\" target=GameNow><img border=0 src=http://img.the9.com/img/c10.gif width=63 height=39 ></a></td>"+ //alt=购物
	SmallItem(iPlaceId)+"</tr></table>";
	
	var snect1 = 
	"<a href="+Redirect(iPlaceId)+"/main.php target=GameNow><img border=0 src=http://img.the9.com/img/c00.gif width=86 height=39 alt=回到中心广场></a>"+
	"<a href="+Redirect(iPlaceId)+"/home/ target=GameNow><img border=0 src=http://img.the9.com/img/c01.gif width=49 height=39 alt=回家></a>"+
	"<a href="+Redirect(iPlaceId)+"/work/ target=GameNow><img border=0 src=http://img.the9.com/img/c02.gif width=49 height=39 alt=工作></a>"+
	"<a href=javascript:displaySubMenu('Menu1') onmouseover=\"displaySubMenu('Menu1')\" onclick=\"displaySubMenu('Menu1')\" target=GameNow><img border=0 src=http://img.the9.com/img/c03.gif width=49 height=39 ></a>"+ //alt=学习
	"<a href=javascript:displaySubMenu('Menu2') onmouseover=\"displaySubMenu('Menu2')\" onclick=\"displaySubMenu('Menu2')\" target=GameNow><img border=0 src=http://img.the9.com/img/c04.gif width=49 height=39 ></a>"+ //alt=生活
	"<a href=javascript:displaySubMenu('Menu3') onmouseover=\"displaySubMenu('Menu3')\" onclick=\"displaySubMenu('Menu3')\" target=GameNow><img border=0 src=http://img.the9.com/img/c05.gif width=49 height=39 ></a>"+ //alt=闲逛
	"<a href=javascript:displaySubMenu('Menu4') onmouseover=\"displaySubMenu('Menu4')\" onclick=\"displaySubMenu('Menu4')\" target=GameNow><img border=0 src=http://img.the9.com/img/c06.gif width=49 height=39 ></a>"+ //alt=社交
	"<a href=javascript:displaySubMenu('Menu5') onmouseover=\"displaySubMenu('Menu5')\" onclick=\"displaySubMenu('Menu5')\" target=GameNow><img border=0 src=http://img.the9.com/img/c07.gif width=49 height=39 ></a>"+ //alt=娱乐
	"<a href=javascript:displaySubMenu('Menu8') onmouseover=\"displaySubMenu('Menu8')\" onclick=\"displaySubMenu('Menu8')\" target=GameNow><img border=0 src=http://img.the9.com/img/c08.gif width=49 height=39 ></a>"+ //alt=投资
	"<a href=javascript:displaySubMenu('Menu6') onmouseover=\"displaySubMenu('Menu6')\" onclick=\"displaySubMenu('Menu6')\" target=GameNow><img border=0 src=http://img.the9.com/img/c09.gif width=49 height=39 ></a>"+ //alt=探险
	"<a href="+Redirect(iPlaceId)+"/mall/shopping_index.htm onmouseover=\"displaySubMenu('Menu7')\" onclick=\"displaySubMenu('Menu7')\" target=GameNow><img border=0 src=http://img.the9.com/img/c10.gif width=63 height=39 ></a>"+ //alt=购物
	"</td></tr></table>";
	
	if(way == 1) return snect;
	if(way == 2) return snect1;
}

function ShowHeaderMainNew(id)
{
	if (arguments.length == 0)	id = 1;
	snect = GetTopMenuNew(GetBannerUrl(id),id,'')+GetSnect(id,1)+
	FM1(183,112,id)+FM2(232,112,id)+FM3(281,112,id)+FM4(330,112,id)+FM5(379,112,id)+FM6(478,112,id)+FM7(527,112,id)+FM8(429,112,id);
	if (navigator.appName=="Netscape") {document.writeln(snect+sEventNS+ShowMarquee(id));}
	else {document.writeln(snect+sEventIE+ShowMarquee(id));}
	OnlineICP();
}

function ShowHeaderIt()
{
	ShowHeaderMainNew(2);
}

function ShowHeaderSport()
{
	ShowHeaderMainNew(3);
}


function ShowHeaderGame()
{
 	ShowHeaderMainNew(4);
}

function ShowHeaderFeeling()
{
  	ShowHeaderMainNew(5);
}

function ShowHeaderPop()
{
  	ShowHeaderMainNew(6);
}

function ShowHeaderOld()
{
  	ShowHeaderMainNew(6);
}

function ShowHeaderMain(id)
{
	if (arguments.length == 0)	id = 11;
	snect = GetTopMenu(id)+
		"<table border=0 width=760 height=77 cellspacing=0 cellpadding=0>" +
		"<tr><td width=155 height=77 valign=top rowspan=3><a href="+Redirect(id)+"/main.php target=GameNow>" +
		"<img border=0 src=http://img.the9.com/img/9logo.gif width=155 height=72 alt=回到中心广场></a></td>" +
		"    <td width=582 height=22 align=right>" +
		"<img border=0 src="+TopMapImgSrc+" usemap=#topmenu width="+TopMapImgWidth+" height="+TopMapImgHeight+"></td></tr>" +
		"<tr><td height=16><img src=http://img.the9.com/img/empty.gif></td></tr>" +
		"<tr><td height=31>"+GetSnect(id,2)+
		FM1(339,78,id)+FM2(388,78,id)+FM3(437,78,id)+FM4(486,78,id)+FM5(534,78,id)+FM6(633,78,id)+FM7(682,78,id)+FM8(583,78,id);
	if (navigator.appName=="Netscape") {document.writeln(snect+sEventNS+ShowMarquee(id));}
	else {document.writeln(snect+sEventIE+ShowMarquee(id));}
	OnlineICP();
}

function ShowHeaderNews()
{
  	ShowHeaderMain(1);
}

function ShowHeaderAgon()
{
  	ShowHeaderMain(11);
}

function ShowHeaderTrea()
{
  	//ShowHeaderMain(9);
	ShowHeaderMainNew(9);
}

function ShowHeaderBbs()
{
	//ShowHeaderMain(14);
	ShowHeaderMainNew(14);
}

function ShowHeaderChat()
{
     ShowHeaderMain(15);
}

function ShowHeader()
{
	var snect = " <table border=0 width=100% height=59 cellspacing=0 cellpadding=0> " +
		" <tr><td width=261><a href=http://www.the9.com/main.php target=GameNow><img border=0 src=http://img.the9.com/img/newsheadlogo.gif width=261 height=59 border=0 alt=回中心广场></a></td> " +
		" <td width=100%><p align=right> " +
		" <a href=javascript:ViewHelp()><img border=0 src=http://img.the9.com/img/helpico.gif width=66 height=35 alt=查看帮助信息></a></td></tr></table> " +
		" <table border=0 width=100% height=1 bgcolor=#000000 cellspacing=0 cellpadding=0> " +
		" <tr><td><img src=http://img.the9.com/img/empty.gif></td></tr></table> " +
		" <table border=0 width=100% cellspacing=0 cellpadding=0> " +
		" <tr><td bgcolor=#cccccc height=3><img src=http://img.the9.com/img/empty.gif></td></tr> " +
		" <tr><td bgcolor=#000000 height=1><img src=http://img.the9.com/img/empty.gif></td></tr></table> " ;
	
	document.writeln(snect);
	OnlineICP();
}

function newsHeader(id)
{
	 if (arguments.length == 0)	id = "";
	 if (id == "")
	 {		 
		var snect = " <table border=0 width=100% height=59 cellspacing=0 cellpadding=0> " +
			" <tr><td width=261><a href=http://www.the9.com/main.php target=GameNow><img src=http://img.the9.com/img/newsheadlogo.gif width=261 height=59 border=0 alt=回中心广场></a></td> " +
			" <td width=100% align=right valign=bottom><a href=\"javascript:ViewHelp()\"><img border=0 src=http://img.the9.com/img/newshelp.gif></a> " +
			" <a href=\"javascript:window.close()\"><img border=0 src=http://img.the9.com/img/newsclose.gif></a>&nbsp;&nbsp;</td></tr></table> " +
			" <table border=0 width=100% cellspacing=0 cellpadding=0> " +
			" <tr><td height=1 bgcolor=#000000><img src=http://img.the9.com/img/empty.gif></td></tr></table>" +
			" <table border=0 width=100% cellspacing=0 cellpadding=0> " +
			" <tr><td bgcolor=#cccccc height=3><img src=http://img.the9.com/img/empty.gif></td></tr> " +
			" <tr><td bgcolor=#000000 height=1><img src=http://img.the9.com/img/empty.gif></td></tr> " +
			" </table> " ;
	 }
	 else
	{
		var snect = " <table border=0 width=100% height=59 cellspacing=0 cellpadding=0> " +
			" <tr><td width=261><img src=http://img.the9.com/img/newsheadlogo.gif width=261 height=59></td> " +
			" <td width=100% align=right valign=bottom>"+	
			"<br><iframe src=/ad/show_banner.php?page_id="+ id +" frameborder=0 marginheight=0"+
			" marginwidth=0 scrolling=no width=468 height=60>"+
			"<script language=JavaScript src=/ad/show_banner.php?page_id="+ id +"></SCRIPT>"+
			"</iframe></td></tr></table>"+
			" <table border=0 width=100% cellspacing=0 cellpadding=0> " +
			" <tr><td height=1 bgcolor=#000000><img src=http://img.the9.com/img/empty.gif></td></tr></table>"+
			" <table border=0 width=100% cellspacing=0 cellpadding=0> " +
			" <tr><td bgcolor=#cccccc height=3><img src=http://img.the9.com/img/empty.gif></td></tr> " +
			" <tr><td bgcolor=#000000 height=1><img src=http://img.the9.com/img/empty.gif></td></tr> " +
			" </table>" +
			" <table border=0 width=100% cellspacing=0 cellpadding=0 bgcolor=#FDFDE8> " +
			"<tr><td width=95% align=right><br> "+
			"<a href=\"javascript:ViewHelp()\"><img border=0 src=http://img.the9.com/img/newshelp.gif></a> " +
			" <a href=\"javascript:window.close()\"><img border=0 src=http://img.the9.com/img/newsclose.gif></a>&nbsp;&nbsp; " +
			"</td></tr></table>";
	}
	 document.writeln(snect);
	 OnlineICP();
}

function ShowCommerceHeader(img_url,name)
{
	var title='';
	id=99;
	if (name == 'joyo') title='卓越音像'; 
	else if (name == 'ebid') title='易必得精品商城';
	snect = GetTopMenuNew(GetBannerUrl(id),id,img_url)+GetSnect(id,1)+
	"<table width=776 border=0 cellspacing=0 cellpadding=0 bgcolor=823C17>" +
    "<tr><td width=48><img src=http://img.the9.com/img/empty.gif></td>" +
    "<td width=34><img src=http://img.the9.com/img/main02.gif></td>" +
    "<td width=660 class=head-l>您当前位置：第九城市-&gt;<a href=http://www.the9.com/main.php target=GameNow class=head-l>中心广场</a>-&gt;"+
    "<a href=http://www.the9.com/mall/shopping_index.htm class=head-l>商业区</a>-&gt;" + title + "</td>" +
    "<td width=18 align=right valign=top><img src=http://img.the9.com/img/main03.gif></td>" +
    "</tr></table>" +
	FM1(183,112,id)+FM2(232,112,id)+FM3(281,112,id)+FM4(330,112,id)+FM5(379,112,id)+FM6(478,112,id)+FM7(527,112,id)+FM8(429,112,id);
	if (navigator.appName=="Netscape") {document.writeln(snect+sEventNS);}
	else {document.writeln(snect+sEventIE);}
	OnlineICP();
}

function ShowFooterNoLink()
{
	var snect = " <table border=0 width=100% cellspacing=0 cellpadding=0> " +
		" <tr><td height=1 width=100% bgcolor=#000000><img src=http://img.the9.com/img/empty.gif></td></tr> " +
		" <tr><td height=3 bgcolor=#cccccc width=100%><img src=http://img.the9.com/img/empty.gif></td></tr></table> " +
		" <table border=0 width=100% cellspacing=0 cellpadding=0 align=center> " +
		" <tr><td height=15></td></tr> " +
		" <tr><td align=center><span class=ct-def1>第九城市计算机技术咨询（上海）有限公司提供全部技术支持</span></td> " +
		" </tr><tr><td height=20></td></tr></table> " ;
	
	document.writeln(snect);	
}

function ShowFooter(id)
{
	 if (arguments.length == 0)	id = "";
	 var snect = "";

	 if (id == 1)
	 {
	 	snect = "<table width=776 border=0 cellspacing=0 cellpadding=0>"+
  				"<tr><td bgcolor=FCE88C width=104>&nbsp;</td>"+
    			"<td bgcolor=FCE88C class=ct-def1 width=757><br>"+
				//"中心广场 | 回家 | 工作 | 学习 | 生活 | 闲逛 | 社交 | 娱乐 | 投资 | 探险 | 购物 "+
				"</td>"+
    			"<td bgcolor=FCE88C width=19><img src=http://img.the9.com/img/main31.gif width=19 height=26></td>"+
  				"</tr></table>";
	}
	else
	{		
		snect = " <table width=760 bgColor=#666666 border=0 cellspacing=0 cellpadding=0 height=3> " +
	             " <tr><td width=100%><img src=http://img.the9.com/img/empty.gif></td></tr> " +
	             " </table> ";
	 }
	 
 	 snect = snect + " <table border=0 width=760 cellspacing=0 cellpadding=0> " +
	    " <tr><td align=center class=ct-def2><br>"+
		"<br><br>| <a href="+Redirect(id)+"/home/ target=GameNow>我的家</a> | <a href="+Redirect(id)+"/news/ target=GameNow>新闻社</a> | <a href="+Redirect(id)+"/agon/ctl_agon.php?action=index target=GameNow>竞猜中心</a> | <a href="+Redirect(id)+"/treasure/ target=GameNow>宝物商城</a> | <a href=http://bbs.the9.com/bbs/bbsindex.php target=GameNow>论坛</a> | <a href="+Redirect(id)+"/chat/ target=GameNow>聊天室</a> |<br>"+
		"| <a href="+Redirect(id)+"/main.php target=GameNow>中心广场</a> | <a href="+Redirect(id)+"/area/it/it.php target=GameNow>电脑网络区</a> | <a href="+Redirect(id)+"/area/sport/sport.php target=GameNow>体育健身区</a> | <a href="+Redirect(id)+"/area/game/game.php target=GameNow>游戏动漫区</a> | <a href="+Redirect(id)+"/area/pop/pop.php target=GameNow>时尚生活区</a> | <a href="+Redirect(id)+"/area/feeling/feeling.php target=GameNow>人文情感区</a> | <a href="+Redirect(id)+business_area_url+" target=GameNow>商业区</a> |<br>"+
		"<br></td></tr> " +
		" <tr><td><p align=center><span class=ct-def1>| <a href=javascript:ViewGuid() class=ct-def2>城市导航图</a> | <a href="+Redirect(id)+"/other/corp_intro.htm target=blank class=ct-def2>公司介绍</a> | <a class=ct-def4 href="+Redirect(id)+"/adser/index.htm>广告服务</a> | <a class=ct-def2 href=mailto:webmaster@the9.com>与我们联系</a> | <a class=ct-def2 href="+Redirect(id)+"/other/feedback.htm target=blank>客户服务</a> |</span><br> " +
		" <span class=ct-def1>第九城市计算机技术咨询（上海）有限公司提供全部技术支持</span></td> " +
		" </tr><tr><td><br></td></tr></table> " ;

 	 document.writeln(snect);
}


function ShowCommerceFooter()
{
	var snect = " <table border=0 width=760 cellspacing=0 cellpadding=0> " +
		    " <tr><td align=center class=ct-def2><br>"+
			"<br><br>| <a href=http://www.the9.com/home/ target=GameNow>我的家</a> | <a href=http://www.the9.com/news/ target=GameNow>新闻社</a> | <a href=http://www.the9.com/agon/ctl_agon.php?action=index target=GameNow>竞猜中心</a> | <a href=http://www.the9.com/treasure/ target=GameNow>宝物商城</a> | <a href=http://bbs.the9.com/bbs/bbsindex.php target=GameNow>论坛</a> | <a href=http://www.the9.com/chat/ target=GameNow>聊天室</a> |<br>"+
			"| <a href=http://www.the9.com/main.php target=GameNow>中心广场</a> | <a href=http://www.the9.com/area/it/it.php target=GameNow>电脑网络区</a> | <a href=http://www.the9.com/area/sport/sport.php target=GameNow>体育健身区</a> | <a href=http://www.the9.com/area/game/game.php target=GameNow>游戏动漫区</a> | <a href=http://www.the9.com/area/pop/pop.php target=GameNow>时尚生活区</a> | <a href=http://www.the9.com/area/feeling/feeling.php target=GameNow>人文情感区</a> | <a href=http://www.the9.com" +business_area_url+" target=GameNow>商业区</a> |<br>"+
			"<br></td></tr> " +
			" <tr><td><p align=center><span class=ct-def1>| <a href=javascript:ViewGuid() class=ct-def2>城市导航图</a> | <a href=http://www.the9.com/other/corp_intro.htm target=blank class=ct-def2>公司介绍</a> | <a class=ct-def4 href=http://www.the9.com/adser/index.htm>广告服务</a> | <a class=ct-def2 href=mailto:webmaster@the9.com>与我们联系</a> | <a class=ct-def2 href=http://www.the9.com/other/feedback.htm target=blank>客户服务</a> |</span><br> " +
			" <span class=ct-def1>第九城市计算机技术咨询（上海）有限公司提供全部技术支持</span></td> " +
			" </tr><tr><td><br></td></tr></table> " ;
	document.writeln(snect);		
}

function JoyCityBBSMain()
{
  var JoyCityHeader="<table width=100% height=150 border='0' cellspacing='0' cellpadding='0' bgcolor=FF9701>"+
					"<tr valign='top'><td>"+
  		"<table border=0 cellspacing=0 cellpadding=0 width=778 bgcolor=FF9701>"+
  		"<tr><td><img src=http://joycity.the9.com/joycity/img/headjs/head01.gif>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/joycity.gif>"+
		"<a href=http://www.the9.com/ target=_blank><img src=http://joycity.the9.com/joycity/img/headjs/the9.gif border=0></a>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/boygirl.jpg>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/leg.gif>"+
		"<img src=http://joycity.the9.com/joycity/img/index/empty.gif width=13 height=78></td></tr>"+
		"<tr><td bgcolor=#CFFFFF><img src=http://joycity.the9.com/joycity/img/headjs/head02.gif>"+
		"<a href=http://joycity.the9.com/joycity/intro/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/intro.gif border=0></a>"+
		"<img src=http://joycity.the9.com/joycity/img/index/empty.gif width=1>"+
		"<a href=http://joycity.the9.com/joycity/intro/bgstory/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/bgstory.gif border=0></a>"+
		"<img src=http://joycity.the9.com/joycity/img/index/empty.gif width=1>"+
		"<a href=http://joycity.the9.com/joycity/fresh/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/new.gif border=0></a>"+
		"<img src=http://joycity.the9.com/joycity/img/index/empty.gif width=5>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/boygirl1.jpg>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/orange.gif width=31 height=21></td></tr>"+
		"<tr><td><img src=http://joycity.the9.com/joycity/img/headjs/head03.gif>"+
		"<a href=http://joycity.the9.com/joycity/manual/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/use.gif border=0></a>"+
		"<a href=http://joycity.the9.com/joycity/text/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/south.gif border=0></a>"+
		"<a href=http://joycity.the9.com/joycity/skill/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/tech.gif border=0></a>"+
		"<a href=http://joycity.the9.com/joycity/feel/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/download.gif border=0></a>"+
		"<a href=http://joycity.the9.com/joycity/download/index.htm target=GameNow>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/player.gif border=0></a>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/leg1.gif>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/foot.gif></td></tr>"+
		"<tr><td><table border=0 width=100% cellspacing=0 cellpadding=0><tr><td>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/head04.gif></td><td>"+
		"<img src=http://joycity.the9.com/joycity/img/headjs/bottom.gif></td>"+
		"<td class=newstopic-h width=493 background=http://joycity.the9.com/joycity/img/headjs/foot1.gif align=center>"+
		"<a href=http://joycity.the9.com/joycity/main.htm target=GameNow>回首页</a>&nbsp;&nbsp;<a href=http://joycity.the9.com/joycity/billing/index.htm target=GameNow>付费中心</a>"+
		"&nbsp;&nbsp;<a href=http://joycity.the9.com/joycity/serve/index.htm target=GameNow>服务体系</a>&nbsp;&nbsp;"+
		"<a href=http://joycity.the9.com/joycity/help/index.htm target=GameNow>帮助</a></td></tr></table></td></tr></table></td></tr></table>";
  document.write(JoyCityHeader);
  OnlineICP();
}
function WarBibleBBSMain()
{
	var WarBibleHeader ="<table width=100% height=109 border='0' cellspacing='0' cellpadding='0' bgcolor=#1F277F>"+
						"<tr valign='top'><td>"+
						"<table width='778' border='0' cellspacing='0' cellpadding='0'>"+
						"<tr><td><img src='http://warbible.the9.com/warbible/img/main/bbstop.gif' width='529' height='109' usemap='#Map' border='0'>"+
						"<map name='Map'><area shape='rect' coords='250,59,324,79' href='http://warbible.the9.com/warbible/main.htm' target=_blank>"+
						"<area shape='rect' coords='407,60,481,80' href='http://warbible.the9.com/warbible/help.htm' target=_blank>"+
						"<area shape='rect' coords='329,59,400,81' href='http://warbible.the9.com/warbible/billing.htm' target=_blank>"+
						"<area shape='rect' coords='34,80,120,102' href='http://warbible.the9.com/warbible/intro.htm' target=_blank>"+
						"<area shape='rect' coords='123,80,202,102' href='http://warbible.the9.com/warbible/fresh.htm' target=_blank>"+
						"<area shape='rect' coords='201,80,278,102' href='http://warbible.the9.com/warbible/text.htm' target=_blank>"+
						"<area shape='rect' coords='278,80,361,102' href='http://warbible.the9.com/warbible/feel.htm' target=_blank>"+
						"<area shape='rect' coords='362,80,440,102' href='http://warbible.the9.com/warbible/download.htm' target=_blank>"+
						"<area shape='rect' coords='442,79,521,102' href='http://warbible.the9.com/warbible/service.htm' target=_blank>"+
						"</map></td><td width='249'>"+
            			"</td></tr></table></td></tr></table>"+
						"<div id='Layer1' style='position:absolute; left:529px; top:0px; width:249px; height:136px; z-index:1'>"+
						"<img src='http://warbible.the9.com/warbible/img/main/bbstop2.gif' width='249' height='136' usemap='#Map2' border='0'>"+
						"<map name='Map2'>"+
						"<area shape='rect' coords='22,57,166,100' href='http://www.the9.com/' target=_blank></map>"+
						"</div>";
	document.write(WarBibleHeader);
}
