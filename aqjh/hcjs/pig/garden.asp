<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<!--#include file="../../showchat.asp"-->
<%
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
sjname1=LCase(trim(request.querystring("sjname")))
if application("aqjh_sanghua")="" or application("aqjh_sanghua")=null then
	Response.Write "<script language=JavaScript>{alert('你想作什么，没有人邀请你呀！');}</script>"
	Response.End
end if
b=application("aqjh_sanghua")
ddd=split(b,"|")
sjname=ddd(0)
myname=ddd(1)
shsj=ddd(2)
erase ddd
if sjname=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('切！还有自已给自已观赏的呀，再这样下去小心变成自恋狂！');}</script>"
	Response.End
end if
if myname<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('你想作什么，人家["&sjname&"]又没有邀请你！');}</script>"
	Response.End
end if
Application.Lock
Application("aqjh_sanghua")=""
Application.UnLock
nowsj=DateDiff("s",shsj,now())
if nowsj>=60 then
	Response.Write "<script language=JavaScript>{alert('错误！您的邀请时间已超过60秒，超时，不能观赏猪场。');}</script>"
	Response.End
end if
Sub Msg (v)
 Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><title>种花</title><meta http-equiv='pragma' content='no-cache'><style type=text/css>body{color:black;font-family:宋体;font-size:9pt;background-color:buttonface;border-bottom:medium none;border-left:medium none;border-right:medium none;border-top:medium none;padding-bottom:0px;padding-left:0px;padding-right:0px;padding-top:0px}</style></head><body leftMargin=0 topMargin=0 marginheight=0 marginwidth=0>"
 Response.Write "<script Language=JavaScript>alert('" & v & "');window.close();</script></body></html>"
 Response.End
End Sub
If Session("aqjh_name") = "" Then Msg "您尚未登录，不能赏猪。"
name = Session("aqjh_name")
If name = "" Then Msg "您尚未登录，不能赏猪。"
if request("myid")<>"" then name=request("myid")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 银两,金币,道德,魅力,体力,内力 FROM 用户 WHERE 姓名='"& session("aqjh_name") &"'",conn,1,3
if rs("银两")<100000 or rs("金币")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "身上没100000银两或金币不足1个，不能赏猪！"
	response.end
end if
if rs("道德")<5000 or rs("魅力")<10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "道德不够1000或魅力不够5000，不能赏猪！"
	Response.End
end if
if rs("体力")<8000 or rs("内力")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "赏花要耗废8000体力和5000内力，你也太虚了吧！"
	response.end
end if
rs("银两")=rs("银两")-100000
rs("金币")=rs("金币")-1
rs("体力")=rs("体力")-8000
rs("内力")=rs("内力")-5000
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE><%=name%>的养猪场</TITLE>
<script>
window.moveTo(40,30);
</script>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta http-equiv=refresh content='150;url=hua.asp'>
<STYLE type=text/css>A {FONT: 9pt 宋体; TEXT-TRANSFORM: none; TEXT-DECORATION: none}
A:link {TEXT-TRANSFORM: none; COLOR: #003300; TEXT-DECORATION: none}
A:visited {	COLOR: #0d660d}
A:active {	COLOR: #ff9900}
A:hover {	COLOR: #ff9900}
.9pt {FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal}
.10pt {FONT-WEIGHT: normal; FONT-SIZE: 10pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal}
BODY {FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal}
DIV {FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal}
TD {FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal}
P {	FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal}
.aq {  font-size: 9pt; line-height: 150%; color: #000000; filter: DropShadow(Color=#efefef, OffX=1, OffY=1, Positive=1)}
</STYLE>

<SCRIPT language=javascript> 
var flowerarr=new Array();
var actname=new Array("睡觉","洗澡","吃饭","锻炼","看病","宰杀","放崽");
var fseed=0;
var theid=-1;
var flowerx=0;
var flowery=0;
var inittime;
var flag=true;
var thefirst=true;

function Flower(fc,ft,fa){
	var tmp=0;
	this.fclass=fc; this.ftype=ft; this.fact=fa;this.edit=true;
	if(fc==0){
		tmp=0;
	}else if((0<fc)&&(fc<25)){
		tmp=1;
	}else if((25<=fc)&&(fc<50)){
		tmp=2;
	}else if((50<=fc)&&(fc<75)){
		if(ft==0){tmp=30;}else if(ft==1){tmp=31;}else if(ft==2){tmp=32;}else if(ft==3){tmp=33;}else if(ft==4){tmp=34;}
	}else if((75<=fc)&&(fc<100)){
		if(ft==0){tmp=40;}else if(ft==1){tmp=41;}else if(ft==2){tmp=42;}else if(ft==3){tmp=43;}else if(ft==4){tmp=44;}
	}else if(fc==100){
		if(ft==0){tmp=50;}else if(ft==1){tmp=51;}else if(ft==2){tmp=52;}else if(ft==3){tmp=53;}else if(ft==4){tmp=54;}
	}
	this.fstyle=tmp;
}
for(var i=0;i<25;i++){	flowerarr[i]=new Flower(0,0,0);}
function gs(seednum){
	fseed=seednum;
}
function setMember(n){
	top.isMember=n;
}
function go(i,fc,ft,fa){
	//if(i>24||fc==''||fc==''||fc==''){
//	flowersact.innerHTML="网络有问题，请刷新页面";
//	return;
//	}
//	if(flowerarr[i].fclass==fc&&flowerarr[i].ftype==ft&&flowerarr[i].fact==fa)
//			flowerarr[i].edit=false;
//	else
			flowerarr[i]=new Flower(fc,ft,fa);
}
function gg(){
	flag=true;
	thefirst=true;
//	return;
	for(var i=0;i<flowerarr.length;i++){
//		if(flowerarr[i].edit==false) continue;
//		if(thefirst||flowerarr[i].fclass==0||flowerarr[i].fclass==1||flowerarr[i].fclass==25||flowerarr[i].fclass==50||flowerarr[i].fclass==75&&flowerarr[i].fclass==100){
			eval("f"+i).src="images/flower"+flowerarr[i].fstyle+".GIF";
			if(top.isMember==0&&i>8) eval("f"+i).style.filter="Alpha(Opacity=30)";//invert()
//		}
		var thefact=eval("mact"+i);
		if(top.isMember==1||i<9) thefact.src="images/a"+flowerarr[i].fact+".gif";
	}
	thefirst=false;
	flowersact.style.visibility='hidden';
	flowerseed.innerHTML=fseed+"颗。";
}
function getfshow(num){
	if(flowerarr.length<1) return;
if(flowerarr[num].fclass==0){flowersact.innerHTML="这里没有小猪，请放入猪崽!"}else{flowersact.innerHTML="小猪最新身体状况：成长值："+flowerarr[num].fclass+" 需要"+actname[(flowerarr[num].fact)];}
flowersact.style.visibility='visible';
}
function highLight(id,fx,fy) {
	if(top.isMember==0&&id>8) {
		flowersact.innerHTML="您的猪场是3X3大小，会员才有5X5的大猪场";
		return ;}
	if(flag){
		noneLight();
		theid=parseInt(id);
		flowerx=fx;
		flowery=fy-15;
		getfshow(id);
		var thef=eval("f"+id);
		thef.style.borderWidth="2px";
		thef.style.borderColor="#70AC58";
		thef.style.borderStyle="solid";
	}}
function noneLight() {
	for(i=0;i<25;i++){
		eval("f"+i).style.border="none";
	}
	showAct(0,0,0);
}
function showAct(type,fx,fy) {
	for(var i=1;i<5;i++){
		var theact=eval("actimg"+i);
		theact.style.pixelLeft=fx;
		theact.style.pixelTop=fy;
		theact.style.visibility="hidden";
		if(type==i) theact.style.visibility="visible";
	}
	linediv.style.visibility="hidden";
}
function act(actnum){
	if(actnum==9){
		if(confirm("您确定要重新建造猪场吗？\n重新建造后现在的小猪崽将全部饿死冻死，站长会再给你10只猪崽！\n点确定是建造，点取消是放弃。")){
			document.flowerout.actid.value=8;
			document.flowerout.flowerid.value=theid;
			document.flowerout.submit();
		}
	return;
	}
	if(theid<0) {alert("请选择要操作的位置!");return;}
	if(!flag) return;
	noneLight();
	if(actnum<5){
	if(flowerarr[theid].fclass==0){alert("此处没有小猪无法操作！");return;}
//alert("此处不存在鲜花无法操作!");return;
		flag=false;
		showAct(actnum,flowerx,flowery);
		inittime=setInterval("actShow()",1000);
		linediv.style.visibility='visible';
		document.flowerout.actid.value=actnum;
		document.flowerout.flowerid.value=theid;
	}if(actnum==7){
		if(flowerarr[theid].fclass!=0){alert("这里已经有小猪了不可以再放!");return;}
		if(confirm("你确定在此位置养猪吗？")) {
		document.flowerout.actid.value=7;
		document.flowerout.flowerid.value=theid;
		document.flowerout.submit();}
	}if(actnum==5){
		if(flowerarr[theid].fclass==0){alert("此处没有小猪无法操作!");return;}
		if(flowerarr[theid].fclass!=100){alert("此猪还没有成熟，不可以宰杀!");return;}
		if(confirm("您确定要宰杀小猪获得金卡等，还是将它换回3只小猪崽？\n点确定是宰杀，点取消是换小猪崽。")) {
			document.flowerout.actid.value=5;
			document.flowerout.flowerid.value=theid;
			document.flowerout.submit();
		}else{
			document.flowerout.actid.value=6;
			document.flowerout.flowerid.value=theid;
			document.flowerout.submit();
		}
	}
}
var ttt=0;
function actShow() {
	eval("tttd"+ttt).style.background="#ff0000";
	if(ttt==9){ttt=-1;actEnd();}
	ttt++;
}
function actEnd() {
	clearInterval(inittime);
	showAct(0,0,0);
	for(var i=0;i<10;i++){
		eval("tttd"+i).style.background="yellow";
	}
	document.flowerout.submit();
}
function changeImage(img, source) {
	if(!flag) return;
    img.src=source;
}
function doPreload(){
	var the_images = new Array('images/a1.gif','images/a2.gif','images/a3.gif','images/a4.gif','images/a5.gif','images/button1f.GIF','images/button2f.GIF','images/button3f.GIF','images/button4f.GIF','images/button5f.GIF','images/flower1.GIF','images/flower2.GIF','images/flower30.GIF','images/flower40.GIF','images/flower50.GIF','images/flower31.GIF','images/flower32.GIF','images/flower33.GIF','images/flower34.GIF','images/flower41.GIF','images/flower42.GIF','images/flower43.GIF','images/flower44.GIF','images/flower51.GIF','images/flower52.GIF','images/flower53.GIF','images/flower54.GIF');
	for(i= 0; i<the_images.length; i++){
   		var an_image = new Image();	
		an_image.src = the_images[i];	
	}
}

var dialogcount=16;
function count() {
	dialogboxbg.style.visibility="visible";
	dialogbox.style.visibility="visible";
	random=Math.random();
		if (random>0&&random<0.1&&dialogcount>0)
		dialogbox.innerHTML="第一次养猪要建造猪场的，以后就不用了！";
    if (random>0.1&&random<0.2&&dialogcount>0)
		dialogbox.innerHTML="选择一个放崽点，点放崽就把小猪放到场里了！";
    if (random>0.2&&random<0.3&&dialogcount>0)
		dialogbox.innerHTML="小猪的成长值100时宰杀，25、50、75时各长大一次！";
   if (random>0.3&&random<0.4&&dialogcount>0)
		dialogbox.innerHTML="现在猪场里有五个品种，有好斗猪、宠物猪、小山猪、皮卡猪、智力猪";
    if (random>0.4&&random<0.5&&dialogcount>0)
		dialogbox.innerHTML="刚开始长的差不多，长大后‘猪大十八变’哦……";
	if (random>=0.5&&random<0.58&&dialogcount>0)
		dialogbox.innerHTML="记住要按照小猪的需要来照顾它，否则它不理你的！";
    if (random>0.58&&random<0.7&&dialogcount>0)
		dialogbox.innerHTML="如果等级大于等于五级，就可以养满场猪了……";
    if (random>0.7&&random<0.8&&dialogcount>0)
		dialogbox.innerHTML="智力猪的品种较好，营养价值也高也！";
	if (random>=0.8&&random<0.88&&dialogcount>0)
		dialogbox.innerHTML="没有小猪崽可以重新建造猪场，站长免费送1只！";
	if (random>=0.88&&random<0.93&&dialogcount>0)
		dialogbox.innerHTML="小猪生长也需要睡觉的，等一下再来照顾它吧！";
    if (random>0.93&&random<1&&dialogcount>0)
		dialogbox.innerHTML="等小猪都长的白白胖胖又可爱时可以邀请朋友来参观，分享你的成果哦！";
	if (dialogcount<1)
		dialogbox.innerHTML="还有不明白的去看江湖帮助吧，那好详细哦！";	
        dialogcount--;
}
</SCRIPT>
<META content="Microsoft FrontPage 6.0" name=GENERATOR></HEAD>
<BODY bgColor=#70ac58 topMargin=0 onload=doPreload(); MARGINWIDTH="0" MARGINHEIGHT="0" >
<table width=100%><tr><td height=350> </td></tr><tr><td align=center><p align=center><font color="#FFFFFF"><a href="http://www.7758530.com" target=_blank><font color="#FF0000">爱情江湖网</font></a> 
        <font color="#000000">授权:<%=Application("aqjh_chatroomname")%></font></font><font style=font-size:9pt color=#000000></font></p></table>
<DIV id=field style="Z-INDEX: 3; LEFT: 160px; WIDTH: 367px; POSITION: absolute; TOP: 40px; HEIGHT: 253px">
<TABLE height=280 cellSpacing=0 cellPadding=0 width="100%" background=images/dw0.gif border=0>
  <TBODY>
  <TR>
    <TD width=7><IMG height=280 src="images/b2.gif" width=25></TD>
    <TD width="90%"></TD>
    <TD width=6> <DIV align=right><IMG height=280 src="images/b1.gif" width=25></DIV></TD>
  </TR>
  </TBODY>
</TABLE>
</DIV>
<DIV id=control style="Z-INDEX: 66; LEFT: 545px; WIDTH: 84px; POSITION: absolute; TOP: 60px; HEIGHT: 206px">
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD width="50%" height=44><IMG onmouseup="changeImage(button5, 'images/button5.GIF');" onmousedown="" id=button5 title=放崽 onclick="" height=44 src="images/button5.gif"  width=42></TD>
    <TD height=44><IMG onmouseup="changeImage(button4, 'images/button4.GIF');" onmousedown="" id=button4 title=锻炼 onclick="" height=44 src="images/button4.gif"  width=42></TD>
  </TR>
  <TR>
    <TD height=44><IMG onmouseup="changeImage(button3, 'images/button3.GIF');" onmousedown="" id=button3 title=沐浴 onclick="" height=44 src="images/button3.gif" width=42></TD>
    <TD height=44><IMG onmouseup="changeImage(button2, 'images/button2.GIF');" onmousedown="" id=button2 title=吃饭 onclick="" height=44 src="images/button2.gif" width=42></TD>
  </TR>
  <TR>
    <TD height=44><IMG onmouseup="changeImage(button1, 'images/button1.GIF');" onmousedown="" id=button1 title=看病 onclick="" height=44 src="images/button1.gif" width=42></TD>
    <TD height=44><IMG onmouseup="changeImage(button6, 'images/button6.GIF');" onmousedown="" id=button6 title=宰杀 onclick="" height=44 src="images/button6.gif" width=42></TD>
   </TR>
  <TR>
    <TD colSpan=2 height=44> <DIV align=center><%if Session("aqjh_name")=name then%><A onmouseup="changeImage('button7', 'images/button7.gif');" onmousedown="" title=第一次养猪必须翻新猪场,如果里面已经有小猪，小猪崽将全部饿死冻死。请考虑清楚~~。 onclick=""  href="#" target=_top><IMG id=button7 height=44 alt="开 垦 猪 场" src="images/button7.gif"  width=55 border=0></A><%end if%></DIV> </TD>
   </TR>
   </TBODY>
   </TABLE>
   </DIV>
    <!-- flower img -->
    <IMG id=f0 style="Z-INDEX: 28; FILTER: ; LEFT: 193px; WIDTH: 50px; POSITION: absolute; TOP: 49px; HEIGHT: 50px" onmousedown=" highLight('0',193,36);" src="images/flower0.gif"> 
    <IMG id=f1 style="Z-INDEX: 4; FILTER: ; LEFT: 255px; WIDTH: 50px; POSITION: absolute; TOP: 49px; HEIGHT: 50px" onmousedown=" highLight('1',255,36);" src="images/flower0.gif">   
    <IMG id=f2 style="Z-INDEX: 5; FILTER: ; LEFT: 319px; WIDTH: 50px; POSITION: absolute; TOP: 49px; HEIGHT: 50px" onmousedown=" highLight('2',319,36);" src="images/flower0.gif">   
    <IMG id=f3 style="Z-INDEX: 6; FILTER: ; LEFT: 193px; WIDTH: 50px; POSITION: absolute; TOP: 102px; HEIGHT: 50px" onmousedown=" highLight('3',193,100);" src="images/flower0.gif">   
    <IMG id=f4 style="Z-INDEX: 7; FILTER: ; LEFT: 255px; WIDTH: 50px; POSITION: absolute; TOP: 102px; HEIGHT: 50px" onmousedown=" highLight('4',255,100);" src="images/flower0.gif">   
    <IMG id=f5 style="Z-INDEX: 8; FILTER: ; LEFT: 319px; WIDTH: 50px; POSITION: absolute; TOP: 102px; HEIGHT: 50px" onmousedown=" highLight('5',319,100);" src="images/flower0.gif">   
    <IMG id=f6 style="Z-INDEX: 9; FILTER: ; LEFT: 193px; WIDTH: 50px; POSITION: absolute; TOP: 155px; HEIGHT: 50px" onmousedown=" highLight('6',193,147);" src="images/flower0.gif">   
    <IMG id=f7 style="Z-INDEX: 10; FILTER: ; LEFT: 255px; WIDTH: 50px; POSITION: absolute; TOP: 155px; HEIGHT: 50px" onmousedown=" highLight('7',255,147);" src="images/flower0.gif">   
    <IMG id=f8 style="Z-INDEX: 11; FILTER: ; LEFT: 319px; WIDTH: 50px; POSITION: absolute; TOP: 155px; HEIGHT: 50px" onmousedown=" highLight('8',319,147);" src="images/flower0.gif">   
    <!-- for member to grow flower -->  
    <IMG onmousedown=" highLight('9',193,200);" id=f9 style="Z-INDEX: 12; FILTER: ; LEFT: 193px; WIDTH: 50px; POSITION: absolute; TOP: 208px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('10',193,250);" id=f10 style="Z-INDEX: 13; FILTER: ; LEFT: 193px; WIDTH: 50px; POSITION: absolute; TOP: 261px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('11',255,200);" id=f11 style="Z-INDEX: 14; FILTER: ; LEFT: 255px; WIDTH: 50px; POSITION: absolute; TOP: 208px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('12',255,250);" id=f12 style="Z-INDEX: 15; FILTER: ; LEFT: 255px; WIDTH: 50px; POSITION: absolute; TOP: 261px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('13',319,200);" id=f13 style="Z-INDEX: 16; FILTER: ; LEFT: 319px; WIDTH: 50px; POSITION: absolute; TOP: 208px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('14',319,250);" id=f14 style="Z-INDEX: 17; FILTER: ; LEFT: 319px; WIDTH: 50px; POSITION: absolute; TOP: 261px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('15',383,36);" id=f15 style="Z-INDEX: 18; FILTER: ; LEFT: 383px; WIDTH: 50px; POSITION: absolute; TOP: 49px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('16',383,100);" id=f16 style="Z-INDEX: 19; FILTER: ; LEFT: 383px; WIDTH: 50px; POSITION: absolute; TOP: 102px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('17',383,147);" id=f17 style="Z-INDEX: 20; FILTER: ; LEFT: 383px; WIDTH: 50px; POSITION: absolute; TOP: 155px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('18',383,200);" id=f18 style="Z-INDEX: 21; FILTER: ; LEFT: 383px; WIDTH: 50px; POSITION: absolute; TOP: 208px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('19',383,250);" id=f19 style="Z-INDEX: 22; FILTER: ; LEFT: 383px; WIDTH: 50px; POSITION: absolute; TOP: 261px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('20',447,36);" id=f20 style="Z-INDEX: 27; FILTER: ; LEFT: 447px; WIDTH: 50px; POSITION: absolute; TOP: 49px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('21',447,100);" id=f21 style="Z-INDEX: 23; FILTER: ; LEFT: 447px; WIDTH: 50px; POSITION: absolute; TOP: 102px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('22',447,147);" id=f22 style="Z-INDEX: 24; FILTER: ; LEFT: 447px; WIDTH: 50px; POSITION: absolute; TOP: 155px; HEIGHT: 50px" src="images/flower0.gif" width=50>   
    <IMG onmousedown=" highLight('23',447,200);" id=f23 style="Z-INDEX: 25; FILTER: ; LEFT: 447px; WIDTH: 50px; POSITION: absolute; TOP: 208px; HEIGHT: 50px" src="images/flower0.gif">   
    <IMG onmousedown=" highLight('24',447,250);" id=f24 style="Z-INDEX: 26; FILTER: ; LEFT: 447px; WIDTH: 50px; POSITION: absolute; TOP: 261px; HEIGHT: 50px" src="images/flower0.gif">   
    <!-- mact img -->  
    <IMG id=mact0 style="Z-INDEX: 37; LEFT: 238px; POSITION: absolute; TOP: 84px" src="images/a0.gif">   
    <IMG id=mact1 style="Z-INDEX: 38; LEFT: 301px; POSITION: absolute; TOP: 84px" src="images/a0.gif">   
    <IMG id=mact2 style="Z-INDEX: 39; LEFT: 364px; POSITION: absolute; TOP: 85px" src="images/a0.gif">   
    <IMG id=mact3 style="Z-INDEX: 40; LEFT: 238px; POSITION: absolute; TOP: 138px" src="images/a0.gif">   
    <IMG id=mact4 style="Z-INDEX: 41; LEFT: 301px; POSITION: absolute; TOP: 138px" src="images/a0.gif">   
    <IMG id=mact5 style="Z-INDEX: 42; LEFT: 364px; POSITION: absolute; TOP: 138px" src="images/a0.gif">   
    <IMG id=mact6 style="Z-INDEX: 43; LEFT: 238px; POSITION: absolute; TOP: 191px" src="images/a0.gif">   
    <IMG id=mact7 style="Z-INDEX: 44; LEFT: 301px; POSITION: absolute; TOP: 191px" src="images/a0.gif">   
    <IMG id=mact8 style="Z-INDEX: 45; LEFT: 364px; POSITION: absolute; TOP: 191px" src="images/a0.gif">   
    <IMG id=mact9 style="Z-INDEX: 46; LEFT: 238px; POSITION: absolute; TOP: 244px" src="images/a0.gif">   
    <IMG id=mact10 style="Z-INDEX: 47; LEFT: 238px; POSITION: absolute; TOP: 297px" src="images/a0.gif">   
    <IMG id=mact11 style="Z-INDEX: 48; LEFT: 301px; POSITION: absolute; TOP: 244px" src="images/a0.gif">   
    <IMG id=mact12 style="Z-INDEX: 49; LEFT: 301px; POSITION: absolute; TOP: 297px" src="images/a0.gif">   
    <IMG id=mact13 style="Z-INDEX: 50; LEFT: 364px; POSITION: absolute; TOP: 244px" src="images/a0.gif">   
    <IMG id=mact14 style="Z-INDEX: 51; LEFT: 364px; POSITION: absolute; TOP: 297px" src="images/a0.gif">   
    <IMG id=mact15 style="Z-INDEX: 52; LEFT: 428px; POSITION: absolute; TOP: 84px" src="images/a0.gif">   
    <IMG id=mact16 style="Z-INDEX: 53; LEFT: 428px; POSITION: absolute; TOP: 138px" src="images/a0.gif">   
    <IMG id=mact17 style="Z-INDEX: 54; LEFT: 428px; POSITION: absolute; TOP: 191px" src="images/a0.gif">   
    <IMG id=mact18 style="Z-INDEX: 55; LEFT: 428px; POSITION: absolute; TOP: 244px" src="images/a0.gif">   
    <IMG id=mact19 style="Z-INDEX: 56; LEFT: 428px; POSITION: absolute; TOP: 297px" src="images/a0.gif">   
    <IMG id=mact20 style="Z-INDEX: 57; LEFT: 492px; POSITION: absolute; TOP: 84px" src="images/a0.gif">   
    <IMG id=mact21 style="Z-INDEX: 58; LEFT: 492px; POSITION: absolute; TOP: 138px" src="images/a0.gif">   
    <IMG id=mact22 style="Z-INDEX: 59; LEFT: 492px; POSITION: absolute; TOP: 191px" src="images/a0.gif">   
    <IMG id=mact23 style="Z-INDEX: 60; LEFT: 492px; POSITION: absolute; TOP: 244px" src="images/a0.gif">   
    <IMG id=mact24 style="Z-INDEX: 61; LEFT: 492px; POSITION: absolute; TOP: 297px" src="images/a0.gif">   
    <!-- act img -->  
    <IMG id=actimg1 style="Z-INDEX: 64; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 0px" src="images/act1.gif">   
    <IMG id=actimg2 style="Z-INDEX: 65; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 0px" src="images/act2.gif">   
    <IMG id=actimg3 style="Z-INDEX: 66; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 0px" src="images/act3.gif">   
    <IMG id=actimg4 style="Z-INDEX: 67; LEFT: 0px; VISIBILITY: hidden; POSITION: absolute; TOP: 0px" src="images/act4.gif">   
    <!-- ps line -->  
    <DIV id=linediv style="Z-INDEX: 72; LEFT: 530px; VISIBILITY: hidden; WIDTH: 90px; POSITION: absolute; TOP: 270px; HEIGHT: 9px" align=center><FONT color=green><B>小猪活动状态::</B></FONT>
    <TABLE style="FONT-WEIGHT: normal; FONT-SIZE: 2pt; BACKGROUND: yellow; WIDTH: 70pt; LINE-HEIGHT: normal; FONT-STYLE: normal; HEIGHT: 2pt; FONT-VARIANT: normal" cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <TR>
    <TD id=tttd0></TD>
    <TD id=tttd1></TD>
    <TD id=tttd2></TD>
    <TD id=tttd3></TD>
    <TD id=tttd4></TD>
    <TD id=tttd5></TD>
    <TD id=tttd6></TD>
    <TD id=tttd7></TD>
    <TD id=tttd8></TD>
    <TD id=tttd9></TD>
    </TR>
    </TBODY>
    </TABLE>
    </DIV>
<DIV id=flowersact style="Z-INDEX: 63; LEFT: 200px; VISIBILITY: hidden; WIDTH: 350px; POSITION: absolute; TOP: 330px">小猪身体情况：0。需要睡觉</DIV>
<IMG id=dialogboxbg style="Z-INDEX: 65; LEFT: 23px; WIDTH: 120px; POSITION: absolute; TOP: 219px; HEIGHT: 100px" src="images/bg10.gif"> 
<DIV id=Layer1 style="Z-INDEX: 2; LEFT: 3px; WIDTH: 600px; POSITION: absolute; TOP: 30px; HEIGHT: 270px">
<DIV id=dialogbox style="Z-INDEX: 15; LEFT: 32px; WIDTH: 100px; POSITION: absolute; TOP: 234px; HEIGHT: 70px">我是猪场的小管家，有什么不懂的请点击我。</DIV>
<IMG id=photoseller onmouseover=count() style="Z-INDEX: 14; LEFT: 15px; WIDTH: 120px; POSITION: absolute; TOP: 132px; HEIGHT: 107px" onclick=count() src="images/help.gif" ;> 
<DIV id=formdiv style="DISPLAY: none; Z-INDEX: -100; LEFT: 0px; WIDTH: 1px; POSITION: absolute; TOP: 0px; HEIGHT: 1px">
<FORM name=flowerout action=show.asp method=post target=fg>
	<INPUT type=hidden name=actid value=-1> 
	<INPUT type=hidden name=flowerid value=-1> 
</FORM>
</DIV>
<DIV id=frdiv style="DISPLAY: none; Z-INDEX: -100; LEFT: 0px; WIDTH: 100px; POSITION: absolute; TOP: 0px; HEIGHT: 100px">
	<IFRAME name=fg marginWidth=100 marginHeight=100 src="" frameBorder=NO noResize width=100 scrolling=no height=20>
	</IFRAME>
</DIV>
<TABLE height=280 cellSpacing=0 cellPadding=0 width=620 align=left border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width="25%">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD><IMG src="images/mimi.gif"></TD></TR>
        <TR>
          <TD>　</TD></TR>
        <TR>
          <TD>
            <TABLE cellSpacing=0 cellPadding=2 border=0>
              <TBODY>
              	<TR>
                	<TD>{<font color=cc0000><%=name%></font>}拥有</TD>
	                <TD><IMG height=20 src="images/02.gif" width=20></TD>
	                <TD  id=flowerseed>00小猪崽。</TD>
                </TR>
              </TBODY>
            </TABLE>
          </TD>
        </TR>
      </TBODY>
     </TABLE></TD></TR></TBODY></TABLE>
</DIV>
</BODY></HTML>
<%
Set conn = Server.CreateObject("ADODB.CONNECTION")
Set rs = Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
aqjh_grade=Session("aqjh_grade")
rs.open "select 会员,会员等级,hua from 用户 where 姓名='"&name&"'",conn,2,2
if rs("hua")<>"" then
else
	conn.execute "update 用户 set hua='3|"&now()&"|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|' where 姓名='" & name &"'"
	Response.Write "<script language=JavaScript>{alert('提示：您第一次来到江湖猪场，站长送你3只小猪崽，好好养哦！');}</script>"
end if

if rs("会员") or rs("会员等级") then
	Response.Write "<script language=JavaScript>{parent.setMember(1);}</script>"
else
	Response.Write "<script language=JavaScript>{parent.setMember(0);}</script>"
end if
fhuasj=split(rs("hua"),"|")
seedsm=fhuasj(0)
hua_sj=fhuasj(1)
for xyz=0 to 99 step 1
randomize timer
if rs("会员") or rs("会员等级") then
czhuad=int(rnd*25)+2
else
czhuad=int(rnd*9)+2
end if
czhuasj=split(fhuasj(czhuad),";")
if czhuasj(0)>=1 then
 exit for
end if
next
for i=2 to 26 step 1
huasj=split(fhuasj(i),";")
huasj_cs=huasj(0)
huasj_zl=huasj(1)
huasj_cz=huasj(2)
if huasj_cs>=100 then
 huasj_cs=100
 huasj_cz=5
end if
if i=czhuad and huasj_cz=0 and huasj_cs<>0 and DateDiff("s",hua_sj,now())>=120 then
hua_sj=now()
randomize timer
czfs=int(rnd*4)+1
huasj_cz=czfs
huasjnew=huasj_cs&";"&huasj_zl&";"&huasj_cz
sfcz=1
else
huasjnew=huasj_cs&";"&huasj_zl&";"&huasj_cz
end if
Response.Write "<script language=JavaScript>{parent.go("&i-2&","&huasj_cs&","&huasj_zl&","&huasj_cz&");}</script>"
ghuasj=ghuasj&huasjnew&"|"
next
ghuasj=seedsm&"|"&hua_sj&"|"&ghuasj
conn.execute "update 用户 set hua='"&ghuasj&"' where 姓名='" & name &"'"
Response.Write "<script language=JavaScript>{parent.gs("&seedsm&");parent.gg();}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>