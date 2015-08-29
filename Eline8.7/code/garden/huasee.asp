<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
toname=LCase(trim(Request.QueryString("toname")))
if toname="大家" then toname=sjjh_name
%>
<HTML><HEAD><TITLE>E线江湖-神秘花园---<%=toname%>的小花园</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
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
</STYLE>

<SCRIPT language=javascript> 
var flowerarr=new Array();
var actname=new Array("休息","浇水","施肥","剪枝","除虫","收获","播种");
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
		if(top.isMember==1||i<8) thefact.src="images/a"+flowerarr[i].fact+".gif";
	}
	thefirst=false;
	flowersact.style.visibility='hidden';
	flowerseed.innerHTML=fseed+"颗。";
}
function getfshow(num){
	if(flowerarr.length<1) return;
if(flowerarr[num].fclass==0){flowersact.innerHTML="此处无花，请播种!"}else{flowersact.innerHTML="成熟度："+flowerarr[num].fclass+" 需要"+actname[(flowerarr[num].fact)];}
flowersact.style.visibility='visible';
}
function highLight(id,fx,fy) {
	if(top.isMember==0&&id>8) {
		flowersact.innerHTML="您的花园是3X3大小，会员才有5X5的大花园";
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
	alert("这里是查看花看花园\n你无权操作!");return;
	if(actnum==9){
		if(confirm("您确定要将花园重新开垦吗？\n点确定是开垦，点取消是放弃。")){
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
	if(flowerarr[theid].fclass==0){alert("此处不存在鲜花无法操作!");return;}
//alert("此处不存在鲜花无法操作!");return;
		flag=false;
		showAct(actnum,flowerx,flowery);
		inittime=setInterval("actShow()",1000);
		linediv.style.visibility='visible';
		document.flowerout.actid.value=actnum;
		document.flowerout.flowerid.value=theid;
	}if(actnum==7){
		if(flowerarr[theid].fclass!=0){alert("这里已经有花了不可以再种!");return;}
		if(confirm("你确定在此位置种花吗？")) {
		document.flowerout.actid.value=7;
		document.flowerout.flowerid.value=theid;
		document.flowerout.submit();}
	}if(actnum==5){
		if(flowerarr[theid].fclass==0){alert("此处不存在鲜花无法操作!");return;}
		if(flowerarr[theid].fclass!=100){alert("此花还没有成熟，不可以采摘!");return;}
		if(confirm("您确定要将花采下收藏，还是将花变成3颗种子？\n点确定是采下，点取消是变种子。")) {
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
		dialogbox.innerHTML="第一次种花要开垦花园的，以后就不用了！";
    if (random>0.1&&random<0.2&&dialogcount>0)
		dialogbox.innerHTML="选择一个种植点，点播种就把种子种下了！";
    if (random>0.2&&random<0.3&&dialogcount>0)
		dialogbox.innerHTML="花儿的成熟度100时收获，25、50、75时各长大一次！";
   if (random>0.3&&random<0.4&&dialogcount>0)
		dialogbox.innerHTML="现在部落有五种颜色的花，紫色的逸韵最珍贵哦！";
    if (random>0.4&&random<0.5&&dialogcount>0)
		dialogbox.innerHTML="不是紫色的花收获的种子一定长成紫色的花，其实应该是……";
	if (random>=0.5&&random<0.58&&dialogcount>0)
		dialogbox.innerHTML="记住要按照花儿的需要来浇水施肥，否则没效果哦！";
    if (random>0.58&&random<0.7&&dialogcount>0)
		dialogbox.innerHTML="如果是会员，就可以种满整个花园的花，……";
    if (random>0.7&&random<0.8&&dialogcount>0)
		dialogbox.innerHTML="一定要保护好密码哦！要是被别人开垦了花园花就都没了，呜呜！";
	if (random>=0.8&&random<0.88&&dialogcount>0)
		dialogbox.innerHTML="忘了留种子可以重新开垦一下，部落免费送一颗！";
	if (random>=0.88&&random<0.93&&dialogcount>0)
		dialogbox.innerHTML="花儿生长也需要休息的，等一下再来照顾它吧！";
    if (random>0.93&&random<1&&dialogcount>0)
		dialogbox.innerHTML="要是满园花儿都开了记得邀请朋友来参观，分享你的成果哦！";
	if (dialogcount<1)
		dialogbox.innerHTML="还有不明白的去看部落帮助吧，那好详细哦！";	
        dialogcount--;
}
</SCRIPT>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY bgColor=#70ac58 topMargin=0 onload=doPreload(); MARGINWIDTH="0" MARGINHEIGHT="0" >
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
    <TD width="50%" height=44><IMG onmouseup="changeImage(button5, 'images/button5.GIF');" onmousedown="changeImage(button5, 'images/button5f.GIF');" id=button5 title=播种 onclick="act('7');" height=44 src="images/button5.gif"  width=42></TD>
    <TD height=44><IMG onmouseup="changeImage(button4, 'images/button4.GIF');" onmousedown="changeImage(button4, 'images/button4f.GIF');" id=button4 title=剪枝 onclick="act('3');" height=44 src="images/button4.gif"  width=42></TD>
  </TR>
  <TR>
    <TD height=44><IMG onmouseup="changeImage(button3, 'images/button3.GIF');" onmousedown="changeImage(button3, 'images/button3f.GIF');" id=button3 title=浇水 onclick="act('1');" height=44 src="images/button3.gif" width=42></TD>
    <TD height=44><IMG onmouseup="changeImage(button2, 'images/button2.GIF');" onmousedown="changeImage(button2, 'images/button2f.GIF');" id=button2 title=施肥 onclick="act('2');" height=44 src="images/button2.gif" width=42></TD>
  </TR>
  <TR>
    <TD height=44><IMG onmouseup="changeImage(button1, 'images/button1.GIF');" onmousedown="changeImage(button1, 'images/button1f.GIF');" id=button1 title=除虫 onclick="act('4');" height=44 src="images/button1.gif" width=42></TD>
    <TD height=44><IMG onmouseup="changeImage(button6, 'images/button6.GIF');" onmousedown="changeImage(button6, 'images/button6f.GIF');" id=button6 title=收获 onclick="act('5');" height=44 src="images/button6.gif" width=42></TD>
   </TR>
  <TR>
    <TD colSpan=2 height=44> <DIV align=center><A onmouseup="changeImage('button7', 'images/button7.gif');" onmousedown="changeImage('button7', 'images/button7f.gif');" title=第一次种花必须开垦花园,如果已经种了花，就会将你的花园重新犁一遍，现在种的花都会犁掉。 onclick="act('9');return false;"  href="#" target=_top><IMG id=button7 height=44 alt="开 垦 花 园" src="images/button7.gif"  width=55 border=0></A></DIV> </TD>
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
    <DIV id=linediv style="Z-INDEX: 72; LEFT: 530px; VISIBILITY: hidden; WIDTH: 90px; POSITION: absolute; TOP: 270px; HEIGHT: 9px" align=center><FONT color=green><B>动作:</B></FONT>
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
<DIV id=flowersact style="Z-INDEX: 63; LEFT: 200px; VISIBILITY: hidden; WIDTH: 350px; POSITION: absolute; TOP: 330px">成熟度：0。需要休息</DIV>
<IMG id=dialogboxbg style="Z-INDEX: 65; LEFT: 23px; WIDTH: 120px; POSITION: absolute; TOP: 219px; HEIGHT: 100px" src="images/bg10.gif"> 
<DIV id=Layer1 style="Z-INDEX: 2; LEFT: 3px; WIDTH: 600px; POSITION: absolute; TOP: 30px; HEIGHT: 270px">
<DIV id=dialogbox style="Z-INDEX: 15; LEFT: 32px; WIDTH: 100px; POSITION: absolute; TOP: 234px; HEIGHT: 70px">我是花园小园丁，有什么不懂的请点击我。</DIV>
<IMG id=photoseller onmouseover=count() style="Z-INDEX: 14; LEFT: 15px; WIDTH: 120px; POSITION: absolute; TOP: 132px; HEIGHT: 107px" onclick=count() src="images/help.gif" ;> 
<DIV id=formdiv style="DISPLAY: none; Z-INDEX: -100; LEFT: 0px; WIDTH: 1px; POSITION: absolute; TOP: 0px; HEIGHT: 1px">
<FORM name=flowerout action=see.asp?toname=<%=toname%> method=post target=fg>
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
          <TD>&nbsp;</TD></TR>
        <TR>
          <TD>
            <TABLE cellSpacing=0 cellPadding=2 border=0>
              <TBODY>
              	<TR>
                	<TD>[<%=toname%>]拥有</TD>
	                <TD><IMG height=20 src="images/02.gif" width=20></TD>
	                <TD  id=flowerseed>00颗。</TD>
                </TR>
              </TBODY>
            </TABLE>
          </TD>
        </TR>
      </TBODY>
     </TABLE></TD></TR></TBODY></TABLE></DIV>

</BODY></HTML>
