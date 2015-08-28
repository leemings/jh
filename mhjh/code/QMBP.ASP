
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>七门八派</title>
<script language="JavaScript">
<!--
var zudza,myname
function MyShowMenu(menuid,x,y,np)
{
//Layer2.style.visibility='hidden';
document.getElementById(menuid).style.visibility='visible';
document.getElementById(menuid).style.left=x+20+'px';
document.getElementById(menuid).style.top=y+'px'
document.mainmove.nowpos.value=np;
//Layer2.style.visibility='visible';
return true;
}
function MyHideMenu(menuid)
{
//Layer1.style.visibility='hidden';
document.getElementById(menuid).style.visibility='hidden';
//document.getElementById('zuza').style.disabled=false;
//Layer1.style.visibility='visible';
return true;
}
function MySetXY(menuid,x,y)
{
var obj=document.getElementById(menuid)
if(obj!=null){
obj.style.left=x+'px';
obj.style.top=y+'px'}
return true;
}
//-->
</script>

<script language="JavaScript">
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("欢迎光临魔幻皇城，点击功能按钮，可以进入另外的天地，完全模拟现实构景，如果你找不到哪个功能，那你就想一想这个功能应该属于哪一类的，就看看相关的功能吧，比如，我想离婚，那当然就要去婚介所了！","侠客提示")){myopenwindow('http://www.yx8.net/','wupin','scrollbars=yes,resizable=yes,width=800,height=600');}}}
</script>
<script language="JavaScript">
function guanfuyamen(wpath,n,wtype){
woiwo=window.open(wpath,n,wtype);
startnodo();
//woiwo.moveTo(x,y);
}
</script>
</head>
<body topmargin="0" leftmargin="0"  background="qmbp.jpg" oncontextmenu="return false;">
<div id="pm" style="position:absolute; left:400px; top:10px; width:40px; height:40px;z-index:2;visibility:visible"> 
    
  <div id="player0" style="position: absolute; left: 0; top: 0; width: 108; height: 65; z-index: 2; visibility: visible"> 
    <img width="80" height="69" src=adv/images/004.gif> </div>
    <div id="player1" style="position:absolute; left:0px; top:0px; width:40px; height:40px;z-index:2;visibility:hidden"> 
      <img width="40" height="40" src=adv/images/004.gif>
    </div>
    <div id="player2" style="position:absolute; left:0px; top:0px; width:40px; height:40px;z-index:2;visibility:hidden"> 
      <img width="40" height="40" src=adv/images/004.gif>
    </div>
    <div id="player3" style="position:absolute; left:0px; top:0px; width:40px; height:40px;z-index:2;visibility:hidden"> 
      <img width="40" height="40" src=adv/images/004.gif>
    </div>
</div>
<form method=POST  name=mainmove>
  <div id="Layer1" style="position:absolute; left:5px; top:0px; width:571px; height:3px; z-index:1"> 
	<input name=nowpos value='' type=hidden>
	<div style="position:absolute;left:300px;top:0px;z-index:1">
    <input type="hidden" name="oldx" size="3" maxlength="3" readonly="true">
    <input name="oldy" type="hidden" size="3" maxlength="3" readonly="true">
	</div>  
  </div> 


<script language="Javascript">
var tmid;
var xp,yp,czmc,rt,pstep,curd,offx,offy; //,ismoving
if(screen.height==0){
offx=0;offy=0;}
else{
offx=0;offy=0;

}

xp=offx;yp=offy;pstep=0;
curd=0;//ismoving=0;
var ox,oy;
ox=xp;oy=yp;
function changedes(fx)
{
   document.getElementById('player'+fx).style.visibility='visible';
   for(var i=0;i<4;i++)
   {
      if(i!=fx)
        document.getElementById('player'+i).style.visibility='hidden';
   }
   
}
function MovePlayer(tx,ty)
{
//alert('ok');
pstep++;
document.mainmove.oldx.value=xp;
document.mainmove.oldy.value=yp;  

//if (ismoving>1)
//{   
//   ismoving=0;
   //return true;
//}

ox=tx;
oy=ty;

if (xp<=ox&&(((oy-yp)<(ox-xp)&&oy>=yp)||(-(oy-yp)<(ox-xp)&&oy<yp))) //abs
{
if(oy>yp)
   curd=3;
else
   curd=3;
changedes(curd);
}
else if(xp<=ox&&(((oy-yp)>(ox-xp)&&oy>=yp)||(-(oy-yp)>(ox-xp)&&oy<yp)))
{
if(oy>yp)
   curd=0;
else if(oy<yp)
   curd=2;
else
   curd=1;
changedes(curd);
}
else if(xp>ox&&(((oy-yp)<(ox-xp)&&oy>=yp)||(-(oy-yp)<(ox-xp)&&oy<yp)))
{
if(oy>yp)
   curd=1;
else
   curd=1;
changedes(curd);
}
else if(xp>ox&&(((oy-yp)>(ox-xp)&&oy>=yp)||(-(oy-yp)>(ox-xp)&&oy<yp)))
{
if(oy>yp)
   curd=0;
else if(oy<yp)
   curd=2;
else
   curd=1;
changedes(curd);
}

if(xp<ox)
   xp+=1;
else if(xp>ox)
   xp-=1;
if(yp<oy)
   yp+=1;
else if(yp>oy)
   yp-=1; 

if(xp==ox&&yp==oy)
{
   if(czmc!='')
   {
      if(czmc=='chat')
	     chat();
	  else if(czmc.indexOf('/')==-1)
	     location.href=czmc;
	  else if(czmc.indexOf('adv/')==0)
	     location.href=czmc;
	  else
         window.open(czmc,"魔幻场景","width=640,height=480,left='+(screen.700)+',top='+(screen.450)+',resizable=no,scrollbars=yes,toolbar=no,menubar=no");//MyShowMenu('CD',ox,oy,czmc);
   }
   else
     {ox+=400;oy+=400;}
   
   rt=0;
   //ismoving=0;
   document.mainmove.oldx.value=ox;
   document.mainmove.oldy.value=oy;   
}

document.getElementById('pm').style.left=xp+'px';
document.getElementById('pm').style.top=yp+'px';
if(tmid!=null)
{
  clearTimeout(tmid);
  tmid=null;
}
if(rt>0)
   tmid=setTimeout('MovePlayer('+ox+','+oy+')',rt);
return true;
}

  </script>

<script language='Javascript'>xp=offx+document.body.scrollLeft+345;yp=offy+document.body.scrollTop+140;document.mainmove.oldx.value=xp;document.mainmove.oldy.value=yp;document.getElementById('pm').style.left=xp+'px';document.getElementById('pm').style.top=yp+'px';</script>
<div id='bingqi' style='position:absolute; left:52px; top:28px; width: 56px; height: 24px;'>
    <input onClick='javascript:rt=30;czmc="mp/Corpman1.asp";MovePlayer(235+offx,370+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='同门弟子' type='button' name=sy2 readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('bingqi',235+offx,370+offy);</script>
  <div id='fangju' style='position:absolute;left:289px; top:267px;'></div>
  <script language='Javascript'>MySetXY('fangju',540+offx,165+offy);</script>
<div id='daoju' style='position:absolute;left:237px; top:283px;'>
    <input onclick='javascript:rt=30;czmc="mp/jk.asp";MovePlayer(285+offx,315+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='弟子工资' type='button' name=sy readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('daoju',285+offx,315+offy);</script>
<div id='biaoju' style='position:absolute;left:141px; top:170px;'>
    <input onclick='javascript:rt=30;czmc="mp/huansu.asp";MovePlayer(35+offx,315+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='出家还俗' type='button' name=sy readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('biaoju',35+offx,315+offy);</script>
  <div id='kezan' style='position:absolute;left:375px; top:178px;'> </div>
  <script language='Javascript'>MySetXY('kezan',537+offx,383+offy);</script>
  <div id='cufang' style='position:absolute;left:160px; top:156px;'></div>
  <script language='Javascript'>MySetXY('cufang',620+offx,190+offy);</script>
<div id='ducang' style='position:absolute;left:173px; top:142px;'>
    <input onclick='javascript:rt=30;czmc="mp/Corpatt.asp";MovePlayer(170+offx,80+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='武功设置' type='button' name=sy readonly size=6 maxlength='8' >
     </div><script language='Javascript'>MySetXY('ducang',170+offx,80+offy);</script>
<div id='gaoligui' style='position:absolute;left:218px; top:453px;'>
    <input onclick='javascript:rt=30;czmc="mp/zm.asp";MovePlayer(382+offx,270+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='掌门工资' type='button' name=sy readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('gaoligui',382+offx,270+offy);</script>
  <div id='diedapi' style='position:absolute;left:83px; top:208px;'></div>
  <script language='Javascript'>MySetXY('diedapi',70+offx,150+offy);</script>
  <div id='yaopu' style='position:absolute;left:175px; top:271px;'> </div>
  <script language='Javascript'>MySetXY('yaopu',510+offx,330+offy);</script>
<div id='sasoutang' style='position:absolute;left:115px; top:278px;'>
    <input onclick='javascript:rt=30;czmc="11.htm";MovePlayer(300+offx,200+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='返回皇城' type='button' name=sy readonly size=6 maxlength='8' >
     </div><script language='Javascript'>MySetXY('sasoutang',300+offx,200+offy);</script>
  <div id='pifa' style='position:absolute;left:245px; top:117px;'></div>
  <script language='Javascript'>MySetXY('pifa',600+offx,340+offy);</script>
<div id='ersou' style='position:absolute;left:201px; top:136px;'>
    <input onclick='javascript:rt=30;czmc="mp/menpai.asp";MovePlayer(125+offx,248+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='门派列表' type='button' name=sy readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('ersou',125+offx,248+offy);</script>
<div id='jiaowai' style='position:absolute;left:483px; top:128px;'>
    <input onclick='javascript:rt=30;czmc="mp/mpjianjie.ASP";MovePlayer(490+offx,100+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='掌门公告' type='button' name=sy readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('jiaowai',490+offx,100+offy);</script>
<div id='adv' style='position:absolute;left:301px; top:51px;'>
    <input onclick='javascript:rt=30;czmc="mp/corpman.asp?corp=";MovePlayer(300+offx,120+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='掌门册封' type='button' name=sy readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('adv',300+offx,120+offy);</script>
<div id='ct' style='position:absolute;left:449px; top:391px;'>
    <input onclick='javascript:rt=30;czmc="mp/Addcorp.asp";MovePlayer(340+offx,26+offy);' onMouseOut="this.style.backgroundColor='#800000';this.style.color='#FFFFFF';" onMouseOver="this.style.backgroundColor='#FFDBAC';this.style.color='#0000FF';" style='background-color: #800000; color: #FFFFFF; border-style: solid; border-width: 1' value='自立门户' type='button' name=sy readonly size=6 maxlength='8' >
  </div><script language='Javascript'>MySetXY('ct',340+offx,26+offy);</script>
  <div id='laofang' style='position:absolute;left:388px; top:365px;'></div>
  <script language='Javascript'>MySetXY('laofang',520+offx,15+offy);</script>
   <div id="CD" style="position:absolute; left:243px; top:108px; width:72px; height:153px; z-index:1; visibility: hidden"> 
    </div>
</form>
  
<!-- onMouseMove="document.mainmove.oldx.value=event.x+document.body.scrollLeft;document.mainmove.oldy.value=event.y+document.body.scrollTop;" -->
<center>
<p></p>
</center>


</body>
</html>