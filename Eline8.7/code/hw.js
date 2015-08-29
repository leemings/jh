function addfav()
	{window.external.addFavorite('http://51eline.com','『E线江湖』')}
function shuaxin()
	{javascript:history.go(0)}
var Mouse_Obj="none";
var pX
var pY
document.onmousemove=D_NewMouseMove;
function m(c_Obj)
{ Mouse_Obj=c_Obj;
pX=parseInt(document.all(Mouse_Obj).style.left)-event.x;
pY=parseInt(document.all(Mouse_Obj).style.top)-event.y; }
function D_NewMouseMove()
{if(Mouse_Obj!="none")
{document.all(Mouse_Obj).style.left=pX+event.x;
document.all(Mouse_Obj).style.top=pY+event.y;
event.returnValue=false;}}
function D_NewMouseUp()
{Mouse_Obj="none";}
  document.onmouseup=mu;
  function mu(){D_NewMouseUp();}
function pop1(win){
parent.w.location.href=win;}
function loc(win){
window.location.href=win;}
function pop(win){
window.open(win,'','toolbar,menubar,scrollbars,resizable,status,location,directories,copyhistory,height=480,width=600')}
function fsetico()
{
	divCount = document.all.ico.all.tags("DIV");
	for (i=0; i<divCount.length; i++) {
	obj = divCount(i);
	win_x=document.body.clientWidth-ico_left;//图标的纵坐标范围
	n=Math.round(win_x/ico_x);//每列可排列的个数
	num=i+1;
if (n!=0){
	l=Math.round(num/n+0.49);//图标的列数
	numy=i-n*l+n;//在每一列中图标的序数
	obj.style.top=(l-1)*ico_y+ico_top;
	obj.style.left=numy*ico_x+ico_left;
	}
	}
}
function show(c_Str)
{
document.all(c_Str).style.display='block';
}
function hide(c_Str)
{
if(document.all(c_Str).style.display=='none')
	{
	return;
	}
document.all(c_Str).style.display='none';
}
function keypress()
{
if(event.keyCode==17){
startclick();
s_img.focus();}
}
function sh(c_Str)
{
if(document.all(c_Str).style.display=='none')
	{
	document.all(c_Str).style.display='block';
	}
	else
	{
	document.all(c_Str).style.display='none';
	}
}
function showmenuie5()
{
	var rightedge=document.body.clientWidth-event.clientX
	var bottomedge=document.body.clientHeight-event.clientY
if (rightedge<ie5menu.offsetWidth)
	ie5menu.style.left=document.body.scrollLeft+event.clientX-ie5menu.offsetWidth
else
	ie5menu.style.left=document.body.scrollLeft+event.clientX
if (bottomedge<ie5menu.offsetHeight)
	ie5menu.style.top=document.body.scrollTop+event.clientY-ie5menu.offsetHeight
else
	ie5menu.style.top=document.body.scrollTop+event.clientY
	ie5menu.style.visibility="visible"
//ie5menu.style.visibility=""
	ie5menu.filters.alpha.opacity=0
	GradientShow()
	return false
}
function hidemenuie5()
{
	ie5menu.style.visibility="hidden"
	GradientClose()}
function GradientShow() //实现淡入的函数
{
	ie5menu.filters.alpha.opacity+=intInterval
	if (ie5menu.filters.alpha.opacity<100) setTimeout("GradientShow()",intDelay)
	}
function GradientClose() //实现淡出的函数
	{ie5menu.filters.alpha.opacity-=intInterval
if (ie5menu.filters.alpha.opacity>0) {
	setTimeout("GradientClose()",intDelay)}
else {ie5menu.style.visibility="hidden"}}
function link(act,txt,hs,ar)
{document.write("<div class=\"menuout\" onMouseOver=this.className='menuover'; onMouseOut=this.className='menuout'; style=\"padding-left:14;padding-top:2;padding-bottom:1\" onclick="+act+">"+txt+"<span style=\"visibility: hidden\">"+ar+"</span><font face=\"Webdings\" style=\"display:"+hs+"\">4</font></div>")}
function hr()
{document.write("<hr>")}
