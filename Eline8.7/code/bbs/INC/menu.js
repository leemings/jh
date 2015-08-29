document.write("<style type=text/css>#master {LEFT: -100px; POSITION: absolute; TOP: 25px; VISIBILITY: visible; WIDTH: 200px; Z-INDEX: 2}#menu {LEFT: 100px; POSITION: absolute; TOP: 25px; VISIBILITY: visible; WIDTH: 18px; Z-INDEX: 5}#top {LEFT: 0px; POSITION: absolute; TOP: 25px; VISIBILITY: visible; WIDTH: 100px; Z-INDEX: 5}#screen {LEFT: 0px; POSITION: absolute; TOP: 31px; VISIBILITY: visible; WIDTH: 100px; Z-INDEX: 5}#screenlinks {LEFT: 0px; POSITION: absolute; TOP: 31px; VISIBILITY: visible; WIDTH: 100px; Z-INDEX: 5}</style>")
var ie = document.all ? 1 : 0
var ns = document.layers ? 1 : 0

if(ie){
	document.write('<style type="text/css">')
	document.write("#screen	{filter:Alpha(Opacity=30);}")
	document.write("</style>")
}

if(ns){
	document.write('<style type="text/css">')
	document.write("#master	{clip:rect(0,150,250,0);}")
	document.write("</style>")
}

var master = new Object("element")
master.curLeft = -100;	master.curTop = 10;
master.gapLeft = 0;		master.gapTop = 0;
master.timer = null;

function moveAlong(layerName, paceLeft, paceTop, fromLeft, fromTop){
	clearTimeout(eval(layerName).timer)
	if(eval(layerName).curLeft != fromLeft){
		if((Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft)) < paceLeft){eval(layerName).curLeft = fromLeft}
		else if(eval(layerName).curLeft < fromLeft){eval(layerName).curLeft = eval(layerName).curLeft + paceLeft}
			else if(eval(layerName).curLeft > fromLeft){eval(layerName).curLeft = eval(layerName).curLeft - paceLeft}
		if(ie){document.all[layerName].style.left = eval(layerName).curLeft}
		if(ns){document[layerName].left = eval(layerName).curLeft}
	}
	if(eval(layerName).curTop != fromTop){
   if((Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop)) < paceTop){eval(layerName).curTop = fromTop}
		else if(eval(layerName).curTop < fromTop){eval(layerName).curTop = eval(layerName).curTop + paceTop}
			else if(eval(layerName).curTop > fromTop){eval(layerName).curTop = eval(layerName).curTop - paceTop}
		if(ie){document.all[layerName].style.top = eval(layerName).curTop}
		if(ns){document[layerName].top = eval(layerName).curTop}
	}
	eval(layerName).timer=setTimeout('moveAlong("'+layerName+'",'+paceLeft+','+paceTop+','+fromLeft+','+fromTop+')',30)
}

function setPace(layerName, fromLeft, fromTop, motionSpeed){
	eval(layerName).gapLeft = (Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft))/motionSpeed
	eval(layerName).gapTop = (Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop))/motionSpeed
	moveAlong(layerName, eval(layerName).gapLeft, eval(layerName).gapTop, fromLeft, fromTop)
}

var expandState = 0

function expand(){
	if(expandState == 0){setPace("master", 0, 10, 10); if(ie){document.menutop.src = "pic/menui.gif"}; expandState = 1;}
	else{setPace("master", -100, 10, 10); if(ie){document.menutop.src = "pic/menuo.gif"}; expandState = 0;}
}
document.write(menustr);

if(ie){var sidemenu = document.all.master;}
if(ns){var sidemenu = document.master;}

function FixY(){
	if(ie){sidemenu.style.top = document.body.scrollTop+10}
	if(ns){sidemenu.top = window.pageYOffset+10}
}

setInterval("FixY()",100);





var currentpos,timer;
function initialize()
{
timer=setInterval ("scrollwindow ()",30);
}
function sc()
{
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos !=document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize

