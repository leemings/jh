<script language="JavaScript">
<!--
showtime();
function showtime()
{
	if (!document.all&&!document.getElementById)
		{return}
		
	var ClockObj=document.getElementById? document.getElementById("clock"): document.all.clock
    var now = new Date()
    var hours = now.getHours()
	var timeValue ="<font color=blue>欢迎你：<%=membername%>，"
    if(hours<3)
		{timeValue  += ("深夜了，还不休息？要注意身体哟，他可是革命的本钱！")}	
    else if(hours<6)
		{timeValue  += ("凌晨了，一觉醒来，看看窗外的天——嗯，还可以再睡一会儿！")}
	else if(hours<8)	
       {timeValue  += ("早晨好，小鸟把你从睡梦中叫醒。呼吸着新鲜的空气——新的一天开始了！")}
	else if(hours<11)	
       {timeValue  += ("上午好，还在忙碌着工作吗？一定要注意劳逸结合哟，千万别累坏了！")}	   
    else if(hours<13) 
       {timeValue += ("中午好，好困哟，可又不敢睡，下午还要工作呢。哎，加把油吧。")}
    else if(hours<17)
       {timeValue +=("下午好，又是工作，看着上司的那张老脸，你的心理是不是暗暗生气？")}
    else if(hours<19)
       {timeValue +=("傍晚好，现在总算可以歇歇了，好好的与家人团聚，好惬意哟！")}
    else
       {timeValue +=("晚上好，辛勤了一天的你可要睡一个好觉了！说不定今晚还可以做一个好梦呢！")}	   
	timeValue  +="</font>"
    ClockObj.innerHTML = timeValue 
}
//-->
</script>

<script language="javascript1.1">
<!--
var dn
c1=new Image(); c1.src="images/img_gupiao/clock/c1.gif"
c2=new Image(); c2.src="images/img_gupiao/clock/c2.gif"
c3=new Image(); c3.src="images/img_gupiao/clock/c3.gif"
c4=new Image(); c4.src="images/img_gupiao/clock/c4.gif"
c5=new Image(); c5.src="images/img_gupiao/clock/c5.gif"
c6=new Image(); c6.src="images/img_gupiao/clock/c6.gif"
c7=new Image(); c7.src="images/img_gupiao/clock/c7.gif"
c8=new Image(); c8.src="images/img_gupiao/clock/c8.gif"
c9=new Image(); c9.src="images/img_gupiao/clock/c9.gif"
c0=new Image(); c0.src="images/img_gupiao/clock/c0.gif"
cb=new Image(); cb.src="images/img_gupiao/clock/cb.gif"
cam=new Image(); cam.src="images/img_gupiao/clock/cam.gif"
cpm=new Image(); cpm.src="images/img_gupiao/clock/cpm.gif"

show3()

function extract(h,m,s,type){
	if (!document.images)
		return
	if (h<=9){
		document.images.a.src=cb.src
		document.images.b.src=eval("c"+h+".src")
	}
	else {
	document.images.a.src=eval("c"+Math.floor(h/10)+".src")
	document.images.b.src=eval("c"+(h%10)+".src")
	}
	if (m<=9){
		document.images.d.src=c0.src
		document.images.e.src=eval("c"+m+".src")
	}
	else {
		document.images.d.src=eval("c"+Math.floor(m/10)+".src")
		document.images.e.src=eval("c"+(m%10)+".src")
	}
	if (s<=9){
		document.g.src=c0.src
		document.images.h.src=eval("c"+s+".src")
	}
	else {
	document.images.g.src=eval("c"+Math.floor(s/10)+".src")
	document.images.h.src=eval("c"+(s%10)+".src")
	}
	if (dn=="AM") document.j.src=cam.src
	else document.images.j.src=cpm.src
}

function show3(){
	if (!document.images)
	return
	var Digital=new Date()
	var hours=Digital.getHours()
	var minutes=Digital.getMinutes()
	var seconds=Digital.getSeconds()
	dn="AM" 
	if ((hours>=12)&&(minutes>=1)||(hours>=13)){
	dn="PM"
	hours=hours-12
	}
	if (hours==0)
	hours=12
	extract(hours,minutes,seconds,dn)
	setTimeout("show3()",1000)
}
//-->
</script>