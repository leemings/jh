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
	var timeValue ="<font color=blue>��ӭ�㣺<%=membername%>��"
    if(hours<3)
		{timeValue  += ("��ҹ�ˣ�������Ϣ��Ҫע������Ӵ�������Ǹ����ı�Ǯ��")}	
    else if(hours<6)
		{timeValue  += ("�賿�ˣ�һ������������������졪���ţ���������˯һ�����")}
	else if(hours<8)	
       {timeValue  += ("�糿�ã�С������˯���н��ѡ����������ʵĿ��������µ�һ�쿪ʼ�ˣ�")}
	else if(hours<11)	
       {timeValue  += ("����ã�����æµ�Ź�����һ��Ҫע�����ݽ��Ӵ��ǧ����ۻ��ˣ�")}	   
    else if(hours<13) 
       {timeValue += ("����ã�����Ӵ�����ֲ���˯�����绹Ҫ�����ء������Ӱ��Ͱɡ�")}
    else if(hours<17)
       {timeValue +=("����ã����ǹ�����������˾��������������������ǲ��ǰ���������")}
    else if(hours<19)
       {timeValue +=("����ã������������ЪЪ�ˣ��úõ�������žۣ������Ӵ��")}
    else
       {timeValue +=("���Ϻã�������һ������Ҫ˯һ���þ��ˣ�˵��������������һ�������أ�")}	   
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