function ybbcode(temp) {

if (getCookie('ybbcode')!='0'){ //是否打开YBB代码
temp = temp.replace(/&amp;/ig,"&");
temp=temp.replace(/\[(\/|)b\]/ig,"<$1b>");
temp=temp.replace(/\[(\/|)i\]/ig,"<$1i>");
temp=temp.replace(/\[(\/|)u\]/ig,"<$1u>");
temp=temp.replace(/\[(\/|)strike\]/ig,"<$1strike>");
temp=temp.replace(/\[(\/|)center\]/ig,"<$1center>");
temp=temp.replace(/\[(\/|)marquee\]/ig,"<$1marquee>");
temp=temp.replace(/\[QUOTE\]/ig,"<BLOCKQUOTE><strong>引用</strong>：<HR Size=1>");
temp=temp.replace(/\[\/QUOTE\]/ig,"<HR SIZE=1><\/BLOCKQUOTE>");
temp=temp.replace(/(\[font=)([^.:;`'"=\]]*)(\])/ig,"<FONT face='$2'>");
temp=temp.replace(/\[\/font\]/ig,"<\/FONT>");
temp=temp.replace(/(\[COLOR=)([^.:;`'"=\]]*)(\])/ig,"<FONT COLOR='$2'>");
temp=temp.replace(/\[\/COLOR\]/ig,"<\/FONT>");
temp=temp.replace(/(\[size=)([0-9]*)(\])/ig,"<FONT size='$2'>");
temp=temp.replace(/\[\/size\]/ig,"<\/FONT>");
temp=temp.replace(/(\[URL\])([^]]*)(\[\/URL\])/ig,"<A TARGET=_blank HREF='$2'>$2</A>");
temp=temp.replace(/(\[URL=)([^]]*)(\])/ig,"<A TARGET=_blank HREF='$2'>");
temp=temp.replace(/\[\/URL\]/ig,"<\/A>");
temp=temp.replace(/(\[EMAIL\])(\S+\@[^]]*)(\[\/EMAIL\])/ig,"<a href=mailto:$2>$2</a>");
temp=temp.replace(/(\[RM=([0-9]*)\,([0-9]*)\,([A-Z]*)\])([^]]*)(\[\/RM\])/ig,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA width=$2 height=$3><PARAM NAME=SRC VALUE=$5><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=$4><\/OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=60 width=$2><PARAM NAME=CONTROLS VALUE=ControlPanel,StatusBar><PARAM NAME=CONSOLE VALUE=Clip1><\/OBJECT>");
temp=temp.replace(/(\[MP=([0-9]*)\,([0-9]*)\])([^]]*)(\[\/MP\])/ig,"<object classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 width=$2 height=$3><param name=ShowStatusBar value=-1><param name=Filename value=$4><\/object>");
temp=temp.replace(/<A HREF=\"([^"]*)\">/ig,"<A TARGET=_blank HREF=\"$1\">");
temp=temp.replace(/<IMG src=\"([^"]*)\">/ig,"<IMG src=\"$1\" border=0 onmousewheel='return yuzi_img(event,this)' onload='javascript:if(this.width>body.clientHeight)this.width=body.clientHeight'>");
}

if (getCookie('ybbflash')!='0'){  //是否打开[FLASH]代码
temp = temp.replace(/(\[FLASH=([0-9]*)\,([0-9]*)\])([^]]*)(\[\/FLASH\])/ig,"<embed width=$2 height=$3 src='$4' wmode=transparent>");
}

if (getCookie('ybbimg')!='0'){  //是否打开[IMG]代码
temp = temp.replace(/(\[IMG\])([^];]*)(\[\/IMG\])/ig,"<img border=0 src=\"$2\" onmousewheel='return yuzi_img(event,this)' onload='javascript:if(this.width>body.clientHeight)this.width=body.clientHeight'>");
}

if (getCookie('ybbbrow')!='0'){  //是否打开表情代码
temp = temp.replace(/(\[em)([0-9]*)(\])/ig,"<IMG border=0 SRC=images/Emotions/$2.gif>");
}

return (temp);
}


function level(experience,membercode,username,moderated)
{
if (membercode=='0'){levelname="尚未激活";levelimage="<img src=images/level/1.gif>";}
else
if (membercode=='2'){levelname="贵宾会员";levelimage="<img src=images/level/18.gif>";}
else
if (membercode=='4'){levelname="管 理 员";levelimage="<img src=images/level/19.gif>";}
else
if (membercode=='5'){levelname="社区区长";levelimage="<img src=images/level/20.gif>";}
else
if(moderated.indexOf("|"+username+"|") != -1 && moderated!=""){levelname="论坛版主";levelimage="<img src=images/level/18.gif>";}
else
if (experience<= 50){levelname="小野人";levelimage="<img src=images/level/1.gif>";}
else
if (experience<= 150){levelname="野人";levelimage="<img src=images/level/2.gif>";}
else
if (experience<= 500){levelname="土人";levelimage="<img src=images/level/3.gif>";}
else
if (experience<= 1000){levelname="咕噜";levelimage="<img src=images/level/4.gif>";}
else
if (experience<= 2000){levelname="大咕噜";levelimage="<img src=images/level/5.gif>";}
else
if (experience<= 4000){levelname="咕噜队长";levelimage="<img src=images/level/6.gif>";}
else
if (experience<= 8000){levelname="初级守卫";levelimage="<img src=images/level/7.gif>";}
else
if (experience<= 12000){levelname="守卫者";levelimage="<img src=images/level/8.gif>";}
else
if (experience<= 20000){levelname="守卫队长";levelimage="<img src=images/level/9.gif>";}
else
if (experience<= 30000){levelname="斗士";levelimage="<img src=images/level/10.gif>";}
else
if (experience<= 50000){levelname="大斗士";levelimage="<img src=images/level/11.gif>";}
else
if (experience<= 70000){levelname="斗士长";levelimage="<img src=images/level/12.gif>";}
else
if (experience<= 100000){levelname="骑士";levelimage="<img src=images/level/13.gif>";}
else
if (experience<= 130000){levelname="青铜骑士";levelimage="<img src=images/level/14.gif>";}
else
if (experience<= 160000){levelname="白银骑士";levelimage="<img src=images/level/15.gif>";}
else
if (experience<= 200000){levelname="黄金骑士";levelimage="<img src=images/level/16.gif>";}
else
if (experience> 200000){levelname="圣骑士";levelimage="<img src=images/level/17.gif>";}
return('');
}

//放大缩小图片
function yuzi_img(e, o)
{
var zoom = parseInt(o.style.zoom, 10) || 100;
zoom += event.wheelDelta / 12;
if (zoom > 0) o.style.zoom = zoom + '%';
return false;
}