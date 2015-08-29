minimizebar="minimize.gif";   //窗口右上角最小化“按钮”的图片
minimizebar2="minimize2.gif"; //鼠标悬停时最小化“按钮”的图片
closebar="close.gif";         //窗口右上角关闭“按钮”的图片
closebar2="close2.gif";       //鼠标悬停时关闭“按钮”的图片
icon="icon.gif";              //窗口左上角的小图标

function noBorderWin(fileName,w,h,titleBg,moveBg,titleColor,titleWord,scr)  //定义一个弹出无边窗口的函数，能数意义见下面“参数说明”，实际使用见最后的实例。
/*
------------------参数说明-------------------
fileName   ：无边窗口中显示的文件。
w　　　　   ：窗口的宽度。
h　　　　   ：窗口的高度。
titleBg    ：窗口“标题栏”的背景色以及窗口边框颜色。
moveBg     ：窗口拖动时“标题栏”的背景色以及窗口边框颜色。
titleColor ：窗口“标题栏”文字的颜色。
titleWord  ：窗口“标题栏”的文字。
scr        ：是否出现滚动条。取值yes/no或者1/0。
--------------------------------------------
*/
{
  var contents="<html>"+
               "<head>"+
	    	   "<title>"+titleWord+"</title>"+
			   "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\">"+
			   "<object id=hhctrl type='application/x-oleobject' classid='clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11'><param name='Command' value='minimize'></object>"+
			   "</head><script>HeiRui_Studio_Player.close()</script>"+
               "<body leftMargin='0' topMargin='0' MARGINHEIGHT='0' MARGINWIDTH='0' scroll=no oncontextmenu='self.event.returnValue=false;' onkeydown='if(event.keyCode==78&&event.ctrlKey)return false;' onselectstart='event.returnValue=false' ondragstart='window.event.returnValue=false'>"+
			   "  <table height=100% width=100% cellpadding=0 cellspacing=0 bgcolor="+titleBg+" id=mainTab>"+
			   "    <tr height=18 style=cursor:default; onmousedown='x=event.x;y=event.y;setCapture();mainTab.bgColor=\""+moveBg+"\";' onmouseup='releaseCapture();mainTab.bgColor=\""+titleBg+"\";' onmousemove='if(event.button==1)self.moveTo(screenLeft+event.x-x,screenTop+event.y-y);'>"+
			   "      <td width=100% style=cursor:hand><map name='FPMap0'><area coords='231, 8, 8' shape='circle' href=### onmousedown=hhctrl.Click();><area coords='248, 10, 8' shape='circle' href=### onmousedown=self.close();></map><img border=0 src=images/HeiRui_r1_c1.gif usemap=#FPMap0 width=276 height=19></td>"+
			   "    </tr>"+
			   "    <tr height=*>"+
			   "      <td colspan=4>"+
			   "        <iframe name=nbw_v6_iframe src="+fileName+" scrolling="+scr+" width=100% height=100% frameborder=0></iframe>"+
			   "      </td>"+
			   "    </tr>"+
			   "  </table>"+
			   "</body>"+
			   "</html>";
  pop=window.open("","HeiRui_Studio_Player","fullscreen=yes");
  pop.resizeTo(w,h);
  pop.moveTo((screen.width-w)/2,(screen.height-h)/2);
  pop.document.writeln(contents);

  if(pop.document.body.clientWidth!=w||pop.document.body.clientHeight!=h)  //如果无边窗口不是出现在纯粹的IE窗口中
  {
    temp=window.open("","nbw_v6");
	temp.close();
	window.showModalDialog("about:<"+"script language=javascript>window.open('','nbw_v6','fullscreen=yes');window.close();"+"</"+"script>","","dialogWidth:0px;dialogHeight:0px");
	pop2=window.open("","nbw_v6");
    pop2.resizeTo(w,h);
    pop2.moveTo((screen.width-w)/2,(screen.height-h)/2);
    pop2.document.writeln(contents);
	pop.close();
  }
}
function DJplay_song()
{
window.open("about:blank","HeiRui_Studio_Player","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,top=10,left=10,width=230,height=350");
}
function open_window(url,windowname,size) 
{ 
window.open(url,windowname,"toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,top=10,left=10,"+ size); 
}
 
function CheckOthers(form)
{
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
//		if (e.name != 'chkall')
			if (e.checked==false)
			{
				e.checked = true;// form.chkall.checked;
			}
			else
			{
				e.checked = false;
			}
	}
}

function CheckAll(form)
{
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
//		if (e.name != 'chkall')
			e.checked = true// form.chkall.checked;
	}
}

nereidFadeObjects = new Object();
nereidFadeTimers = new Object();
function nereidFade(object, destOp, rate, delta){
if (!document.all)
return
    if (object != "[object]"){
        setTimeout("nereidFade("+object+","+destOp+","+rate+","+delta+")",0);
        return;
    }
    clearTimeout(nereidFadeTimers[object.sourceIndex]);
    diff = destOp-object.filters.alpha.opacity;
    direction = 1;
    if (object.filters.alpha.opacity > destOp){
        direction = -1;
    }
    delta=Math.min(direction*diff,delta);
    object.filters.alpha.opacity+=direction*delta;
    if (object.filters.alpha.opacity != destOp){
        nereidFadeObjects[object.sourceIndex]=object;
        nereidFadeTimers[object.sourceIndex]=setTimeout("nereidFade(nereidFadeObjects["+object.sourceIndex+"],"+destOp+","+rate+","+delta+")",rate);
    }
}
