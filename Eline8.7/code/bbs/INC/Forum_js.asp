<!--#include file="../z_plus_conn.asp"-->
<script language="javascript">
function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
</script>
<script Language="JavaScript">
//***********Ĭ�����ö���.*********************
tPopWait=50;//ͣ��tWait�������ʾ��ʾ��
tPopShow=5000;//��ʾtShow�����ر���ʾ
showPopStep=20;
popOpacity=99;

//***************�ڲ���������*****************
sPop=null;
curShow=null;
tFadeOut=null;
tFadeIn=null;
tFadeWaiting=null;

document.write("<style type='text/css'id='defaultPopStyle'>");
document.write(".cPopText {  <%=forum_body(4)%> color:#000000; border: 1px #000000 solid;font-color: font-size: 12px; padding-right: 4px; padding-left: 4px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0)}");
document.write("</style>");
document.write("<div id='dypopLayer' style='position:absolute;z-index:1000;' class='cPopText'></div>");


function showPopupText(){
var o=event.srcElement;
if(o!=null){
	MouseX=event.x;
	MouseY=event.y;
	if(o.alt!=null && o.alt!=""){
		o.dypop=o.alt;
		o.alt="";
	}
	if(o.title!=null && o.title!="" && o.tagName!="FORM"){
		o.dypop=o.title;
		o.title="";
	}
	if(o.dypop!=sPop) {
			sPop=o.dypop;
			clearTimeout(curShow);
			clearTimeout(tFadeOut);
			clearTimeout(tFadeIn);
			clearTimeout(tFadeWaiting);	
			if(sPop==null || sPop=="") {
				dypopLayer.innerHTML="";
				dypopLayer.style.filter="Alpha()";
				dypopLayer.filters.Alpha.opacity=0;	
				}
			else {
				if(o.dyclass!=null) popStyle=o.dyclass;
					else popStyle="cPopText";
				curShow=setTimeout("showIt()",tPopWait);
			}
		}
	}
}

function showIt(){
		dypopLayer.className=popStyle;
		dypopLayer.innerHTML=sPop;
		popWidth=dypopLayer.clientWidth;
		popHeight=dypopLayer.clientHeight;
		if(MouseX+12+popWidth>document.body.clientWidth) popLeftAdjust=-popWidth-24
			else popLeftAdjust=0;
		if(MouseY+12+popHeight>document.body.clientHeight) popTopAdjust=-popHeight-24
			else popTopAdjust=0;
		dypopLayer.style.left=MouseX+12+document.body.scrollLeft+popLeftAdjust;
		dypopLayer.style.top=MouseY+12+document.body.scrollTop+popTopAdjust;
		dypopLayer.style.filter="Alpha(Opacity=0)";
		fadeOut();
}

function fadeOut(){
	if(dypopLayer.filters.Alpha.opacity<popOpacity) {
		dypopLayer.filters.Alpha.opacity+=showPopStep;
		tFadeOut=setTimeout("fadeOut()",1);
		}
		else {
			dypopLayer.filters.Alpha.opacity=popOpacity;
			tFadeWaiting=setTimeout("fadeIn()",tPopShow);
			}
}

function fadeIn(){
	if(dypopLayer.filters.Alpha.opacity>0) {
		dypopLayer.filters.Alpha.opacity-=1;
		tFadeIn=setTimeout("fadeIn()",1);
		}
}
document.onmouseover=showPopupText;

function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }

//�����˵���ش���
 var h;
 var w;
 var l;
 var t;
 var topMar = 1;
 var leftMar = -2;
 var space = 1;
 var isvisible;
 var MENU_SHADOW_COLOR='#999999';//���������˵���Ӱɫ
 var global = window.document
 global.fo_currentMenu = null
 global.fo_shadows = new Array

function HideMenu() 
{
 var mX;
 var mY;
 var vDiv;
 var mDiv;
	if (isvisible == true)
{
		vDiv = document.all("menuDiv");
		mX = window.event.clientX + document.body.scrollLeft;
		mY = window.event.clientY + document.body.scrollTop;
		if ((mX < parseInt(vDiv.style.left)) || (mX > parseInt(vDiv.style.left)+vDiv.offsetWidth) || (mY < parseInt(vDiv.style.top)-h) || (mY > parseInt(vDiv.style.top)+vDiv.offsetHeight)){
			vDiv.style.visibility = "hidden";
			isvisible = false;
		}
}
}

function ShowMenu(vMnuCode,tWidth) {
	vSrc = window.event.srcElement;
	vMnuCode = "<table id='submenu' cellspacing=1 cellpadding=3 style='width:"+tWidth+"' class=tableborder1 onmouseout='HideMenu()'><tr height=23><td nowrap align=left class=tablebody1>" + vMnuCode + "</td></tr></table>";

	h = vSrc.offsetHeight;
	w = vSrc.offsetWidth;
	l = vSrc.offsetLeft + leftMar+4;
	t = vSrc.offsetTop + topMar + h + space-2;
	vParent = vSrc.offsetParent;
	while (vParent.tagName.toUpperCase() != "BODY")
	{
		l += vParent.offsetLeft;
		t += vParent.offsetTop;
		vParent = vParent.offsetParent;
	}

	menuDiv.innerHTML = vMnuCode;
	menuDiv.style.top = t;
	menuDiv.style.left = l;
	menuDiv.style.visibility = "visible";
	isvisible = true;
	makeRectangularDropShadow(submenu, MENU_SHADOW_COLOR, 4)
}

function makeRectangularDropShadow(el, color, size)
{
	var i;
	for (i=size; i>0; i--)
	{
		var rect = document.createElement('div');
		var rs = rect.style
		rs.position = 'absolute';
		rs.left = (el.style.posLeft + i) + 'px';
		rs.top = (el.style.posTop + i) + 'px';
		rs.width = el.offsetWidth + 'px';
		rs.height = el.offsetHeight + 'px';
		rs.zIndex = el.style.zIndex - i;
		rs.backgroundColor = color;
		var opacity = 1 - i / (i + 1);
		rs.filter = 'alpha(opacity=' + (100 * opacity) + ')';
		el.insertAdjacentElement('afterEnd', rect);
		global.fo_shadows[global.fo_shadows.length] = rect;
	}
}
//�û�״̬
var mystatus= '<a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=online&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"����\">����</a><br><a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=hidden&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"����\">����</a><br><a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=hidden3&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"æµ\">æµ</a><br><a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=hidden4&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"���ϻ���\">���ϻ���</a><br><a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=hidden5&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"�뿪\">�뿪</a><br><a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=hidden6&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"�����绰\">�����绰</a><br><a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=hidden7&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"����Ͳ�\">����Ͳ�</a><br><a style=font-size:9pt;line-height:14pt; href=\"cookies.asp?action=hidden8&userid=<%=userid%>\" onclick=document.all.mystat.innerText=\"����\">����</a>'
//�û��������
var manage= '<a style=font-size:9pt;line-height:14pt; href=\"JavaScript:openScript(\'messanger.asp?action=new\',500,400)\">������</a><br><a style=font-size:9pt;line-height:14pt; href=\"dispuser.asp?id=<%=userid%>&boardid=<%=boardid%>&action=permission\">������ʲô</a><br><a style=font-size:9pt;line-height:14pt; href=\"topicwithme.asp?s=2\">�ҷ��������</a><br><a style=font-size:9pt;line-height:14pt; href=\"topicwithme.asp?s=1\">�Ҳ��������</a><br><a style=font-size:9pt;line-height:14pt; href=\"mymodify.asp\">���������޸�</a><br><a style=font-size:9pt;line-height:14pt; href=\"modifypsw.asp\">�û������޸�</a><br><a style=font-size:9pt;line-height:14pt; href=\"modifyadd.asp\">��ϵ�����޸�</a><br><a style=font-size:9pt;line-height:14pt; href=\"usersms.asp\">�û����ŷ���</a><br><a style=font-size:9pt;line-height:14pt; href=\"friendlist.asp\">�༭�����б�</a><br><a style=font-size:9pt;line-height:14pt; href=\"favlist.asp\">�û��ղع���</a><!--<br><a style=font-size:9pt;line-height:14pt; href=\"myfile.asp\">�����ļ�����</a>--><br><a style=font-size:9pt;line-height:14pt; href=\"z_fastpost.asp\">���ٷ���ѡ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"z_visual.asp\">�����������</a>'
//ģ���б�
var stylelist = '<%
myCache.name="stylelist"
if myCache.valid then
response.write replace(myCache.value,"{boardid}",boardid)

else
dim stylelist
stylelist="<a style=font-size:9pt;line-height:12pt; href=\""cookies.asp?action=stylemod&skinid=0&boardid={boardid}\"">�ָ�Ĭ������</a><br>"
set rs=conn.execute("select id,skinname from config order by id")
do while not rs.eof
stylelist=stylelist & "<a style=font-size:9pt;line-height:12pt; href=\""cookies.asp?action=stylemod&skinid="&rs(0)&"&boardid={boardid}\"">"&rs(1)&"</a><br>"
rs.movenext
loop
myCache.add stylelist,dateadd("n",9999,now)
response.write replace(stylelist,"{boardid}",boardid)
end if
%>'
//��̳״̬
var boardstat= '<a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?boardid=<%=boardid%>\">��������ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?action=lasttopicnum&boardid=<%=boardid%>\">������ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?action=lastbbsnum&boardid=<%=boardid%>\">������ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?reaction=online&boardid=<%=boardid%>\">����ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?reaction=onlineinfo&boardid=<%=boardid%>\">�������</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?reaction=onlineUserinfo&boardid=<%=boardid%>\">�û�������ͼ��</a>'
//��̳�ղ�
//var downlist= '<a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=0&boardid=<%=boardid%>\">�ļ������</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=1&boardid=<%=boardid%>\">ͼƬ�����</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=2&boardid=<%=boardid%>\">Flash���</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=3&boardid=<%=boardid%>\">���ּ����</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=4&boardid=<%=boardid%>\">��Ӱ�����</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp\">�ؿ�����</a>'
<!--��������������ʼ-->
<%set rs= server.createobject ("adodb.recordset")
sqlgroup = "select * from [group] order by groupid"
rs.open sqlgroup,connplus,0,1
dim rsplus,pid
dim padstr,menustr,submenustr,menucount
if not (rs.bof and rs.eof) then
	menustr=""
	menucount=0
	do while not rs.eof
		response.write "var plus"&rs("groupid")&"='"
		set rsplus= server.createobject ("adodb.recordset")
		sqlgroup="select * from plus where groupid="&rs("groupid")&" order by plusid"
		rsplus.open sqlgroup,connplus,0,1
		do while not rsplus.eof
			if rsplus("window")<>"limit" then
				response.write "<a "&rsplus("window")&" style=font-size:9pt;line-height:14pt; href=\'"&rsplus("url")&"\'>"&rsplus("plusname")&"</a><br>"
			else
				response.write "<a style=font-size:9pt;line-height:14pt; href=\""#\"" onclick=\""window.open(\'"&rsplus("url")&"\',\'\',\'scrollbars=no,width="&rsplus("width")&",height="&rsplus("height")&",resizable=0,top=0,left=0\')\"">"&rsplus("plusname")&"</a><br>"
			end if
			rsplus.movenext
		loop
		rsplus.close
		response.write "';"&vbcrlf
		menustr=menustr&"<font color=999900><font face=Wingdings>1</font>"&rs("groupname")&"</font><br>"
		menucount=menucount+1
		rsplus.open sqlgroup,connplus,0,1
		do while not rsplus.eof
			if rsplus("window")<>"limit" then
				submenustr="<a "&rsplus("window")&" href=\'"&rsplus("url")&"\'>"&rsplus("plusname")&"</a><br>"
			else
				submenustr="<a href=\""#\"" onclick=\""window.open(\'"&rsplus("url")&"\',\'\',\'scrollbars=no,width="&rsplus("width")&",height="&rsplus("height")&",resizable=0,top=0,left=0\')\"">"&rsplus("plusname")&"</a><br>"
			end if
			rsplus.movenext
			if rsplus.eof then
				submenustr="��"&submenustr
			else
				submenustr="��"&submenustr
			end if
			menucount=menucount+1
			menustr=menustr&submenustr
		loop
		rsplus.close
		set rsplus=nothing
		rs.movenext
	loop
	padstr="var menustr='<div id=master>"
	padstr=padstr&"<div id=menu onmouseover=\""javascript:expand()\"">"
	padstr=padstr&"<table border=0 cellpadding=0 cellspacing=0 width=18>"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td width=\""100%\""><a href=\""javascript:expand()\"" onFocus=this.blur()><img alt=�������չ��/�رտ�ݲ˵� border=0 height=70 name=menutop src=\""pic/menuo.gif\"" width=18></a></td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</div>"
	padstr=padstr&"<div id=top>"
	padstr=padstr&"<table border=0 cellpadding=0 cellspacing=0 width=100>"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td width=\""100%\""><img border=0 height=6 src=\""pic/menutop.gif\""  width=100></td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</div>"
	padstr=padstr&"<div id=screen onmouseout=\""javascript:expand()\"">"
	padstr=padstr&"<table border=0 cellpadding=5 cellspacing=0 width=100>"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td bgcolor=#3399FF width=\""100%\"">"
	padstr=padstr&"<table bgcolor=#3399FF border=0 cellpadding=0 cellspacing=0 width=\""100%\"" height="&(menucount*16+28)&">"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td width=\""100%\"">"
	padstr=padstr&"<table border=0 cellpadding=5 cellspacing=1 width=\""100%\"">"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td bgcolor=#ecf6f5 width=\""100%\""><br><br><br><br><br></td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</div>"
	padstr=padstr&"<div id=screenlinks>"
	padstr=padstr&"<table border=0 cellpadding=6 cellspacing=0 width=100>"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td style=\""FILTER: alpha(opacity=90)\"" width=\""100%\"">"
	padstr=padstr&"<table bgcolor=#336666 border=0 cellpadding=0 cellspacing=0 width=\""100%\"">"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td width=\""100%\"">"
	padstr=padstr&"<table border=0 cellpadding=6 cellspacing=1 width=\""100%\"">"
	padstr=padstr&"<tbody>"
	padstr=padstr&"<tr>"
	padstr=padstr&"<td bgcolor=#ecf6f5 width=\""100%\"" align=\""left\"">"
	padstr=padstr&"<font color=999900>�������񵼺�</font><br>"
	padstr=padstr&menustr
	padstr=padstr&"<a href=\""logout.asp\"")\""><font color=999900><font face=Wingdings>v</font>�˳�����</font></a><br>"
	padstr=padstr&"</td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</td>"
	padstr=padstr&"</tr>"
	padstr=padstr&"</tbody>"
	padstr=padstr&"</table>"
	padstr=padstr&"</div>"
	padstr=padstr&"</div>"
	response.write padstr&"';"
end if
rs.close
set rs=nothing%>
<!--����������������-->
</script>
<script language="JavaScript">
nereidFadeObjects = new Object();
nereidFadeTimers = new Object();

function nereidFade(object, destOp, rate, delta) {
	if (!document.all) return
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
function strLength(str) {
	var WINNT_CHINESE;
	WINNT_CHINESE = (("��̳").length==2)
	if(WINNT_CHINESE) {
		var l,t,c;
		var i;
		l=str.length;
		t=l;
		for(i=0;i<l;i++) {
			c=str.charCodeAt(i);
			if(c<0 || c>255) t=t+1;
		}
		return t;
	} else  {
		return str.length;
	}
}
//һ���˵�
var ALLboardlist= '<%
myCache.name="ALLBoardList"
Dim ALLBoardList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
select case rs(2)
case 0
ALLBoardList = ALLBoardList & ""
case 1
ALLBoardList = ALLBoardList & "<font color=cc0000><b>��</b></font>"
case 2
ALLBoardList = ALLBoardList & "����"
end select
ALLBoardList = ALLBoardList & "<a style=font-size:9pt;line-height:12pt; href=\""list.asp?boardid="&rs(0)&"\"">"&replace(rs(1),"'","��")&"</a><br>"
rs.MoveNext
loop
set rs=nothing
myCache.add ALLBoardList,dateadd("n",60,now)
response.write ALLBoardList
%>'


//�����˵�
var boardlist= '<%
if BoardParentID>0 then
myCache.name="boardlist"
Dim boardlist,boardrootid
set rs=conn.execute("select rootid from board where boardid="&boardid)
boardrootid=rs(0)
set rs=nothing
set rs=conn.execute("select boardid,boardtype,depth from board where rootid="&boardrootid&" and depth<>0 order by rootid,orders")
do while not rs.EOF
select case rs(2)
case 0
boardlist = boardlist & ""
case 1
boardlist = boardlist & "<font color=cc0000><b>��</b></font>"
case 2
boardlist = boardlist & "����"
end select
BoardList = BoardList & "<a style=font-size:9pt;line-height:12pt; href=\""list.asp?boardid="&rs(0)&"\"">"&replace(rs(1),"'","��")&"</a><br>"
rs.MoveNext
loop
set rs=nothing
myCache.add boardlist,dateadd("n",60,now)
response.write boardlist
end if
%>'
</script>
