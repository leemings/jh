<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=LCase(trim(Request.QueryString("name")))
if name="大家" or name="" then name=aqjh_name
s=Hour(time())
if s>6 and s<17 then home=1
if s>=17 and s<=20 then home=2
if s>=21 or s<=6 then home=3
%>
<HTML><HEAD><TITLE>
<%if name<>aqjh_name then%>
查看--<%=name%>---<%else%>
<%if session("myroom")="h_拥有者" then%>我自己<%else%>配偶<%end if%><%end if%>的小屋客厅</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="style.css" rel=stylesheet>
</head>

<BODY bgColor=#ffffff leftMargin=0 topMargin=0>
<SCRIPT language=JavaScript>
var canchat=false;
function doMouseMove(){
	if((1==event.button)&&(document.eldragged!=null)){
		var mousex=event.clientX+document.body.scrollLeft;
		var mousey=event.clientY+document.body.scrollTop;
		document.eldragged.style.pixelLeft=mousex-document.eldragged.offsetx; 
		document.eldragged.style.pixelTop=mousey-document.eldragged.offsety;
		event.returnValue=false;
	}
}
function checkAtt(el,att){
	while(el!=null){
		if(el.getAttribute(att)!=null)
			return el;
		el=el.parentElement;
	}
	return null;
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function getMaxZindex(){
	if(document.maxzindex==null){
		document.maxzindex=0;
		document.divs=document.all.tags("div");
		for(var i=0;i<document.divs.length;i++)
		
document.maxzindex=(document.maxzindex<document.divs[i].style.zIndex)?document.divs[i].style.zIndex:document.maxzindex;
	}
	return document.maxzindex;
}
function getiZindex(){
	if(document.izindex==null){
		document.izindex=0;
		for(var i=0;i<document.frames.length;i++)
		
document.izindex=(document.izindex<document.frames.style.zIndex)?document.frames.style.zIndex:document.izindex;
	}
	return document.izindex;
}

function doMouseDown(){
	document.eldragged=checkAtt(event.srcElement,"dragarea");
	document.eldragged=checkAtt(document.eldragged,"candrag");
	if(null!=document.eldragged){
		getMaxZindex();
		if(document.eldragged.getAttribute("toztop")!=null)
			document.eldragged.style.zIndex=document.maxzindex++;
		var mousex=event.clientX+document.body.scrollLeft;
		var mousey=event.clientY+document.body.scrollTop;
		document.eldragged.offsetx=mousex-document.eldragged.style.pixelLeft;
		document.eldragged.offsety=mousey-document.eldragged.style.pixelTop;
	}
}
function doSelect(){
	
return((document.eldragged==null)&&null==checkAtt(event.srcElement,"candrag")&&null==checkAtt(event.srcElement,"nosel"));
}
top.document.onmousedown=doMouseDown;
top.document.onmousemove=doMouseMove;
top.document.onmouseup=new Function("document.eldragged=null;");
top.document.ondragstart=doSelect;
top.document.onselectstart=doSelect;

function getDwin(divname,dwloc,url){
dwloc.innerHTML="<IFRAME src="+url+" width='100%' height='100%' frameborder=0 width='100%' height='100%' scrolling='auto' marginwidth='0' marginheight='0' noresize style='visibility:inherit;z-index:34'></IFRAME>&nbsp;.";
}
function changeImage(an_image, newSource){
	var theimg=eval("window.document."+an_image);
    theimg.src = newSource;
}
function resetDwin(){
	email.style.visibility="hidden";
	pet.style.visibility="hidden";
}
function shLayers(n){
	if(n.style.visibility=="visible"){
		n.style.visibility="hidden";
	}else if(n.style.visibility=="hidden"){
		n.style.visibility="visible";
	}
}
function showDwin(w1,w2,url){
	if(!canchat){
		resetDwin();
//		shLayers(w1);
		w1.style.visibility="visible";
		getDwin(w1,w2,url);
	}else{
	//
	}
}
</SCRIPT>
<script language=javascript>window.moveTo(100,50);</script>
<table border="0" height="436" cellpadding="0" cellspacing="0" width="630">
  <tr> 
    <td width="561" height="380" nowrap><img src="images/home<%=home%>.jpg" width="561" height="436"></td>
    <td valign="top" background="images/homedw1.jpg" width="218" height="380">
<table width="100%" border="0" cellpadding="1" cellspacing="2" align="left">
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(email,mailwin,'xiuxi.asp');return false">家中休息</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(email,mailwin,'jl.asp');return false">江湖事记</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(email,mailwin,'money.asp');return false">夫妻钱柜</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(pet,petwin,'pet.asp');return false">我的宠物</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="../diary/mydiary.asp" target="_blank" >心情日记</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(email,mailwin,'set.asp');return false">小屋设置</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(email,mailwin,'xiuli.asp');return false">修理维护</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(email,mailwin,'uproom.asp');return false">小屋升级</a></td>
        </tr>
        <tr> 
          <td height="20" align="center"><img src="images/sign019.gif" width="10" height="10"><a href="#" onClick="showDwin(email,mailwin,'changeroom.asp');return false">我要换房</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">&nbsp; </td>
        </tr>
        <tr> 
          <td height="20" align="center"> <a href="#" onClick="showDwin(email,mailwin,'../hcjs/jhzb/myhua.asp');return false"><img src="images/flower.gif" width="45" height="30" alt="我的鲜花！" border=0></a></td>
        </tr>
        <tr> 
          <td height="20" align="center">&nbsp; </td>
        </tr>
        <tr> 
          <td height="20" align="center">&nbsp; </td>
        </tr>
        <tr> 
          <td height="20" align="center">&nbsp; </td>
        </tr>
      </table>
</td>
</tr>
<tr>
<td colspan="2" background="images/homedw2.jpg">
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.Open "select house.*,用户.* from house inner join 用户 on house.h_拥有者=用户.姓名 where house."& session("myroom") &"='"& name &"'", conn, 1, 1
Response.Write "房间主人资料:" & rs("姓名") & " " & rs("等级") & "级 配偶:" & rs("配偶")
peiou=rs("配偶")
if rs("配偶")<>"无" then
	if Instr(LCase(","&Application("aqjh_useronlinename" & session("nowinroom")))," "&LCase(PEIOU)&" ")=0 then
		Response.Write " 不在聊天室!"
	else
		Response.Write " 在聊天室中!"
	end if
end if
Response.Write " 银两:" & rs("银两") & " 存款:" & rs("存款") & " 存点:" & rs("allvalue")
rs.close
set rs=nothing
conn.close
set conn=nothing   
%>
    </td>
  </tr>
</table>
<SPAN class=transparent></SPAN><SPAN class=p5></SPAN> 
<div id=email style="HEIGHT: 329px; WIDTH: 431px;LEFT: 87px; POSITION: absolute; TOP: 25px; VISIBILITY: hidden;  Z-INDEX: 17; overflow: visible" toztop candrag> 
  <table border=0 cellpadding=0 cellspacing=0 height=17 width="100%">
    <tbody> 
    <tr> 
      <td class=tdline colspan=3><img height=8 src="images/mail1.gif" width=26></td>
    </tr>
    <tr> 
      <td width=26><img height=17 src="images/mail2.gif" 
    width=26></td>
      <td background="images/homechat4.gif" class=cu dragarea></td>
      <td align=right width=23><a href="#" onMouseDown="changeImage('Image13','images/homechata2.gif')" 
      onMouseOut="changeImage('Image13','images/homechata3.gif')"><img border=0 height=17 name=Image13 onMouseUp=shLayers(email); 
      src="images/homechata3.gif" width=15></a><img height=17 
      src="images/homechat5.gif" width=8></td>
    </tr>
    </tbody>
  </table>
  <table border=0 cellpadding=0 cellspacing=0 height="100%" width="100%">
    <tbody> 
    <tr> 
      <td colspan=3> 
        <table border=0 cellpadding=0 cellspacing=0 height="100%" width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#f8d880 width=1></td>
            <td bgcolor=#f8b410 width=2></td>
            <td bgcolor=#996600 width=1></td>
            <td id=mailwin></td>
            <td bgcolor=#ffffcc width=1></td>
            <td bgcolor=#f8b410 width=2></td>
            <td bgcolor=#996600 width=1></td>
          </tr>
          </tbody>
        </table>
      </td>
    </tr>
    <tr> 
      <td bgcolor=#ffff99 height=3> 
        <table border=0 cellpadding=0 cellspacing=0 height=3 width="100%">
          <tbody> 
          <tr bgcolor=#f8d880> 
            <td colspan=5 height=1></td>
            <td bgcolor=#996600 class=tdline rowspan=3 width=1></td>
          </tr>
          <tr bgcolor=#f8b410> 
            <td class=tdline colspan=5 height=2></td>
          </tr>
          <tr bgcolor=#996600> 
            <td colspan=5 height=1></td>
          </tr>
          </tbody>
        </table>
      </td>
    </tr>
    </tbody>
  </table>
</div>
<div id=pet style="HEIGHT: 293px; WIDTH: 342px;LEFT: 135px; POSITION: absolute; TOP: 55px; VISIBILITY: hidden;  Z-INDEX: 17; overflow: visible" toztop candrag> 
<font color=#ffffff id=petwin></font>
</div>
</BODY></HTML>
