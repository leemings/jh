<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
thesoft=Request.ServerVariables("HTTP_USER_AGENT")
session("yx8_mhjh_inchat")="in"
chatroomsn=session("yx8_mhjh_userchatroomsn")
servername=Request.ServerVariables("server_name")
scriptname=lcase(Request.ServerVariables("script_name"))
if instr(LCase(scriptname),"chatroom.asp")=0 then Response.Redirect "../error.asp?id=012"
useragent=Request.ServerVariables("http_user_agent")
allhttp=Request.ServerVariables("all_http")
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then Response.Redirect "error.asp?id=014"
onlinenum=Application("yx8_mhjh_allonlinenum")
if onlinenum>=yx8_mhjh_maxonline then Response.Redirect "../error.asp?id=058"
nowtime=now()
chatroomnname=application("yx8_mhjh_systemname"&chatroomsn)
onlinelist=Application("yx8_mhjh_onlinelist")
chatroomnum=Application("yx8_mhjh_chatroomnum")
maxnosaytime=Application("yx8_mhjh_nosaytime")
onlinename=Application("yx8_mhjh_onlinename"&chatroomsn)
username=session("yx8_mhjh_username")
usercop=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
'if instr(Application("yx8_mhjh_allonline"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename1"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename2"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename3"),";"&username&";")<>0 then Response.Redirect "../error.asp?id=017"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
rs.Open "select 头像,性别,任务名称,杀人,积分,形态,会员,身份 from 用户 where 姓名='"&username&"'",conn
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=016"
txxx=rs("头像")
jifen=rs("积分")
mysex=rs("性别")
xt=rs("形态")
sr=rs("杀人")
hy=rs("会员")
sf=rs("身份")
zs=rs("任务名称")
rs.Close
set rs=nothing
conn.Close
set conn=nothing
if username="" then Response.Redirect "chaterror.asp?id=004"
if usercop=""  then Response.Redirect "chaterror.asp?id=004"
if instr(onlinename,";"&username&";")=0 then
session("yx8_mhjh_userlasttalktime")=nowtime
session("yx8_mhjh_userlastsavetime")=nowtime
dim newonlinelist()
dim newonlinename()
newonlinenum=0
for i=1 to  chatroomnum
redim preserve newonlinename(i)
newonlinename(i)=";"
next
i=1
j=1
addsuc=false
onlinelistubd=ubound(onlinelist)
do while i<onlinelistubd
if addsuc=false and strcomp(onlinelist(i),username,1)=1 then
newonlinenum=newonlinenum+2
newonlinename(onlinelist(i+5))=newonlinename(onlinelist(i+5))&onlinelist(i)&";"
newonlinename(chatroomsn)=newonlinename(chatroomsn)&username&";"
		redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5),newonlinelist(j+6),newonlinelist(j+7),newonlinelist(j+8),newonlinelist(j+9),newonlinelist(j+10),newonlinelist(j+11)
newonlinelist(j)=username
newonlinelist(j+1)=mysex
newonlinelist(j+2)=usercop
newonlinelist(j+3)=txxx
newonlinelist(j+4)=session("yx8_mhjh_userlasttalktime")
newonlinelist(j+5)=session("yx8_mhjh_userchatroomsn")
		newonlinelist(j+6)=onlinelist(i)
		newonlinelist(j+7)=onlinelist(i+1)
		newonlinelist(j+8)=onlinelist(i+2)
		newonlinelist(j+9)=onlinelist(i+3)
		newonlinelist(j+10)=onlinelist(i+4)
		newonlinelist(j+11)=onlinelist(i+5)
		j=j+12
addsuc=true
elseif datediff("s",onlinelist(i+4),nowtime)<=maxnosaytime then
newonlinenum=newonlinenum+1
newonlinename(onlinelist(i+5))=newonlinename(onlinelist(i+5))&onlinelist(i)&";"
		redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5)
		newonlinelist(j)=onlinelist(i)
		newonlinelist(j+1)=onlinelist(i+1)
		newonlinelist(j+2)=onlinelist(i+2)
		newonlinelist(j+3)=onlinelist(i+3)
		newonlinelist(j+4)=onlinelist(i+4)
		newonlinelist(j+5)=onlinelist(i+5)
		j=j+6
elseif datediff("s",onlinelist(i+4),nowtime)>maxnosaytime then
Application.Lock
Application("yx8_mhjh_allonline")=replace(Application("yx8_mhjh_allonline"),";"&onlinelist(i)&";",";")
Application.UnLock
end if
i=i+6
loop
if addsuc=false then
newonlinenum=newonlinenum+1
newonlinename(chatroomsn)=newonlinename(chatroomsn)&username&";"
	redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5)
newonlinelist(j)=username
newonlinelist(j+1)=mysex
newonlinelist(j+2)=session("yx8_mhjh_usercorp")
newonlinelist(j+3)=txxx
newonlinelist(j+4)=session("yx8_mhjh_userlasttalktime")
newonlinelist(j+5)=session("yx8_mhjh_userchatroomsn")
end if
Application.Lock
if newonlinenum=0 then
dim index(0)
Application("yx8_mhjh_onlinelist")=index(0)
else
Application("yx8_mhjh_onlinelist")=newonlinelist
end if
if sr>9 then
msg="在江湖路上，江湖通缉犯%%正鬼头鬼脑溜进江湖，大家赶快为江湖除害呀，打死了还有赏金呢！”<bgsound src=../mid/scream.wav loop=3>"
elseif xt="鬼魂" then
msg="人生自古谁无死？留取丹心照汗青！此仇不报非君子，忍辱偷生看明朝。突然之间，江湖道上刮起一阵阴风，冤死之魂%%悄悄飘进快乐江湖，哎~~~好可怜，赶紧去城外的凤凰蘖盘吧！<bgsoundsrc=../mid/scream.wav loop=3>"
elseif jifen<500 then
msg="少年不识愁滋味，欲罢还修，欲乐还忧，却道天凉好个秋。新人%%满怀期望进入快乐江湖，愿大家给%%以帮助和照顾，赶紧使用命令里‘帮扶’功能吧，好处多多呢！<bgsoundsrc=../mid/scream.wav loop=3>"
elseif username=Application("yx8_mhjh_admin") then
msg="此生自断天休问，独倚危楼。独倚危楼，不信人间别有愁。江湖站长%%进入快乐江湖"
elseif usercop="官府" then
msg="新丰美酒斗十千，咸阳游侠多少年。相逢意气为君饮，系马高楼垂柳边。官府捕头%%进入快乐江湖，大家小心了！"
elseif sf="掌门" then
msg="何事沉吟？小窗斜日，立遍春阴。翠袖天寒，青衫人老，一样伤心。"&usercop&"掌门%%进入快乐江湖"
elseif sf="副掌门" then
msg="男儿生世间，及壮当封侯。战伐有功业，焉能守旧丘？"&usercop&"副掌门%%进入快乐江湖"
elseif hy="True" then
msg="五陵年少金市东，银鞍白马度春风。落花踏尽游何处，笑入胡姬酒肆中。江湖会员%%进入快乐江湖"
elseif zs<>"无" then
msg="天外有天，人外有人，十年磨一剑，潇洒江湖行！<img src=image/bt.gif width=52 height=41>"&zs&"%%进入快乐江湖"
else
msg=replace(yx8_mhjh_guestjoin,"@#",Application("yx8_mhjh_systemname"&chatroomsn))
end if
Application("yx8_mhjh_allonlinenum")=newonlinenum
for i=1 to chatroomnum
Application("yx8_mhjh_onlinename"&i)=newonlinename(i)
next
Application.UnLock
erase newonlinelist
talkpoint=clng(Application("yx8_mhjh_talkpoint"))
talkarr=Application("yx8_mhjh_talkarr")
dim newtalkarr(600)
j=1
for i=11 to 600 step 10
newtalkarr(j)=talkarr(i)
newtalkarr(j+1)=talkarr(i+1)
newtalkarr(j+2)=talkarr(i+2)
newtalkarr(j+3)=talkarr(i+3)
newtalkarr(j+4)=talkarr(i+4)
newtalkarr(j+5)=talkarr(i+5)
newtalkarr(j+6)=talkarr(i+6)
newtalkarr(j+7)=talkarr(i+7)
newtalkarr(j+8)=talkarr(i+8)
newtalkarr(j+9)=talkarr(i+9)
j=j+10
next
newtalkarr(591)=talkpoint+1
newtalkarr(592)=1
newtalkarr(593)=0
newtalkarr(594)=username
newtalkarr(595)=username
newtalkarr(596)=""
newtalkarr(597)="000000"
newtalkarr(598)="000000"
if username=Application("yx8_mhjh_admin")then 
   if Application("yx8_mhjh_yinshen")=True then 
      newtalkarr(599)=""
   else 
      newtalkarr(599)="<font color=FF0000>【公告】</font><img src=image/horse.gif width=32 height=31>"&msg&"<font class=timsty>"&time()&"</font>"
   end if
else
   newtalkarr(599)="<font color=FF0000>【公告】</font><img src=image/horse.gif width=32 height=31>"&msg&"<font class=timsty>"&time()&"</font>"
end if
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkarr")=newtalkarr
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application.UnLock
erase newtalkarr
session("yx8_mhjh_usertalkpoint")=talkpoint
end if
%>
<head>
<title>快乐江湖</title>
<script language=javascript>
if(window.name!="chat")
{ var i=1;
while (i<=50)
{
window.alert("你想作什么呀，黑我？这里是不行的，去别处玩去吧！哈！慢慢点50次！！");
i=i+1;
}
top.location.href="../exit.asp"
}
var myname='<%=username%>';
var tbclu=true;
var hang=0;
var clsok=0;
var masklist=' ';
function listmask(){resfrm.location.href='masklist.asp?masklist='+masklist;}
function mask(un){if(masklist.indexOf(' '+un+' ')==-1){masklist=masklist+un+' ';}else{var regexp=eval('/ '+un+' /gi');masklist=masklist.replace(regexp,' ');}}
function getmasklist(){return(masklist);}
function lrclutch(){if(this.talkfrm.document.talkform.lrclutch.value=="功能开"){this.talkfrm.document.talkform.lrclutch.value="功能关";this.talkfrm.document.talkform.talkmsg.size=46;this.mainfrm.cols="60,*,150";}else{this.talkfrm.document.talkform.lrclutch.value="功能开";this.talkfrm.document.talkform.talkmsg.size=56;this.mainfrm.cols="0,*,150";}this.talkfrm.document.talkform.talkmsg.focus();}
function youxia(){if(this.talkfrm.document.talkform.youxia.value=="→"){this.talkfrm.document.talkform.youxia.value="←";this.talkfrm.document.talkform.talkmsg.size=46;this.mainfrm.cols="0,*,0";}else{this.talkfrm.document.talkform.youxia.value="→";this.talkfrm.document.talkform.talkmsg.size=56;this.mainfrm.cols="0,*,150";}this.talkfrm.document.talkform.talkmsg.focus();}
function tbclutch(){if(this.talkfrm.document.talkform.tbclutch.value=="关分屏"){this.talkfrm.document.talkform.tbclutch.value="开分屏";this.msgfrm.rows="20,*,0,11,70,0,0";tbclu=false;}else{this.talkfrm.document.talkform.tbclutch.value="关分屏";this.msgfrm.rows="20,*,*,11,70,0,0";tbclu=true;}this.talkfrm.document.talkform.talkmsg.focus();}
function chgsendto(st){this.talkfrm.document.talkform.sendto.options[0].value=st;this.talkfrm.document.talkform.sendto.options[0].text=st;this.talkfrm.document.talkform.sendto.options[0].selected=true;this.talkfrm.document.talkform.talkmsg.focus();}
function settalk(str1,str2){this.talkfrm.document.talkform.talkmsg.value=str1+' '+str2;this.talkfrm.document.talkform.talkmsg.focus();}
function showmsg(isact,isprivacy,username,sendto,expression,namecolor,wordcolor,msg){var msgtmp='';
if (this.talkfrm.document.talkform.tupian.checked != true){msg=msg.replace(/(\[xq\])(\S+)(\[\/xq\])/gi,"");}else{msg=msg.replace(/(\[xq\])(\S+)(\[\/xq\])/gi,"<img src=\'..\/chatroom\/image\/image\/$2\'></img>");}
msg=msg.replace(/(\[img\])(\S+)(\[\/img\])/gi,"<img src=\'..\/chatroom\/image\/image\/$2\'></img>");
hang=hang+1;
if(hang>500 && clsok==0)
{if(confirm("你的屏幕显示的发言行数已经超过了500，考虑到您电脑的性能问题 ，是否清屏？点击“确定”清屏，点击“取消”以后不再出现此提示。")){hang=0;parent.resfrm.location.href='Refresh.asp';}else{clsok=1;}}
if (this.talkfrm.document.talkform.dwtx.checked != true){msg=msg.replace(".wav","");msg=msg.replace(".mid","")}
if(isact=='1' || isact=='2'){
msgtmp="<font color='"+wordcolor+"'>"+msg+"<\/font>";
if(username==myname){msgtmp=msgtmp.replace(/##/gi,"<font color=FF0000>〖<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000>我</font></a>〗</font>");}
else{msgtmp=msgtmp.replace(/##/gi,"<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color="+namecolor+">"+username+"</font></a>");}
if(sendto==myname){msgtmp=msgtmp.replace(/%%/gi,"<font color=FF0000>〖<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000>你</font></a>〗</font>");}
else{msgtmp=msgtmp.replace(/%%/gi,"<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color="+namecolor+">"+sendto+"</font></a>");}
msgtmp="<span class=linesty>"+msgtmp+"</span><br>";}
else{
if(isprivacy=='1'){msgtmp="<span class=linesty>【传音入密】";}
else{msgtmp="<span class=linesty>";}
if(myname==username){msgtmp=msgtmp+"<font color=FF0000>〖<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm'><font color=FF0000>我</font></a>〗</font>";}
else{msgtmp=msgtmp+"<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm'><font color="+namecolor+">"+username+"</font></a>";}
msgtmp=msgtmp+expression;
if(myname==sendto){msgtmp=msgtmp+"对<font color=FF0000>〖<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm'><font color=FF0000>你</font></a>〗</font>说：";}
else{msgtmp=msgtmp+"对<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm'><font color="+namecolor+">"+sendto+"</font></a>说：";}msgtmp=msgtmp+"<font color="+wordcolor+">"+msg+"</font></span><br>";}
if(masklist.indexOf(' '+username+' ')==-1 | ( masklist.indexOf(' '+username+' ')!=-1 & isact=="2" & sendto==myname)){if(tbclu==false){this.msgfrm0.document.write(msgtmp);}else{if(myname==username | myname==sendto){this.msgfrm1.document.write(msgtmp);}else{this.msgfrm0.document.write(msgtmp);}}}}
function chgbgcolor(bgcolor){
this.msgfrm0.document.bgColor=bgcolor;
this.msgfrm1.document.bgColor=bgcolor;
}
function tbygn(){
if(this.talkfrm.document.talkform.tbygn.value=='↓'){
this.talkfrm.document.talkform.tbygn.value='↑';
this.tbymd.rows="20,100%,0" ;
}
else{
this.talkfrm.document.talkform.tbygn.value='↓';
this.tbymd.rows="20,78%,*" ;
}
this.talkfrm.document.talkform.talkmsg.focus();		
}
function refresh1(tx){
parent.msgfrm0.document.open();
parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css.css'></head><\script src=data/"+tx+" ><\/script><body oncontextmenu=self.event.returnValue=false text=000000><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><font color=FF0000>【风云变－震天下】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临快乐江湖<font class=timsty><%=time()%></font><br>");
}
function ziti(ttx){
parent.msgfrm0.document.open();
parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href="+ttx+"></head><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><body text=000000 oncontextmenu=self.event.returnValue=false><font color=FF0000>【风云变－震天下】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临快乐江湖<font class=timsty><%=time()%></font><br>");parent.msgfrm1.document.open();parent.msgfrm1.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href="+ttx+"></head><body text=000000 ><font color=FF0000>【风云变－震天下】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临快乐江湖<font class=timsty><%=time()%></font><br>");
}
</script>
</head>
<frameset name=mainfrm id=mainfrm cols="0,*,148" border=0 frameborder="0" framespacing="0">
<frame name="optfrm" src="option.asp" marginheight=0 marginwidth=0 scrolling=no target="_self">
<frameset name="msgfrm" rows="20,1*,1*,11,70,0,0">
<frame name="topfrm" src="welcome.asp" marginheight=0 marginwidth=0 scrolling=no>
<frame name="msgfrm0" src="about:blank" marginheight=3 marginwidth=0 scrolling=auto border=3 bordercolor=CCCCCC frameborder="Yes">
<frame name="msgfrm1" src="about:blank" marginheight=2 marginwidth=0 scrolling=auto border=3 bordercolor=CCCCCC frameborder="Yes">
<frame name="titlefrm" src="title.asp" marginheight=0 marginwidth=0 scrolling=no border=0 bordercolor=CCCCCC frameborder="no" noresize>
<frame name="talkfrm" src="talk.asp" marginheight=0 marginwidth=0 scrolling=no>
<frame name="sendfrm" src="about:blank">
<frame name="getfrm" src="about:blank">
</frameset>
<frameset rows="27,73%,*" border=0 frameborder="0" name=tbymd>
<frame name="chgfrm" src="about:blank" marginheight=0 marginwidth=0 scrolling=no>
<frame name="resfrm" src="onlinelist.asp" marginheight=0 marginwidth=0 scrolling=auto>
<frame noresize name="f4" src="f4.asp" marginheight=0 marginwidth=0 scrolling=no target="_self">
</frameset>

</frameset>

