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
rs.Open "select ͷ��,�Ա�,��������,ɱ��,����,��̬,��Ա,��� from �û� where ����='"&username&"'",conn
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=016"
txxx=rs("ͷ��")
jifen=rs("����")
mysex=rs("�Ա�")
xt=rs("��̬")
sr=rs("ɱ��")
hy=rs("��Ա")
sf=rs("���")
zs=rs("��������")
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
msg="�ڽ���·�ϣ�����ͨ����%%����ͷ���������������ҸϿ�Ϊ��������ѽ�������˻����ͽ��أ���<bgsound src=../mid/scream.wav loop=3>"
elseif xt="���" then
msg="�����Թ�˭��������ȡ�����պ��࣡�˳𲻱��Ǿ��ӣ�����͵����������ͻȻ֮�䣬�������Ϲ���һ�����磬ԩ��֮��%%����Ʈ�����ֽ�������~~~�ÿ������Ͻ�ȥ����ķ�����̰ɣ�<bgsoundsrc=../mid/scream.wav loop=3>"
elseif jifen<500 then
msg="���겻ʶ����ζ�����ջ��ޣ����ֻ��ǣ�ȴ�������ø������%%��������������ֽ�����Ը��Ҹ�%%�԰������չˣ��Ͻ�ʹ���������������ܰɣ��ô�����أ�<bgsoundsrc=../mid/scream.wav loop=3>"
elseif username=Application("yx8_mhjh_admin") then
msg="�����Զ������ʣ�����Σ¥������Σ¥�������˼���г����վ��%%������ֽ���"
elseif usercop="�ٸ�" then
msg="�·����ƶ�ʮǧ���������������ꡣ�������Ϊ������ϵ���¥�����ߡ��ٸ���ͷ%%������ֽ��������С���ˣ�"
elseif sf="����" then
msg="���³�����С��б�գ����鴺���������캮���������ϣ�һ�����ġ�"&usercop&"����%%������ֽ���"
elseif sf="������" then
msg="�ж������䣬��׳����ս���й�ҵ�������ؾ���"&usercop&"������%%������ֽ���"
elseif hy="True" then
msg="�������ٽ��ж�����������ȴ��硣�仨̤���κδ���Ц����������С�������Ա%%������ֽ���"
elseif zs<>"��" then
msg="�������죬�������ˣ�ʮ��ĥһ�������������У�<img src=image/bt.gif width=52 height=41>"&zs&"%%������ֽ���"
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
      newtalkarr(599)="<font color=FF0000>�����桿</font><img src=image/horse.gif width=32 height=31>"&msg&"<font class=timsty>"&time()&"</font>"
   end if
else
   newtalkarr(599)="<font color=FF0000>�����桿</font><img src=image/horse.gif width=32 height=31>"&msg&"<font class=timsty>"&time()&"</font>"
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
<title>���ֽ���</title>
<script language=javascript>
if(window.name!="chat")
{ var i=1;
while (i<=50)
{
window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ȥ�ɣ�����������50�Σ���");
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
function lrclutch(){if(this.talkfrm.document.talkform.lrclutch.value=="���ܿ�"){this.talkfrm.document.talkform.lrclutch.value="���ܹ�";this.talkfrm.document.talkform.talkmsg.size=46;this.mainfrm.cols="60,*,150";}else{this.talkfrm.document.talkform.lrclutch.value="���ܿ�";this.talkfrm.document.talkform.talkmsg.size=56;this.mainfrm.cols="0,*,150";}this.talkfrm.document.talkform.talkmsg.focus();}
function youxia(){if(this.talkfrm.document.talkform.youxia.value=="��"){this.talkfrm.document.talkform.youxia.value="��";this.talkfrm.document.talkform.talkmsg.size=46;this.mainfrm.cols="0,*,0";}else{this.talkfrm.document.talkform.youxia.value="��";this.talkfrm.document.talkform.talkmsg.size=56;this.mainfrm.cols="0,*,150";}this.talkfrm.document.talkform.talkmsg.focus();}
function tbclutch(){if(this.talkfrm.document.talkform.tbclutch.value=="�ط���"){this.talkfrm.document.talkform.tbclutch.value="������";this.msgfrm.rows="20,*,0,11,70,0,0";tbclu=false;}else{this.talkfrm.document.talkform.tbclutch.value="�ط���";this.msgfrm.rows="20,*,*,11,70,0,0";tbclu=true;}this.talkfrm.document.talkform.talkmsg.focus();}
function chgsendto(st){this.talkfrm.document.talkform.sendto.options[0].value=st;this.talkfrm.document.talkform.sendto.options[0].text=st;this.talkfrm.document.talkform.sendto.options[0].selected=true;this.talkfrm.document.talkform.talkmsg.focus();}
function settalk(str1,str2){this.talkfrm.document.talkform.talkmsg.value=str1+' '+str2;this.talkfrm.document.talkform.talkmsg.focus();}
function showmsg(isact,isprivacy,username,sendto,expression,namecolor,wordcolor,msg){var msgtmp='';
if (this.talkfrm.document.talkform.tupian.checked != true){msg=msg.replace(/(\[xq\])(\S+)(\[\/xq\])/gi,"");}else{msg=msg.replace(/(\[xq\])(\S+)(\[\/xq\])/gi,"<img src=\'..\/chatroom\/image\/image\/$2\'></img>");}
msg=msg.replace(/(\[img\])(\S+)(\[\/img\])/gi,"<img src=\'..\/chatroom\/image\/image\/$2\'></img>");
hang=hang+1;
if(hang>500 && clsok==0)
{if(confirm("�����Ļ��ʾ�ķ��������Ѿ�������500�����ǵ������Ե��������� ���Ƿ������������ȷ���������������ȡ�����Ժ��ٳ��ִ���ʾ��")){hang=0;parent.resfrm.location.href='Refresh.asp';}else{clsok=1;}}
if (this.talkfrm.document.talkform.dwtx.checked != true){msg=msg.replace(".wav","");msg=msg.replace(".mid","")}
if(isact=='1' || isact=='2'){
msgtmp="<font color='"+wordcolor+"'>"+msg+"<\/font>";
if(username==myname){msgtmp=msgtmp.replace(/##/gi,"<font color=FF0000>��<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000>��</font></a>��</font>");}
else{msgtmp=msgtmp.replace(/##/gi,"<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color="+namecolor+">"+username+"</font></a>");}
if(sendto==myname){msgtmp=msgtmp.replace(/%%/gi,"<font color=FF0000>��<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000>��</font></a>��</font>");}
else{msgtmp=msgtmp.replace(/%%/gi,"<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color="+namecolor+">"+sendto+"</font></a>");}
msgtmp="<span class=linesty>"+msgtmp+"</span><br>";}
else{
if(isprivacy=='1'){msgtmp="<span class=linesty>���������ܡ�";}
else{msgtmp="<span class=linesty>";}
if(myname==username){msgtmp=msgtmp+"<font color=FF0000>��<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm'><font color=FF0000>��</font></a>��</font>";}
else{msgtmp=msgtmp+"<a href='javascript:parent.chgsendto(\""+username+"\");'target='talkfrm'><font color="+namecolor+">"+username+"</font></a>";}
msgtmp=msgtmp+expression;
if(myname==sendto){msgtmp=msgtmp+"��<font color=FF0000>��<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm'><font color=FF0000>��</font></a>��</font>˵��";}
else{msgtmp=msgtmp+"��<a href='javascript:parent.chgsendto(\""+sendto+"\");'target='talkfrm'><font color="+namecolor+">"+sendto+"</font></a>˵��";}msgtmp=msgtmp+"<font color="+wordcolor+">"+msg+"</font></span><br>";}
if(masklist.indexOf(' '+username+' ')==-1 | ( masklist.indexOf(' '+username+' ')!=-1 & isact=="2" & sendto==myname)){if(tbclu==false){this.msgfrm0.document.write(msgtmp);}else{if(myname==username | myname==sendto){this.msgfrm1.document.write(msgtmp);}else{this.msgfrm0.document.write(msgtmp);}}}}
function chgbgcolor(bgcolor){
this.msgfrm0.document.bgColor=bgcolor;
this.msgfrm1.document.bgColor=bgcolor;
}
function tbygn(){
if(this.talkfrm.document.talkform.tbygn.value=='��'){
this.talkfrm.document.talkform.tbygn.value='��';
this.tbymd.rows="20,100%,0" ;
}
else{
this.talkfrm.document.talkform.tbygn.value='��';
this.tbymd.rows="20,78%,*" ;
}
this.talkfrm.document.talkform.talkmsg.focus();		
}
function refresh1(tx){
parent.msgfrm0.document.open();
parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css.css'></head><\script src=data/"+tx+" ><\/script><body oncontextmenu=self.event.returnValue=false text=000000><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><font color=FF0000>�����Ʊ䣭�����¡�</font>��ӭ<font color=FF0000>��<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>��</font>���ٿ��ֽ���<font class=timsty><%=time()%></font><br>");
}
function ziti(ttx){
parent.msgfrm0.document.open();
parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href="+ttx+"></head><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><body text=000000 oncontextmenu=self.event.returnValue=false><font color=FF0000>�����Ʊ䣭�����¡�</font>��ӭ<font color=FF0000>��<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>��</font>���ٿ��ֽ���<font class=timsty><%=time()%></font><br>");parent.msgfrm1.document.open();parent.msgfrm1.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href="+ttx+"></head><body text=000000 ><font color=FF0000>�����Ʊ䣭�����¡�</font>��ӭ<font color=FF0000>��<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>��</font>���ٿ��ֽ���<font class=timsty><%=time()%></font><br>");
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

