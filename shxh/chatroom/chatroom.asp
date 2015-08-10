<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=session("Ba_jxqy_userchatroomsn")
servername=Request.ServerVariables("server_name")
scriptname=lcase(Request.ServerVariables("script_name"))
if instr(scriptname,"chatroom.asp")=0 then Response.Redirect "../error.asp?id=012"
//useragent=Request.ServerVariables("http_user_agent")
//if instr(useragent,"MSIE ")=0 then Response.Redirect "../error.asp?id=013"
allhttp=Request.ServerVariables("all_http")
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then Response.Redirect "error.asp?id=014"
onlinenum=Application("Ba_jxqy_allonlinenum")
if onlinenum>=Application("Ba_jxqy_maxonline") then Response.Redirect "../error.asp?id=058"
advertisemenheight=Application("Ba_jxqy_advertisemenheight")
nowtime=now()
chatroomnname=application("Ba_jxqy_systemname"&chatroomsn)
onlinelist=Application("Ba_jxqy_onlinelist")
chatroomnum=Application("Ba_jxqy_chatroomnum")
maxnosaytime=Application("Ba_jxqy_nosaytime")
onlinename=Application("Ba_jxqy_onlinename"&chatroomsn)
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "chaterror.asp?id=004"
if instr(onlinename,";"&username&";")=0 then
session("Ba_jxqy_userlasttalktime")=nowtime
session("Ba_jxqy_userlastsavetime")=nowtime
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
		redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5),newonlinelist(j+6),newonlinelist(j+7),newonlinelist(j+8),newonlinelist(j+9),newonlinelist(j+10),newonlinelist(j+11),newonlinelist(j+12),newonlinelist(j+13),newonlinelist(j+14),newonlinelist(j+15)
		newonlinelist(j)=username
		newonlinelist(j+1)=session("Ba_jxqy_usersex")
		newonlinelist(j+2)=session("Ba_jxqy_usercorp")
		newonlinelist(j+3)=session("Ba_jxqy_userip")
		newonlinelist(j+4)=session("Ba_jxqy_userlasttalktime")
		newonlinelist(j+5)=session("Ba_jxqy_userchatroomsn")
		newonlinelist(j+6)=session("Ba_jxqy_usermoral")
		newonlinelist(j+7)=session("Ba_jxqy_useridentity")
		newonlinelist(j+8)=onlinelist(i)
		newonlinelist(j+9)=onlinelist(i+1)
		newonlinelist(j+10)=onlinelist(i+2)
		newonlinelist(j+11)=onlinelist(i+3)
		newonlinelist(j+12)=onlinelist(i+4)
		newonlinelist(j+13)=onlinelist(i+5)
		newonlinelist(j+14)=onlinelist(i+6)
		newonlinelist(j+15)=onlinelist(i+7)
		j=j+16
		addsuc=true
	elseif datediff("s",onlinelist(i+4),nowtime)<=maxnosaytime then
		newonlinenum=newonlinenum+1
		newonlinename(onlinelist(i+5))=newonlinename(onlinelist(i+5))&onlinelist(i)&";"
		redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5),newonlinelist(j+6),newonlinelist(j+7)
		newonlinelist(j)=onlinelist(i)
		newonlinelist(j+1)=onlinelist(i+1)
		newonlinelist(j+2)=onlinelist(i+2)
		newonlinelist(j+3)=onlinelist(i+3)
		newonlinelist(j+4)=onlinelist(i+4)
		newonlinelist(j+5)=onlinelist(i+5)
		newonlinelist(j+6)=onlinelist(i+6)
		newonlinelist(j+7)=onlinelist(i+7)
		j=j+8
	elseif datediff("s",onlinelist(i+4),nowtime)>maxnosaytime then
		Application.Lock
		Application("Ba_jxqy_allonline")=replace(Application("Ba_jxqy_allonline"),";"&onlinelist(i)&";",";")
		Application.UnLock
	end if
	i=i+8
loop
if addsuc=false then
	newonlinenum=newonlinenum+1
	newonlinename(chatroomsn)=newonlinename(chatroomsn)&username&";"
	redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5),newonlinelist(j+6),newonlinelist(j+7)
	newonlinelist(j)=username
	newonlinelist(j+1)=session("Ba_jxqy_usersex")
	newonlinelist(j+2)=session("Ba_jxqy_usercorp")
	newonlinelist(j+3)=session("Ba_jxqy_userip")
	newonlinelist(j+4)=session("Ba_jxqy_userlasttalktime")
	newonlinelist(j+5)=session("Ba_jxqy_userchatroomsn")
	newonlinelist(j+6)=session("Ba_jxqy_usermoral")
	newonlinelist(j+7)=session("Ba_jxqy_useridentity")
end if
Application.Lock
if newonlinenum=0 then
	dim index(0)
	Application("Ba_jxqy_onlinelist")=index(0)
else	
	Application("Ba_jxqy_onlinelist")=newonlinelist
end if
msg=replace(Application("Ba_jxqy_guestjoin"),"@#",Application("Ba_jxqy_systemname"&chatroomsn))
Application("Ba_jxqy_allonlinenum")=newonlinenum
for i=1 to chatroomnum
	Application("Ba_jxqy_onlinename"&i)=newonlinename(i)
next
Application.UnLock
erase newonlinelist
talkpoint=clng(Application("Ba_jxqy_talkpoint"))
talkarr=Application("Ba_jxqy_talkarr")
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
newtalkarr(599)="<font color=FF0000>【公告】</font>"&msg&"<font class=timsty>"&time()&"</font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkarr")=newtalkarr
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application.UnLock
erase newtalkarr
session("Ba_jxqy_usertalkpoint")=talkpoint
end if
%>
<head>
<title><%=Application("Ba_jxqy_systemname"&chatroomsn)%></title>
<script language=javascript>
var myname='<%=username%>';
var advertisemenheight=<%=advertisemenheight%>;
var tbclu=true;
var masklist=' ';
function listmask(){resfrm.location.href='masklist.asp?masklist='+masklist;}
function mask(un){if(masklist.indexOf(' '+un+' ')==-1){masklist=masklist+un+' ';}else{var regexp=eval('/ '+un+' /gi');masklist=masklist.replace(regexp,' ');}}
function getmasklist(){return(masklist);}
function lrclutch(){
	var mf = document.getElementById("mainfrm");
	//alert(mf);
	if(this.talkfrm.document.talkform.lrclutch.value=="开功能菜单"){this.talkfrm.document.talkform.lrclutch.value="关功能菜单";this.talkfrm.document.talkform.talkmsg.size=46;this.mainfrm.cols="60,*,140";}else{this.talkfrm.document.talkform.lrclutch.value="开功能菜单";this.talkfrm.document.talkform.talkmsg.size=56;this.mainfrm.cols="0,*,140";}this.talkfrm.document.talkform.talkmsg.focus();
	}
function tbclutch(){if(this.talkfrm.document.talkform.tbclutch.value=="关分屏窗口"){this.talkfrm.document.talkform.tbclutch.value="开分屏窗口";this.msgfrm.rows=advertisemenheight+",*,0,70,0,0";tbclu=false;}else{this.talkfrm.document.talkform.tbclutch.value="关分屏窗口";this.msgfrm.rows=advertisemenheight+",*,*,70,0,0";tbclu=true;}this.talkfrm.document.talkform.talkmsg.focus();}
function chgsendto(st){this.talkfrm.document.talkform.sendto.options[0].value=st;this.talkfrm.document.talkform.sendto.options[0].text=st;this.talkfrm.document.talkform.sendto.options[0].selected=true;this.talkfrm.document.talkform.talkmsg.focus();}
function settalk(str1,str2){this.talkfrm.document.talkform.talkmsg.value=str1+' '+str2;this.talkfrm.document.talkform.talkmsg.focus();}
function showmsg(isact,isprivacy,username,sendto,expression,namecolor,wordcolor,msg){var msgtmp='';
	msg=msg.replace(/(\[msg\])(.+)(\[\/msg\])(<font class=timsty>\S+)/i,"<marquee><font color=FF0000>【传言】</font>"+username+"：$2$4<\/marquee>")
	msg=msg.replace(/(\[img\])(\S+)(\[\/img\])/gi,"<img src=\'..\/images\/image\/$2\'></img>");
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
</script>
</head>
<frameset id=mainfrm name=mainfrm cols="0,*,140" border=0>
	<frame name="optfrm" src="option.asp" marginheight=0 marginwidth=0 scrolling=yes>
	<frameset name="msgfrm" rows="<%=advertisemenheight%>,1*,1*,70,0,0"  >
		<frame name="topfrm" src="welcome.asp" marginheight=0 marginwidth=0 scrolling=no>
		<frame name="msgfrm0" src="about:blank" marginheight=3 marginwidth=0 scrolling=auto border=3 bordercolor=CCCCCC frameborder="Yes">
		<frame name="msgfrm1" src="about:blank" marginheight=3 marginwidth=0 scrolling=auto border=3 bordercolor=CCCCCC frameborder="Yes">
		<frame name="talkfrm" src="talk.asp" marginheight=0 marginwidth=0 scrolling=no>
		<frame name="sendfrm" src="about:blank">
		<frame name="getfrm" src="about:blank">
	</frameset>
	<frameset rows="20,*" border=0>
		<frame name="chgfrm" src="about:blank" marginheight=0 marginwidth=0 scrolling=no>
		<frame name="resfrm" src="onlinelist.asp" marginheight=0 marginwidth=0 scrolling=auto>
	</frameset>
</frameset>		