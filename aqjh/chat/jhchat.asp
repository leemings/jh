
<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const2.asp"-->
<!--#include file="../const3.asp"-->
<!--#include file="../const4.asp"-->
<!--#include file="../chk.asp"-->
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
application("npc_sxr")=aqjh_name
application("npc_sxsj")=now()
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(aqjh_name)&" ")=0 then ydl=0
	if ydl=1 and clng(nowinroom)<>i then 
		Session.Abandon
		Response.Redirect "../error.asp?id=140"
		Response.End 
	end if
next 
'����
chatroomname=trim(Application("aqjh_chatroomname"&session("nowinroom")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_yjdh=1 then Chkyjdh()
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if aqjh_disproxy=1 then  Chkproxy()
Dim SplitReflashPage
Dim DoReflashPage
dim shuaxin_time
DoReflashPage=true
shuaxin_time=chatshuaxin_time
ReflashTime=Now()
if (not isnull(session("ReflashTime"))) and cint(shuaxin_time)>0 and DoReflashPage then
	if DateDiff("s",session("ReflashTime"),Now())<cint(shuaxin_time) then
   	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3><font style='font-size:9pt'>��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��<b><font color=ff0000>"&shuaxin_time&"</font></b>��������ˢ�±�ҳ��<BR>���ڴ�ҳ�棬���Ժ򡭡�</font>"
	response.end
	else
	session("ReflashTime")=Now()
	end if
elseif isnull(session("ReflashTime")) and cint(shuaxin_time)>0 and DoReflashPage then
	Session("ReflashTime")=Now()
end if
'���û����ϴ���
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM sm where a='����'",conn,2,2
mybanner=rs("c")
jhauto=rs("d")
rs.close
rs.open "select k from r where a='"&chatroomname&"'",conn,3,3
if not(rs.eof and rs.bof) then
	jrht=rs("k")
else
	jrht="��ӭ����["&chatroomname&"]ף���ĵĿ���~~"
end if
rs.close
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
tjrf=rs("ͨ��")
newuser=rs("times")-1
if newuser=3 then
newuser=2
end if
aqjh_id=rs("id")
hydj=rs("��Ա�ȼ�")
jhsf=rs("���")
if tjrf=True then
	jhmp="ͨ����"
else
	jhmp=rs("����")
end if
nickname=rs("����")
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
jhjh=rs("����")
hymd=rs("��������")
mywife=rs("��ż")
guojia=rs("����")
zhiwei=rs("ְλ")
mmp=rs("����")
myzs=rs("ת��")
dj=rs("�ȼ�")
hk=rs("����")
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
 jhsf="���"
end if
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then 
 jhsf="����"
end if
if rs("��ż")=Application("aqjh_user") and rs("�Ա�")="Ů" then
 jhsf="վ������"
end if
if Application("aqjh_mengzhu")=aqjh_name then 
 jhsf="��������"
end if
userinfo2=0
mytimes="<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
if aqjh_name=Application("aqjh_user") then
	zhanz="<font color=red>[�ٸ�]</font>"
else
	zhanz="<font color=red>[�ٸ�]</font>"
end if
if dj<18 then
 aqjh_userinto=aqjh_userinto1
else
 aqjh_userinto=aqjh_userinto2
end if
if hydj>0 then
 aqjh_userinto=aqjh_userinto3
end if
if jhsf="���" then
 aqjh_userinto=aqjh_userinto4
end if
if aqjh_grade=4 then
 aqjh_userinto=aqjh_userinto5
elseif aqjh_grade=5 then
 aqjh_userinto=aqjh_userinto6
end if
if hk<1 or isnull(hk) then
    conn.execute "update �û� set ����=0 where ����='"&aqjh_name&"'"
else
    userinfo2=1
    aqjh_userinto=aqjh_userinto7
    conn.execute "update �û� set ����=����+18,����=����-1 where ����='"&aqjh_name&"'"
end if
if tjrf=True then
 userinfo2=1
 aqjh_userinto=aqjh_userinto8
end if
if jhsf="����" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto9
end if
if jhsf="��������" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto10
end if
if aqjh_grade>5 and aqjh_grade<10 then
 aqjh_userinto=aqjh_userinto11
elseif aqjh_grade=10 then
 aqjh_userinto=aqjh_userinto12
end if
if zhiwei="����" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto13
end if
if zhiwei="ة��" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto14
end if
if zhiwei="����" then
 userinfo2=1
 aqjh_userinto=aqjh_userinto15
end if
aqjh_userinto=aqjh_userinto&mytimes
if newuser=1 then
	userinfo2=1
	conn.Execute "update �û� set ����=����+5000000,��=��+10,ľ=ľ+10,ˮ=ˮ+10,��=��+10,��=��+10,times=times+1 where ����='"&aqjh_name&"'"
	Response.Write "<script Language=Javascript>alert('���������˷ѡ��\n\n��ӭ����"&aqjh_name&"��\n�������������������ȷ������վ�����͵����˷�\n����5���򡢽�ľˮ������10�㣡\n��Ǯ�����񿨸�1�ţ�\nֻҪ��ϲ����������������ǹ�ͬ�ļ�԰��\n��������᲻�Ͻ�ʶ��������˷�������\n�����кܶ����ʵ��������Ȥ��������\nף���ڱ�������Զ���ġ���죡\n\n�������Ľ������İ� ���� ����');</script>"
      aqjh_userinto="<img src=img/new.gif>##��ȡ��վ���͵����˷�:����<font color=#cc0000><b>5����</b></font>/��ľˮ������<font color=#cc0000><b>10��</b></font>�������������ˣ���Ҷ�������Ҫ����չ˰�!<br>[<font color=ff0000>%%</font>]��һ���������ǽ������Ҷ���չˡ���"
end if
rs.close
rs.open "select id,a,d,g,h,i,l from vh where b='"&aqjh_name&"' and c=true and j=false",conn
if rs.bof and rs.eof then
 aqjh_userinto=aqjh_userinto
elseif userinfo2=0 then
	aqjh_userinto=rs("d")
	vhwj=rs("h")
	vhname=rs("g")
	id=rs("id")
	vhnj=rs("i")
	sj=DateDiff("d",rs("l"),now())
	randomize timer
	ii=int(rnd()*10)
        if aqjh_grade=10 then
	aqjh_userinto="���ǵ�"&zhanz&"##����"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">���ڸ�·���ɵĴ�ӵ�£�������ʹ⻷������"&Application("aqjh_chatroomname")&"�������һ������⣡"&mytimes&"<br><font color=red>��������Ϣ��</font>"&zhanz&"%%���������ˣ����·Ӣ�ۻ�ӭ��������[%%]��������1000000�㣡��������1000000��"
	conn.execute "update �û� set ����=����+1000000,����=����+1000000 where ����='"&aqjh_name&"'"
	end if
        if aqjh_grade=5 then
	aqjh_userinto="##<font color=blue>["&mmp&"����]</font>����"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">���ڵ��ӵ���ͬ�£��������������˽��������һ����С���ǣ������ˣ���"&mytimes&"<br><font color=red>��������Ϣ��</font>["&mmp&"]����[%%]���������ˣ�����е��ӻ�ӭ������[%%]��������10000�㣡��������10000��"
	conn.execute "update �û� set ����=����+10000,����=����+10000 where ����='"&aqjh_name&"'"
	end if
        if aqjh_grade>5 and aqjh_grade<10 then
	aqjh_userinto="##[<font color=red>�ٸ�</font>"&jhsf&"]����"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">���ڵ��������У�ɱ�����ڵ������ˡ�"&Application("aqjh_chatroomname")&"����ŭ�������������Ҳ����"&mytimes&"<br><font color=red>��������Ϣ��</font>[�ٸ�"&jhsf&"][%%]���������ˣ�����С�Ĵ��¡�������[%%]��������50000�㣡��������50000��"
	conn.execute "update �û� set ����=����+50000,����=����+50000 where ����='"&aqjh_name&"'"
	end if
     if aqjh_grade>4 then
	sj=30 
	vhnj=50
     else
	if sj>42 or vhnj<1 then
		conn.execute "update vh set j=true where id="&id
                aqjh_userinto="##ƮƮȻ�ģ������ˡ�"&Application("aqjh_chatroomname")&"������С�ó����ɾ�����·�������񣡡�"
	else
		if sj>35 or vhnj<10 then
			if ii<3 then
				conn.execute "update vh set j=true where id="&id
				aqjh_userinto="##["&mmp&""&jhsf&"]�������Լҵ�"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">Ҫ����"&Application("aqjh_chatroomname")&"���������İ������ݳ���Щ����--���϶���ʹ�ù��Ȼ����ʧ�ޣ����ˣ�"&mytimes&"<br><font color=red>��������Ϣ��</font>[%%]���������ˣ�ֻ���ܲ����뽭������������졣������"
			else
				if Isnull(vhname) or vhname="" or vhname="��" then vhname=rs("a")
				aqjh_userinto=replace(aqjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				aqjh_userinto=aqjh_userinto&"<br><font color=red>��������Ϣ��</font>��СϺ�����ؿ���[%%]����¶��һ����ɷĽɷ�ı��飬��ʱ�Լ�Ҳ���С�����[%%]��������28�㣡��������1000��"
				conn.execute "update �û� set ����=����+28,����=����+1000 where ����='"&aqjh_name&"'"
				conn.execute "update vh set i=i-1 where id="&id
			end if
		else
			if ii<2 then
				aqjh_userinto="##["&mmp&""&jhsf&"]�������Լҵ�"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">Ҫ����"&Application("aqjh_chatroomname")&"�����������Ͻ�ͨ������ֻ��һ·�ܲ�ǰ��������"&Application("aqjh_chatroomname")&"ʱ��������졣"&mytimes&"<br><font color=red>��������Ϣ��</font>��СϺ��������������[%%]����¶��һ������ı��顣����[%%]��������18�㣡"
				conn.execute "update �û� set ����=����+18 where ����='"&aqjh_name&"'"
			else
				if Isnull(vhname) or vhname="" or vhname="��" then vhname=rs("a")
				aqjh_userinto=replace(aqjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				aqjh_userinto=aqjh_userinto&"<br><font color=red>��������Ϣ��</font>��СϺ�����ؿ���[%%]����¶��һ����ɷĽɷ�ı��飬��ʱ�Լ�Ҳ���С�����[%%]��������28�㣡��������1000��"
				conn.execute "update �û� set ����=����+28,����=����+1000 where ����='"&aqjh_name&"'"
			end if
			conn.execute "update vh set i=i-1 where id="&id
		end if
	end if
    end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
scrname=Request.ServerVariables("SCRIPT_NAME")
if Instr(LCase(scrname),"jhchat.asp")=0 then Response.Redirect "../error.asp?id=002"
Response.Write "<bgsound src=img\luckybug1.mid loop=1 volume=50>"
if Application("aqjh_closedoor")="1" then Response.Redirect "../error.asp?id=100"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("aqjh_disproxy")="1" and (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
t=s & ":" & f & ":" & m
'Session("aqjh_lastsaytime")=sj

if Session("aqjh_inthechat")<>"1" then
if InStr(LCase(Application("aqjh_useronlinename"&nowinroom))," " & LCase(aqjh_name) & " ")<>0 then Response.Redirect "../error.asp?id=300"
Session("aqjh_inthechat")="1"
Session("aqjh_savetime")=now()
Session("aqjh_lasttime")=sj
myzanli=0
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 then myzanli=1
'����0,�Ա�1,����2,���3,ͷ��4,�ȼ�5,id6,��Ա�ȼ�7,����8,����9,��ż10,ת��11,����12,ְλ13,����14
myonline = aqjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & aqjh_jhdj& "|" & aqjh_id& "|" & hydj&"|"&myzanli&"|"&"����"&"|"&mywife&"|"&myzs&"|"&guojia&"|"&zhiwei&"|"&jhjh
Application.Lock
dim newonlinelist()
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")

'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'���������ֺ���ƴ������
'if yjl=0 and len(onuser(2))<len(jhmp) then			'���������ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(2))>len(jhmp) then			'���������ֳ���,�̵���ǰ
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'�����ֺ���ƴ������
'if yjl=0 and len(onuser(0))<len(aqjh_name) then	'�����ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(0))>len(aqjh_name) then	'�����ֳ���,�̵���ǰ
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if
Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)& " "&aqjh_name & " "
erase  newonlinelist
erase  onlinelist
Application.UnLock
Session("SayCount")=Application("SayCount")
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
		Application("SayCount")=Application("SayCount")+1
		i="SayStr"&YuShu(Application("SayCount"))
		Application(i)=Str
	Application.UnLock()
End Sub
act="����"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
aqjh_userinto=Replace(aqjh_userinto,"%%","<a href=javascript:parent.sw(\'[" & aqjh_name & "]\'); target=f2>" & aqjh_name &"</a>")
says="<font color=#cc0000>����½������</font><font color=008800>" & Replace(aqjh_userinto,"##","<img src="&jhtx&"><a href=javascript:parent.sw(\'[" & aqjh_name & "]\'); target=f2>" & aqjh_name &"</a><font color=#000000>ID:"& aqjh_id &"</b></font>") & "<font color=#000000>[ <font color=blue>�Ա�:</font><font color=red>"& sex &"</font> <font color=blue>�ȼ�:</font><font color=red>"& aqjh_jhdj&"</font> <font color=blue>����:</font><font color=red>"& guojia &"</font> <font color=blue>ְλ:</font><font color=red>"& zhiwei &"</font> <font color=blue>����:</font><font color=red>"& jhmp &"</font> <font color=blue>���:</font><font color=red>"& jhsf &"</font> <font color=blue>��ż:</font><font color=red>"& mywife &"</font> <font color=blue>��Ա�ȼ�:</font><font color=red>"& hydj &"</font> <font color=blue>ת��:</font><font color=red>"& myzs &"</font> <font color=blue>����:</font><font color=red>"& jhjh &"</font>]</font><font class=t>(" & t & ")</font><bgsound src=readonly/okok.wma loop=1>"
if newuser=1 then
randomize()
rnd2=int(rnd*400)+1
says=says & "<input type=button value=��ȡ���˷� onClick=""javascript:xrf"&rnd2&".disabled=1;window.open(\'sjfunc/xrf.asp\')"" name=xrf"&rnd2&" >"
'says=says & "<input type=button value=��ȡ���˷� >"
end if
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
end if
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
hiddenadmin="|"&Application("hidden_admin")&"|"
hiddenadmin=Replace(hiddenadmin,",","|")
if aqjh_grade>=9 then hiddenadmin="||"
randomize()
pic=int(rnd*3)
r=int(rnd*136)
b=int(rnd*136)
pcinfo = Request.ServerVariables("HTTP_USER_AGENT")
If InStr(pcinfo, "NT 5") <> 0 Then
 gdsd = 4
Else
 gdsd = 7
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<style>
BODY{background-image: URL(bbg1.jpg);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}\n");</style>
<script Language="Javascript">
var Banner=new Array();
var wen = new Array();
var huida = new Array();
<%
banners=Split(Trim(mybanner),";")
total=UBound(banners)
for o = 0 to total-1
	temp1=banners(o)
	temp1=replace(temp1,chr(34),"")
	temp1=replace(temp1,chr(13),"")
	 Response.Write "Banner[" & o & "]="& chr(34) & temp1 & chr(34) & ";" & chr(13) & chr(10)
next
erase banners
banners=Split(Trim(jhauto),";")
total=UBound(banners)
for o = 0 to total-1
	temp=Split(Trim(banners(o)),"|")
	temp1=temp(0)
	temp2=temp(1)
	temp1=replace(temp1,chr(34),"")
	temp1=replace(temp1,chr(13),"")
	temp2=replace(temp2,chr(34),"")
	temp2=replace(temp2,chr(13),"")
	 Response.Write "wen[" & o & "]="& chr(34) & temp1 & chr(34) & ";" & chr(13) & chr(10)
 	 Response.Write "huida[" & o & "]="& chr(34) & temp2 & chr(34) & ";" & chr(13) & chr(10)
 	 erase temp
next
erase banners
%>
function shake(n) {if (window.top.moveBy) {for (i = 10; i > 0; i--) {for (j = n; j > 0; j--) {window.top.moveBy(0,i);window.top.moveBy(i,0);window.top.moveBy(0,-i);window.top.moveBy(-i,0);}}}}
if(window!=window.top){window.alert("��ʹ��ie�����ʹ�ñ�ϵͳ��");top.location.href="../exit.asp"}
if(window.name!="aqjh"){ var i=1;while (i<=50){window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ȥ�ɣ�����������50�Σ���");i=i+1;}top.location.href="../exit.asp"}
var crm="<%=chatroomname%>",bgc="<%=Application("aqjh_chatcolor")%>",systitle="<%=Application("aqjh_tltie")%>";
var myn="<%=aqjh_name%>",mywife="<%=mywife%>",chatbgcolor="<%=chatbgcolor%>",chatimage="<%=chatimage%>";
var hiddenadmin="<%=hiddenadmin%>",cs=<%=aqjh_grade%>,automan="<%=Application("aqjh_automanname")%>";
var myroom=<%=session("nowinroom")%>,slbox=0,lst=0,tbclu=true,mdcls=true,listfaces=false;
var myguojia="<%=session("guojia")%>",myzhiwei="<%=session("zhiwei")%>";
var showhao=false,myxq=1,mymp=1,jhtx="<%=jhtx%>",showpy="<%=hymd%>",showmp="<%=jhmp%>",headhigh=16;
var showsex="<%=sex%>",showseek="",bc = 1,clsok=0,hang = 0 ;showtype = 0,bgimg=""
var mysayc=0,lastsay=new Array(31);
for(mysayc=0;mysayc<31;mysayc++){lastsay[mysayc]=0;}
var gdsd = <%=gdsd%>;
<%if aqjh_dongtai=1 then%>jsjsstr="<\script src=gg.js></\script>";<%else%>jsjsstr="";<%end if%>
var badword = new Array(<%=aqjh_badword%>);
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/����������������������������������������?������������������" ;
var writejs=false;
var Maxwrite=100;
var askjs="��ϵͳ���Ѵ��ڽű����ã�����ָ���Զ��������ʹ������ָ��";
var askqp2="<br><font size=2><font color=red>����ʾ��</font>�Ի������أ�10���Ӻ��Զ�������</font><br>";
var writeNUM=0;
var askqp="<br><font size=2><font color=red>����ʾ��</font>����<a href='javascript:parent.qp()'>[��]</a>����Լ������Դ����ʾ���κ��Զ�������</font><br>";
document.write("<title>��ӭ����[<%=Application("aqjh_chatroomname")%>],ף���ĵĿ��ģ�</title>");
function write(cls){var fsize,lheight;
if(cls==1){fsize=this.f2.document.af.fs.value;lheight=this.f2.document.af.lh.value;}else{fsize='10';lheight='125';}
this.f1.document.open();
this.f1.document.writeln("<html><head><title>�Ի���</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f1.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"����\";CURSOR:url('3.cur');font-size:" + fsize + "pt}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL(imgg/<%=r%>.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f1.document.writeln("<\Script Language=\"JavaScript1.1\">var autoScrollOn=1;var scrollOnFunction;var scrollOffFunction;");
this.f1.document.writeln("function scrollit(){if(!parent.f2.document.af.as.checked){autoScrollOn=0;return true;}else{autoScrollOn=1;StartUp();return true;}}");
this.f1.document.writeln("function scrollWindow(){if(autoScrollOn==1){this.scroll(0,65000);parent.f0.scroll(0,65000);setTimeout('scrollWindow()',200);}}");
this.f1.document.writeln("function scrollOn(){autoScrollOn=1;scrollWindow();}");
this.f1.document.writeln("function scrollOff(){autoScrollOn=0;}");
this.f1.document.writeln("function StartUp(){this.onblur=scrollOnFunction;this.onfocus=scrollOffFunction;scrollWindow();}");
this.f1.document.writeln("scrollOnFunction=new Function('scrollOn()');");
this.f1.document.writeln("scrollOffFunction=new Function('scrollOff()');");
this.f1.document.writeln("StartUp();</\script>");
this.f1.document.writeln("<body background='"+bgimg+"' oncontextmenu=self.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f1.document.writeln("<\script language='javascript' src='QQ.js'><\/script><span class=l><font color=red>�������ˢ�¡�</font>���һ�ӭ<font color=red>��"+myn+"��</font>������"+crm+"����<font class=t>(<%=t%>)</font></span>");
this.f1.document.writeln("<%=roomtemp%><br>");
this.f0.document.open();
this.f0.document.writeln("<html><head><title>������ʾ</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f0.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"����\";CURSOR:url('2.cur');font-size:" + fsize + "pt;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL(imgg/<%=b%>.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f0.document.writeln("<body background='"+bgimg+"'oncontextmenu=self.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f0.document.writeln("<\script language='javascript' src='QQ.js'><\/script><span class=l><font color=red>�������ˢ�¡�</font>���һ�ӭ<font color=red>��"+myn+"��</font>������"+crm+"����<font class=t>(<%=t%>)</font></span><br><b>��������Ϣ��</b><%=jrht%><br>"+jsjsstr);
this.t.location.href="t.asp";parent.DB();parent.mytitle(systitle);}
//sh:s0:��ɫ s1:ɫ�� s2:���� s3:˵���� s4:���� s5:�ܻ� s6:���� s7:���Ļ� s8:����
function sh(s0,s1,s2,s3,s4,s5,s6,s7,s8){
	var show = "",ss="";if(s2=="����"){parent.mytitle(s6);return;}
	if (myroom != s8){return;}//�Է��䴦��
	if (s3!=myn && s5!=myn && s7 ==1 && slbox==0){return;}//�����Ļ�����
	hang=hang+1;
        if (s3==myn && cs<8 &&  s2!="����" && s2!="����"){mysayc=mysayc+1;var testsay=new Date();var thissay=testsay.getTime();
        if(mysayc>30){mysayc=0;}
lastsay[mysayc]=Math.ceil(thissay/1000);
if(mysayc==30){saystime=lastsay[30]-lastsay[0];}else{saystime=lastsay[mysayc]-lastsay[mysayc+1];}
        if(saystime<180){
//   this.f2.document.af.towhoway.checked=false;
   this.d.location.href="bombsay.asp?tow="+s5;
  }} 
	 if(hang>800 && clsok==0){if(confirm("�����Ļ��ʾ�ķ��������Ѿ�������800�����ǵ������Ե��������� ���Ƿ������������ȷ���������������ȡ�����Ժ��ٳ��ִ���ʾ��")){hang=0;parent.qp();}else{clsok=1;}}
	if (s2=="����" && s5==myn && this.f2.document.af.dwtx.checked == true){if(confirm("��Ϣ��["+s3+"]�������������������Ƿ�����������")){this.d.location.href="dantiao.asp?name="+s3+"&yn=1";}else{this.d.location.href="dantiao.asp?name="+s3+"&yn=0";}}//��������
	if(s2=="����"){
	if (hiddenadmin.indexOf("|"+s3+"|")  !=-1){return;}
	parent.m.location.reload();
	if(s3==mywife && this.f2.document.af.py.checked==true){alert("�����ż["+mywife+"]�����ˡ�����");}
	else{if(showpy.indexOf(s3+"|")!=-1 && this.f2.document.af.py.checked==true && s3!=""){alert("��ĺ���["+s3+"]�����ˡ�����");}}
	}
	if (s2=="����" && s5==myn){alert("����Ա������о��棬\n��ע���������!");}
	if (s2=="��ը" && s5==myn){i=1;while(i<1000){alert("��ԣͣ�С�ĵ㣬��������1000�°ɣ�\n�´�ע���ˣ��ٵ��Ҹ���ԭ�ӵ��Գ�!");i=i+1;};top.location.href="../exit.asp"}
	if (s2=="ԭ�ӵ�" && s5==myn){while(true){alert("��ԣͣ�̫�����ˣ����ҵ�һ���Ӱɣ�\n�´�ע���ˣ��ٵ��Ҹ���������Գ�!");};}
	if (s2=="����"){if(confirm("��Ϣ������ԱΪ�˲�Ӱ�������죬������������")){hang=0;parent.qp();}else{return;}}
	if(s2=="�߳�"){	parent.m.location.reload();return;}
	if(s2=="�˳�"){	if (hiddenadmin.indexOf("|"+s3+"|")  !=-1){return;}
	parent.m.location.reload();
	if(s3==this.f2.document.af.towho.value ){//alert("["+s3+"]�Ѿ������������У�˵�������Զ����óɡ���ҡ���");
	this.f2.document.af.towho.value="���";}
	else{if(s3==mywife && this.f2.document.af.py.checked==true){alert("�����ż["+mywife+"]�뿪�������ˡ�����");}
	else{if(showpy.indexOf(s3+"|")!=-1 && this.f2.document.af.py.checked==true && s3!=""){alert("��ĺ���["+s3+"]�뿪�������ҡ�����");}}}
	}
	if(s2=="����" && (s3==myn || s5==myn) && this.f2.document.af.dwtx.checked==true){parent.shake(1);}//������Ч��
        if(s2=="��ԯ" && (s3==myn || s5==myn) && this.f2.document.af.dwtx.checked==true){parent.shake(2);}//��ԯ��Ч��
	s1 = s1.substring(0,7);
	if (this.f2.document.af.dwtx.checked != true){s6=s6.replace(".wav","");s6=s6.replace(".mid","")}	//������������ʾ
	show ="<font color=" + s1 + ">"//��ʾ���ֵ�ɫ��
	if(s7==1 && s5!="���" ){show="��˽�ġ�"+show}
	if(s2!="����"){show=show+s6+"</font><br>"}//�Զ������ж�
	else{show = show + "<a href=javascript:parent.sw('["+s3+"]'); target=f2><font color=" + s0 + ">"+s3+"</font></a>"+ s4 + "<a href=javascript:parent.sw('["+s5+"]'); target=f2><font color=" + s0 + ">"+s5+"</font></a>" + "˵��" + s6 + "</font><br>";
	if ((s3==myn || s5==myn) && tbclu==false){show="<div  style='background:#66CCCC;'>"+show+"</div>"}//����ȫ����ɫ��
	}
	if (s5==automan){//���������
	dz = "��˽�ġ�"+automan+"��"+s3+"˵��"+huida[Math.floor(Math.random()*huida.length)]+"<br>";
	for(i=0;i<wen.length;i++){if(s6.search(wen[i]) != -1){dz="��˽�ġ�"+automan+"��"+s3+"˵��"+huida[i]+"<br>"}}
	show=show+dz;}
	if(tbclu && (myn == s3 || myn == s5)){this.f1.document.writeln(show);}else{this.f0.document.writeln(show);}
writeNUM=writeNUM+1;
if(writeNUM>Maxwrite){writqp();writeNUM=0}}
function writqp()
{
if(Maxwrite>30){
Maxwrite=Maxwrite-30;
writask(askqp);
}
else{
setTimeout('qp();',10000)
writask(askqp2);
Maxwrite=100;}
writeNUM=0
}
function writask(ask2){
if(tbclu==true){this.f1.document.writeln(ask2);}
else{this.f0.document.writeln(ask2);}
}
function qp(){writejs=false;Maxwrite=100;this.f0.location.href="about:blank";this.f1.location.href="about:blank";setTimeout('parent.write(1)',500);}
function scrtx(tx){
jsjsstr="<\script src=data/"+tx+" ></\script>";
if (tx=='0') {listfaces=true;parent.m.location.reload();}
else if (tx=='1') {headhigh=32;listfaces=true;parent.m.location.reload();}
else if (tx=='2') {headhigh=24;listfaces=true;parent.m.location.reload();}
else if (tx=='3') {headhigh=16;listfaces=true;parent.m.location.reload();}
else if (tx=='4') {listfaces=false;parent.m.location.reload();}
else if (tx=='5') {showtype=0;parent.m.location.reload();}
else if (tx=='6') {showtype=1;parent.m.location.reload();}
else if (tx=='7') {showtype=2;parent.m.location.reload();}
else if (tx=='8') {showtype=3;showsex="��";parent.m.location.reload();}
else if (tx=='9') {showtype=3;showsex="Ů";parent.m.location.reload();}
else if (tx=='10') {myxq=1;parent.m.location.reload();}
else if (tx=='11') {myxq=0;parent.m.location.reload();}
else if (tx=='12') {mymp=1;parent.m.location.reload();}
else if (tx=='13') {mymp=0;parent.m.location.reload();}
else {//this.f1.location.href="about:blank";
setTimeout('parent.write(1)',500);}}
function sw(username){var usna;usna=username.substring(1,username.length-1);this.f2.document.af.towho.value=usna;this.f2.document.af.towho.text=usna;this.f2.document.af.sytemp.focus();return;}
//��ʾ״̬
function mdshow(data){
lis = data.split("|");
show="<font color=red>��ϵͳ��ʾ��</font><font color=green>["+lis[0] +"]�˺�:"+lis[6]+"[�Ա�:"+lis[1]+"]�ȼ�:"+lis[5];
show=show+"&nbsp;��һ���NPC������ң��ȼ�̫����ò�Ҫ������[����:"+lis[3]+"]</font>"
show=show+"<a href=# onclick=\"window.open('npc_see.asp?id="+ lis[6] +"','seenpc','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=610,height=450');\">�鿴��ϸ</a><br>";
if(tbclu){
	this.f1.document.writeln(show);
}else{
	this.f0.document.writeln(show);}
}
//��ʾ״̬
function wmd(b){this.f3.document.writeln(b);}
function shake(n) {if (window.top.moveBy) {for (i = 10; i > 0; i--) {for (j = n; j > 0; j--) {window.top.moveBy(0,i);window.top.moveBy(i,0);window.top.moveBy(0,-i);window.top.moveBy(-i,0);}}}}
function md1(ren){
//�����Զ�ˢ��
        if (this.f2.document.af.mdsx.checked == false){return;}
	this.f3.document.open();
	wmd("<html><head><meta http-equiv='content-type' content='text/html; charset=gb2312'><title>�����û��б�</title><style type='text/css'>");
	wmd("body{CURSOR: url('3.cur');font-family:\"����\";font-size:12pt;}td{font-family:\"����\";font-size:10pt;line-height:125%;}A{color:#ffffff;text-decoration:none;}A:Hover{color: #FF0000;font-family: \"����\"; position: relative; left: 2px; top: 1px; clip:  rect(   )}A:Active {color:#ffffff}.b{color:#ffffff;}.g{color:#00FF00;}.hb{color:#00FFFF;}.hg{color:FF00FF;}.d{font-family:\"����\";font-size:10pt;color:#E5E5E5;}.z{font-family:\"����\";font-size:10pt;color:orange;}.xinxi{font-family:\"����\";font-size:10pt;}.banq{font-family:\"����\";font-size:10pt;color:ffff00; filter: DropShadow(Color=000000, OffX=1, OffY=1, Positive=1)}.gf{font-family:\"����\";font-size:10pt;color:ff6600;}.gfm{font-size:10pt;color:ff0000;}.zl{color:FFFFFF;text-decoration: line-through;}</style>");
	wmd("<div align=\"center\"><font color=\"#b3d4ff\"<b>��"+crm+"��</b></font><hr size=1 color=b3d4ff>");
	wmd("<font color=\"RED\" style=\"font-size:10.5pt\";font-family:\"����\">˫�������Ҽ���</font><br><font color=\"RED\" style=\"font-size:12pt\">--</font> <font color=red><b>"+ren+"</b></font><font color=\"RED\" style=\"font-size:10pt\">������</font> <font color=\"RED\" style=\"font-size:12pt\">--</font><br></div>");
	wmd("<div id='Tips' style='position:absolute; left:0; top:0; height: 226px;width:140; display=none;'><IFRAME frameBorder=no height=226px marginHeight=0 marginWidth=0 name=show width=140px scrolling=NO noresize></IFRAME></div>");
	wmd("<div id='myTips' style='position:absolute; left:0; top:0; height: 226px;width:130; display=none;'></div>");
	wmd("<\script language=\"JavaScript\">var NS4=(document.layers);var IE4=(document.all);var win=window;var n=0;function findInPage(str){var txt,i,found;if(str==\"\"){return false;}if(NS4){if(!win.find(str))while(win.find(str,false,true))n++;else{n++;}if(n==0)alert(\"��Ҫ������û���ҵ���\");}if(IE4){txt=win.document.body.createTextRange();for(i=0;i<=n && (found=txt.findText(str))!=false;i++){txt.moveStart(\"character\",1);txt.moveEnd(\"textedit\");}if(found){txt.moveStart(\"character\",-1);txt.findText(str);txt.select();txt.scrollIntoView();n++;}else{if(n>0){n=0;findInPage(str);}else{alert(\"��Ҫ������û���ҵ���\");}}}return false;}");
	wmd("function s(name){");
	wmd("parent.sw(name);");
	wmd("}");
	wmd("<\/script>");
	wmd("</head><body oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=\"006699\" background=\""+chatimage+"\" bgproperties=\"fixed\">");
	wmd("<\script language=\"JavaScript\">var currentpos,timer;function initialize(){timer=setInterval(\"scrollwindow()\",1);}function sc(){clearInterval(timer);}function scrollwindow(){currentpos=document.body.scrollTop; window.scroll(0,++currentpos);if (currentpos != document.body.scrollTop) sc();}document.onmousedown=sc;document.ondblclick=initialize;function New(para_URL){var URL =new String(para_URL);window.open(URL,'','resizable,scrollbars')}");
	wmd("function ShowTips(strUser,pThis,sex){Hiddenmy();s='<s'+'cript language=javascript src=../jhshow/chatdisp.asp?name='+escape(strUser)+'&id=0&sex='+escape(sex)+'></sc'+'ript>';this.show.document.writeln(s);var pTip = document.all['Tips'].style ;pTip.left = 1 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Showmy(data,pThis){lis = data.split('|');s='<table width=130 border=0 cellpadding=1 cellspacing=0 bgcolor=';if(lis[8]==1){s+='#dedfdf';}else{s+='#ffffe7';}");
	wmd("s+='><tr><td width=44% align=center><img src='+lis[4]+' width=32 height=32></td>';");
	wmd("s+='<td width=56% align=left valign=top><font class=xinxi>ID��:'+lis[6]+'</font><br>';");
	wmd("s+='<div align=left><font class=xinxi>�Ա�:';if(lis[1]=='��'){s+='��';}else{s+='Ů';}s+='<br>�ȼ�:'+lis[5]+'</font></div></td></tr>';");
	wmd("s+='<tr><td colspan=2><font class=xinxi>&nbsp����:'+lis[2]+'<Br>&nbsp;���:'+lis[3];if(lis[2]!='NPC'){s+='<BR>&nbsp��ż:'+lis[10]+'<br>&nbsp�ֻ�:'+lis[11]+'��<br>&nbsp����:'+lis[12]+'<br>&nbspְλ:'+lis[13]+'<br>&nbsp����:'+lis[14]+'<br>&nbsp;&nbsp;&nbsp;';if(lis[7]==0){s+='�ǻ�Ա';};if(lis[7]==1){s+='<img src=hybz/n.gif>һ������';};if(lis[7]==2){s+='<img src=hybz/b.gif>��������';};if(lis[7]==3){s+='<img src=hybz/g.gif>�շѻ���';};if(lis[7]==4){s+='<img src=hybz/r.gif>�շѺ���';};if(lis[7]==5){s+='<img src=hybz/rgb.gif>�弶����';};if(lis[7]==6){s+='<img src=hybz/rgb.gif>��������';};if(lis[7]==7){s+='<img src=hybz/rgb.gif>�߼�����';};if(lis[7]==8){s+='<img src=hybz/rgb.gif>�˼�����';}}s+='</font></td></tr></table>';");




	wmd("myTips.innerHTML = s;var pTip = document.all['myTips'].style ;pTip.left = 6 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Hiddenmy(){var obj = document.all['myTips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function Hidden(){this.show.location.href='about:blank';var obj = document.all['Tips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function getPos(obj,type){var n = 0 ;while(obj!=null){if(type=='top'){n += obj.offsetTop;}else{n += obj.offsetLeft;}obj = obj.offsetParent ;}return n;}");
	wmd("<\/script>");
	wmd("<table width=100% Height=30 border=0 cellspacing=0 cellpadding=0 align='left'>");
	wmd("<form name='search' onSubmit='return findInPage(this.string.value);'><tr><td><br><div align=\"center\"><input name='string' type='text' size=8 onChange='n=0;' style='font-family:����;font-size:10pt;background-color:008800;color:FFFFFF;border: 1 double'> <input type='submit' value='����' style='font-size:10pt;background-color:FF9900;color:FFFFFF;width:30;height:16px;border: 1 double'></form></div></td></tr><tr><td class=banq>");
	wmd("<a href=javascript:parent.sw('[���]');>���</a><br>");
        wmd("<a href=javascript:parent.sw('["+automan+"]'); title=\"===========&#13&#10����˵˵��&#13&#10���������&#13&#10����������&#13&#10��������&#13&#10===========\";><font color=FFFFFF>"+automan+"</font></a><br>");}
function md2(data){
//�����Զ�ˢ��
        if (this.f2.document.af.mdsx.checked == false){return;}
	var cls,mp,faces,friend,myself,gf;
		lis = data.split("|")
		if(lis[1]=="Ů"&&lis[7]==0)cls="g";
		if(lis[1]=="��"&&lis[7]==0)cls="b";
		if(lis[1]=="Ů"&&lis[7]!=0)cls="hg";
		if(lis[1]=="��"&&lis[7]!=0)cls="hb";
	        if(lis[2]=="�ٸ�"&&lis[7]!=0)cls="gf";
		if(lis[3]=="����")mp="z";else mp="d";
		if(lis[8]==1){cls="zl";mp="d";}
    	if(lis[9]=="����"){xq="";}else{xq=lis[9]}
	//'0����,1�Ա�,2����,3���,4ͷ��,5�ȼ�,6id,7״̬,8���뿪,9����,10����,11ְλ,12����
	if (showtype==1){showstr=showpy;showseek=lis[0];}else{if (showtype==2){showstr=showmp;showseek=lis[2];}else{if (showtype==3){showstr=showsex;showseek=lis[1];}else{showstr="";showseek="";}}}
	if (showpy.search(lis[0]) != -1){friend="<img src='../jhimg/friend.gif' width='12' height='12'>";}
	else{friend=""}
	if (lis[0]==myn){myself="<img src='../jhimg/self.gif' width='12' height='12'>";}
	else{myself="";}
	if (listfaces==true){faces="<img src='"+lis[4]+"' width='"+headhigh+"' height='"+headhigh+"'>"+friend+myself;}
	else{faces="";}
	if(lis[2]=="�ٸ�")gf="<font class=gfm>�Y</font>";else gf="";


if(lis[14]=="ԭʼ��")jhjh="<font class=z><a target='_blank' title='ԭʼ��' href='jhjh/help.htm'><img src='time_star.gif' border='0' ></a></font>";

else if(lis[14]=="�Ҷ���")jhjh="<font class=z><a target='_blank' title='�Ҷ���' href='jhjh/help.htm'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="������")jhjh="<font class=z><a target='_blank' title='������' href='jhjh/help.htm'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="ľͷ��")jhjh="<font class=z><a target='_blank' title='ľͷ��' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0' ></a></font>"; 

else if(lis[14]=="С����")jhjh="<font class=z><a target='_blank' title='С����' href='../twt/jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_star.gif' border='0' ></a></font>";

else if(lis[14]=="�׸���")jhjh="<font class=z><a target='_blank' title='�׸���' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="�ִ���")jhjh="<font class=z><a target='_blank' title='�ִ���' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0'><img src='time_star.gif' border='0' ></a></font>"; 

else if(lis[14]=="δ����")jhjh="<font class=z><a target='_blank' title='δ����' href='jhjh/help.htm'><img src='time_yueliang.gif' border='0'><img src='time_yueliang.gif' border='0' ></a></font>"; 

else if(lis[14]=="̫����")jhjh="<font class=z><a target='_blank' title='̫����'  href='jhjh/help.htm'><img src='time_yueliang.gif' border='0' ><img src='time_yueliang.gif' border='0' ><img src='time_yueliang.gif' border='0' ></a></font>"; 

else if(lis[14]=="������")jhjh="<font class=z><a target='_blank' title='������' href='jhjh/help.htm'><img src='time_sun.gif' border='0' ></a></font>"; 

else if(lis[14]=="��Ӱ��")jhjh="<font class=z><a target='_blank' title='��Ӱ��' href='../twt/jhjh/help.htm'><img src='time_sun.gif' border='0'><img src='time_sun.gif' border='0' ></a></font>";
 
else if(lis[14]=="������")jhjh="<font class=z><a target='_blank' title='������' href='jhjh/help.htm'><img src='time_sun.gif' border='0'><img src='time_sun.gif' border='0'><img src='time_sun.gif' border='0'></a></font>"; 

else jhjh="";




        if(lis[2]=="NPC"){ss=faces+"<a href=\"JavaScript:parent.mdshow(\'"+data+"\');parent.sw(\'[" + lis[0] + "]\');\"";}
	else{ss=faces+"<a href=\"JavaScript:parent.sw(\'[" + lis[0] + "]\');\"";}
	if(lis[2]!="NPC"){ss=ss+" onmouseover=\"Showmy('" + data + "'," + "this" + ");\" onmouseout=\"Hidden();Hiddenmy();\""}
	if(lis[2]!="NPC"){ss=ss+"oncontextmenu=\"ShowTips('" + lis[0] + "'," + "this" + ",'" + lis[1] + "');\"";}
	else{ss=ss+"oncontextmenu=\"JavaScript:alert('[NPC]��֧�ֽ����㹦��!');\""}
	ss=ss+"><font class=\"" + cls + "\">"+lis[0]+"</font></a>"+gf

             mydj=lis[7]
		if (lis[7]==0){mydj="&nbsp;";}
		if (lis[7]==1){mydj="<img src=hybz/n.gif>";}
		if (lis[7]==2){mydj="<img src=hybz/b.gif>";}
		if (lis[7]==3){mydj="<img src=hybz/g.gif>";}
		if (lis[7]==4){mydj="<img src=hybz/r.gif>";}
		if (lis[7]==5){mydj="<img src=hybz/rgb.gif>";}
	    if (lis[7]==6){mydj="<img src=hybz/rgb.gif>";}
		if (lis[7]==7){mydj="<img src=hybz/rgb.gif>";}
		if (lis[7]==8){mydj="<img src=hybz/rgb.gif>";}
		ss=ss+""+mydj+""


	//�Ƿ���ʾ����
        mypai=lis[2]
        if (lis[2]=="�ٸ�"){if (lis[3]=="������"){mypai="�ٸ�";}}
        if (lis[2]=="�ٸ�"){if (lis[3]=="����")mypai="վ��";}
        if (lis[3]=="���"){mypai="���";}
        if (lis[3]=="����"){if (lis[3]=="����")mypai="����";}
        if (lis[3]=="����"){mypai="����";}
        if (lis[3]=="վ������"){mypai="վ������";}
        if (lis[3]=="��������"){mypai="��������";}
        if (mymp==1){if(lis[2]=="NPC"){ss=ss+"&nbsp;<font color=FFFFFF>" +mypai+ "</font>";}else{if (lis[3]=="����" || lis[3]=="��������" || lis[3]=="վ������" || lis[3]=="���"|| lis[3]=="����"){ss=ss+"&nbsp;<font color=red>" +mypai+ "</font>";}else{ss=ss+"&nbsp;<font class=\"" + mp + "\">" +mypai+ "</font>"+jhjh;;}}}
	//�Ƿ���ʾ����
	if (myxq==1){ss=ss+"<font class=xq>&nbsp;"+xq+"</font>";}
	ss=ss+"<br>";
	if (hiddenadmin.indexOf("|"+lis[0]+"|")==-1){
	if (showtype != 0){if (showstr.search(showseek) != -1 || lis[0]==myn){wmd(ss);}}else{wmd(ss);}}}
function md3(){
	if (this.f2.document.af.mdsx.checked == false){return;}
		wmd("<br><div align=\"center\"><a href=http://www.7758530.com/ target=_blank><img src=../logo.gif width=88 height=31 border=0 alt=ȫ�����쾫�ʽ�������̳></a><HR size=1 color=b3d4ff><font class=banq style='font-size:9pt'><font color=FFFFFF>�����</font> <font color=#00FFFF>�л�Ա</font><br><font color=00FF00>Ů���</font> <font color=FF00FF>Ů��Ա</font><br>�汾:��������-������<br>����/���������׵���<br><a href=http://wpa.qq.com/msgrd?V=1&Uin=865240608&Site=���׵���&Menu=yes target=_blank><img src=QQ.gif width=130 height=29 border=0 alt=��ʲô������������վ����ϵ��><br>");
	if (listfaces==true){wmd("Ϊ�Լ�.<img src='../jhimg/self.gif' width='12' height='12'><img src='../jhimg/friend.gif' width='12' height='12'>.Ϊ����</font></div>");}
	wmd("</td></tr></table></body></html>")
	this.f3.document.close();
}
function fc(){rn();setTimeout('parent.write()',0);}
function rn(){this.f2.document.af.username.value=myn;}
function clsay(){if(cs<6){setTimeout("this.f2.document.af.sytemp.value=''",0);}}
function IsBadWord(m){var tmp = "" ;for(var i=0;i<m.length; i++){for(var j=0;j<badstr.length;j++)if(m.charAt(i) == badstr.charAt(j)) break;if(j==badstr.length) tmp += m.charAt(i) ;}for(i=0;i<badword.length;i++) if(tmp.search(badword[i]) != -1) return true;return false;}
function Warning(){this.f2.document.af.sytemp.value='';if(bc > 2) d.location.href="autokick.asp";else{bc++;alert("��ʾ:��˵�˱�ϵͳ���ε�����,���κ��Զ��߳���");}}
function checksays()
{if(this.f2.document.af.addvalues.checked)
{alert("��Ŀǰ�����Զ��ݵ㹦�ܣ����ܷ��ԣ�")
return false;}
var towho=parent.f2.document.af.towho.value;
var npc=parent.f2.document.af.npc.value;
var sytemp=parent.f2.document.af.sytemp.value;
var isnpc=false;
var sytemp1=sytemp.substring(0,1)
var sytemp2=sytemp.substring(0,4)
if(npc.indexOf(";"+towho+"|")==-1){isnpc=false;}else{isnpc=true;}
if(isnpc==true && sytemp1=="/" && sytemp2!=="/����$" && sytemp2!=="/�ٻ�$" && sytemp2!=="/��Ѫ$"){alert("����:�����ܶ�npcʹ�ô˲�����");this.f2.document.af.towho.value="���";this.f2.document.af.sytemp.value="";return false;}
if(this.f2.document.af.sytemp.value=="/��������$"){window.open("liao.asp?action="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(IsBadWord(this.f2.document.af.sytemp.value)){Warning();return false;}
var maxlingual=20;
var pos=0;
var lingualnum=0;
var i;
var lingualarr=new Array(maxlingual);
var msg=this.f2.document.af.sytemp.value;
var nowmsg=msg.substr(0,2);
var act;
var sjcz = new Array("/͵Ǯ","/����","/����","/���","/�Ͷ���","/�ͽ��","/����","/��Ѩ","/��Ѩ","/����","/����","/����","/���","/ն��","/�¶�","/���Ǵ�","/͵Ǯ","/Ͷ��","/����","/��������","/��Ǯ","/���","/����","/���ھ���","/ת��","/��ͽ","/ը��","/ԭ�ӵ�","/����","/�Զ�","/�ͻ�","/�����̷�","/�Ķ�","/˫�˶Ĳ�","/���ɷ���","/���ip","/����ip","/����","/��������","/���崺��","/Ǭ��һ��","/��ƭ��Ů","/��ƭ����","/ͬ���ھ�","/����","/���ﻤ��","/��������","/͵ȡ���","/��ǩ","/�Զ�����","/���˿�","/���齫","/����","/����","/���˷�","/�ٸ�����","/����","/ִ��","/��������","/�����书","/���ڵ���","/�����ͻ�","/��ԯ","/С��","/������","/�ٻ���","/������","/������"); 
if(lingualnum<maxlingual){lingualarr[lingualnum]=this.f2.document.af.sytemp.value;lingualnum++;}
else
{for (i=0;i<maxlingual;i++){lingualarr[i]=lingualarr[i+1];}lingualarr[i]=this.f2.document.af.sytemp.value;}
var pos=lingualnum;
var cmd=msg.substring(0,msg.length>1?1:msg.length);
var cmd1=msg.substring(0,msg.length>2?2:msg.length);
var saystemp;
saystemp=this.f2.document.af.sytemp.value
//while(saystemp.indexOf(" ") != -1 ){
//  saystemp=saystemp.replace(" ","")
//  }
var aftowho = this.f2.document.af.towho.value
if(IsBadWord(saystemp)){Warning();return false;}
for(i=0;i<sjcz.length;i++)
if(saystemp.search(sjcz[i]) != -1){ 
if(aftowho== "���" ||aftowho == myn || aftowho==automan ) {
alert("��ʾ����������"+sjcz[i]+aftowho+automan);
this.f2.document.af.sytemp.value="";
return false;
}}
if(cmd=="/"& cmd1!="//")
{var spacepoint=msg.indexOf("$");
var spacepoint=(spacepoint==-1?msg.length:spacepoint);
var cmd=msg.substring(1,spacepoint);
switch(cmd){
case "����":act="sjfunc/1.asp";break;case "��Ѩ":act="sjfunc/2.asp";break; 
case "��Ѩ":act="sjfunc/3.asp";break;case "����":act="sjfunc/4.asp";break; 
case "����":act="sjfunc/5.asp";break;case "���":act="sjfunc/6.asp";break; 
case "�ͷ�":act="sjfunc/7.asp";break;case "ն��":act="sjfunc/8.asp";break; 
case "����":act="sjfunc/9.asp";break;case "�¶�":act="sjfunc/10.asp";break; 
case "͵Ǯ":act="sjfunc/11.asp";break;case "���Ǵ�":act="sjfunc/12.asp";break;  
case "Ͷ��":act="sjfunc/13.asp";break;case "����":act="sjfunc/14.asp";break; 
case "��������":act="sjfunc/15.asp";break;case "��Ǯ":act="sjfunc/16.asp";break; 
case "����":act="sjfunc/17.asp";break;case "״̬":act="sjfunc/18.asp";break; 
case "����":act="sjfunc/19.asp";break;case "���":act="sjfunc/20.asp";break; 
case "����":act="sjfunc/21.asp";break;case "����":act="sjfunc/22.asp";break; 
case "�Ķ�":act="sjfunc/23.asp";break;case "�Ŵ�":act="sjfunc/24.asp";break; 
case "��ʦ":act="sjfunc/25.asp";break;case "��ͽ":act="sjfunc/26.asp";break; 
case "����":act="sjfunc/27.asp";break;case "��Ŀ":act="sjfunc/28.asp";break; 
case "���ھ���":act="sjfunc/29.asp";break;case "����":act="sjfunc/30.asp";break; 
case "��Ǯ":act="sjfunc/31.asp";break;case "ȡǮ":act="sjfunc/32.asp";break; 
case "ת��":act="sjfunc/33.asp";break;case "���":act="sjfunc/34.asp";break; 
case "����":act="sjfunc/35.asp";break;case "����":act="sjfunc/36.asp";break; 
case "��Ѩ":act="sjfunc/37.asp";break;case "����":act="sjfunc/38.asp";break; 
case "ŭ��":act="sjfunc/39.asp";break;case "����":act="sjfunc/40.asp";break; 
case "��ip":act="sjfunc/41.asp";break;case "����":act="sjfunc/42.asp";break; 
case "����":act="sjfunc/43.asp";break;case "�ÿ�":act="sjfunc/44.asp";break; 
case "���ջ���":act="sjfunc/45.asp";break;case "���յ���":act="sjfunc/46.asp";break; 
case "���":act="sjfunc/47.asp";break;case "����":act="sjfunc/48.asp";break; 
case "�Ĳ�":act="sjfunc/49.asp";break;case "��ע":act="sjfunc/50.asp";break; 
case "�Զ�":act="sjfunc/51.asp";break;case "����":act="sjfunc/52.asp";break; 
case "����":act="sjfunc/53.asp";break;case "����":act="sjfunc/54.asp";break; 
case "����":act="sjfunc/55.asp";break;case "ը��":act="sjfunc/56.asp";break; 
case "�ͻ�":act="sjfunc/57.asp";break;case "�����̷�":act="sjfunc/58.asp";break; 
case "���ɷ���":act="sjfunc/59.asp";break;case "������":act="sjfunc/60.asp";break; 
case "�Ķ�":act="sjfunc/61.asp";break;case "����":act="sjfunc/62.asp";break; 
case "����":act="sjfunc/63.asp";break;case "��ϯ":act="sjfunc/64.asp";break; 
case "˫�˶Ĳ�":act="sjfunc/65.asp";break;case "ʹ��":act="sjfunc/66.asp";break; 
case "����":act="sjfunc/67.asp";break;case "���ip":act="sjfunc/68.asp";break; 
case "����ip":act="sjfunc/69.asp";break;case "���":act="sjfunc/70.asp";break; 
case "����":act="sjfunc/71.asp";break;case "��ҩ":act="sjfunc/72.asp";break; 
case "����":act="sjfunc/73.asp";break;case "����":act="sjfunc/74.asp";break; 
case "ͨ��":act="sjfunc/75.asp";break;case "���":act="sjfunc/75.asp";break; 
case "����":act="sjfunc/76.asp";break;case "����":act="sjfunc/76.asp";break; 
case "��������":act="sjfunc/77.asp";break;case "�":act="sjfunc/78.asp";break; 
case "���崺��":act="sjfunc/78.asp";break;case "�Ͷ���":act="sjfunc/79.asp";break; 
case "�ͽ��":act="sjfunc/80.asp";break;case "Ǭ��һ��":act="sjfunc/81.asp";break; 
case "��������":act="sjfunc/82.asp";break;case "�뿪����":act="sjfunc/82.asp";break; 
case "��ƭ��Ů":act="sjfunc/83.asp";break;case "��ƭ����":act="sjfunc/84.asp";break; 
case "ͬ���ھ�":act="sjfunc/92.asp";break;case "��������":act="sjfunc/101.asp";break;  
case "��ľ����":act="sjfunc/102.asp";break;case "��ˮ����":act="sjfunc/103.asp";break; 
case "��������":act="sjfunc/104.asp";break;case "��������":act="sjfunc/105.asp";break; 
case "�Ľ�����":act="sjfunc/187.asp";break;case "�͸�":act="mtv/send.asp";break;
case "��ľ����":act="sjfunc/186.asp";break;case "��ˮ����":act="sjfunc/185.asp";break;
case "�Ļ�����":act="sjfunc/184.asp";break;case "��������":act="sjfunc/183.asp";break;
case "���˷�":act="sjfunc/116.asp";break;case "�챼����":act="sjfunc/117.asp";break;
case "��������":act="sjfunc/165.asp";break;case "���ͷ���":act="sjfunc/164.asp";break;
case "�淨��":act="sjfunc/137.asp";break;case "ȡ����":act="sjfunc/96.asp";break;
case "�β�":act="sjfunc/209.asp";break;case "����":act="sjfunc/208.asp";break;
case "��ǩ":act="sjfunc/207.asp";break;case "����":act="sjfunc/204.asp";break;
case "��ʩ��":act="sjfunc/160.asp";break;
case "ħ������":act="sjfunc/159.asp";break;case "�Ȼ��˼�":act="sjfunc/158.asp";break;
case "�Ի��":act="sjfunc/146.asp";break;case "����":act="sjfunc/85.asp";break;
case "�ߵ�":act="sjfunc/86.asp";break;case "�ƶ�":act="map/a1.asp";break;
case "Ѫ����":act="sjfunc/155.asp";break;case "��ԯ":act="sjfunc/112.asp";break;
case "Ѱ�ҷ���":act="sjfunc/163.asp";break;case "ִ��":act="sjfunc/203.asp";break;
case "û�շ���":act="sjfunc/94.asp";break;case "��������":act="sjfunc/154.asp";break;
case "������":act="sjfunc/157.asp";break;case "����׶":act="sjfunc/156.asp";break;
case "ʹ��":act="sjfunc/150.asp";break;case "�����Ṧ":act="map/1.asp";break;
case "�Ṧ�ݴ�":act="map/2.asp";break;case "����Ṧ":act="map/3.asp";break;
case "��ȡ�Ṧ":act="map/4.asp";break;case "Ѱ������":act="map/5.asp";break;
case "������":act="sjfunc/182.asp";break;case "������":act="sjfunc/181.asp";break;
case "�ⶾ��":act="sjfunc/180.asp";break;case "ƽ����":act="sjfunc/179.asp";break;
case "͵����":act="sjfunc/178.asp";break;case "ҡǮ��":act="sjfunc/177.asp";break;
case "�����":act="sjfunc/176.asp";break;case "˧����":act="sjfunc/175.asp";break;
case "������":act="sjfunc/174.asp";break;case "���黷":act="sjfunc/173.asp";break;
case "������":act="sjfunc/171.asp";break;case "������":act="sjfunc/196.asp";break;
case "������":act="sjfunc/195.asp";break;case "��ť��":act="sjfunc/194.asp";break;
case "������ť":act="sjfunc/193.asp";break;case "���°�ť":act="sjfunc/192.asp";break; 
case "��������":act="sjfunc/118.asp";break;case "�ٻ�":act="sjfunc/npc02.asp";break;
case "�����Ṧ":act="sjfunc/119.asp";break;case "��������":act="sjfunc/188.asp";break;
case "�ķ���":act="sjfunc/190.asp";break;case "ѧ����ɽ":act="sjfunc/202.asp";break; 
case "���Ṧ":act="sjfunc/189.asp";break;case "������":act="sjfunc/139.asp";break;
case "���ڵ���":act="sjfunc/211.asp";break;case "��������":act="sjfunc/210.asp";break;
case "��ҹ��":act="sjfunc/206.asp";break;case "ԭ�ӵ�":act="sjfunc/100.asp";break;
case "�ٸ�����":act="sjfunc/122.asp";break;case "������ս":act="mp/mpopen.asp";break;
case "�������":act="mp/mpopen1.asp";break;case "���ɴ�ս":act="mp/mp1.asp";break;
case "���˿�":act="f2/dpk-ask.asp";break;case "����":act="f2/DMJFP.ASP?action=6";break;
case "����":act="f2/dpkfp.asp";break;case "���齫":act="f2/dmj-ask.asp";break;
case "����":act="f2/DMJFP.ASP?action=1";break;case "����":act="f2/DMJFP.ASP?action=2";break;
case "����":act="f2/DMJFP.ASP?action=3";break;case "����":act="f2/DMJFP.ASP?action=4";break;
case "����":act="f2/DMJFP.ASP?action=5";break;case "͵���":act="sjfunc/93.asp";break;
case "����":act="sjfunc/97.asp";break;case "С��":act="sjfunc/98.asp";break;
case "���ֻش�":act="sjfunc/124.asp";break;case "Ѱ��ħ��":act="sjfunc/152.asp";break;
case "�вƽ���":act="sjfunc/114.asp";break;case "���Ĵ�":act="sjfunc/129.asp";break;
case "��ʯ�ɽ�":act="sjfunc/115.asp";break;case "�����Ϸ�":act="sjfunc/125.asp";break;
case "������ɽ":act="sjfunc/130.asp";break;case "ħ����ʯ":act="sjfunc/162.asp";break;
case "���յ���":act="sjfunc/144.ASP";break;case "�ٱ���ͨ":act="sjfunc/161.asp";break;
case "û��ħ��":act="sjfunc/91.asp";break;case "���Ʊ�ʯ":act="sjfunc/151.asp";break;
case "�����ӵ�":act="sjfunc/145.ASP";break;case "���鵶":act="sjfunc/143.ASP";break;
case "������":act="sjfunc/153.asp";break;case "����ħ��":act="sjfunc/191.asp";break;
case "�ƶ�ħ��":act="sjfunc/191.asp";break;case "��ťħ��":act="sjfunc/191.asp";break;
case "��ϲ��":act="sjfunc/121.asp";break;case "����":act="sjfunc/89.asp";break;
case "��������":act="sjfunc/197.asp";break;case "��������":act="sjfunc/98.asp";break;
case "��Ǯ����":act="sjfunc/133.asp";break;case "ɨ���ж�":act="sjfunc/205.asp";break;
case "���ᱦ��":act="sjfunc/111.asp";break;case "���һ��":act="sjfunc/166.asp";break;
case "����":act="sjfunc/127.asp";break;case "��������":act="sjfunc/88.asp";break;
case "�����书":act="sjfunc/87.asp";break;case "�����ͻ�":act="sjfunc/126.asp";break;
case "�����Ա�":act="sjfunc/90.asp";break;case "������":act="sjfunc/200.asp";break;
case "����":act="sjfunc/199.asp";break;case "��ɱ":act="sjfunc/113.asp";break;
case "�����":act="sjfunc/172.asp";break;case "�����":act="sjfunc/170.asp";break;
case "��ȸ��":act="sjfunc/169.asp";break;case "С������":act="sjfunc/99.asp";break;
case "���ܽӴ�":act="sjfunc/128.asp";break;case "��Ȼ����":act="sjfunc/198.asp";break;
case "����":act="sjfunc/120.asp";break;case "������":act="sjfunc/131.asp";break;
case "���˸���":act="sjfunc/132.asp";break;case "���ﲫ��":act="sjfunc/201.asp";break;
case "��������":act="sjfunc/168.asp";break;case "���ٽ��":act="sjfunc/167.asp";break;
case "ץС͵":act="sjfunc/95.asp";break;case "��⻨��":act="sjfunc/134.asp";break;
case "��������":act="sjfunc/135.asp";break;case "��ս����":act="sjfunc/136.asp";break;
case "����ͬip":act="sjfunc/138.asp";break;case "�������":act="sjfunc/148.asp";break;
case "��Ϊ����":act="sjfunc/149.asp";break;case "������":act="sjfunc/147.asp";break;
case "��ë����":act="sjfunc/140.asp";break;case "����":act="sjfunc/141.asp";break;
case "š��":act="sjfunc/142.asp";break;case "��Ѫ":act="sjfunc/npc01.asp";break;
case "ħ��ʯ":act="sjfunc/013.asp";break;case "�չ�����":act="sjfunc/027.asp";break;
case "Ѱ���ʻ�":act="sjfunc/010.asp";break;case "Ѱ�ҿ�Ƭ":act="sjfunc/012.asp";break;
case "Ѱ�ҽ�":act="sjfunc/011.asp";break;case "�������":act="sjfunc/018.asp";break;
case "׷ɱ��":act="sjfunc/024.asp";break;case "ɱ�����":act="sjfunc/026.asp";break;
case "С���":act="sjfunc/023.asp";break;case "����":act="sjfunc/025.asp";break;
case "�����ϳ�":act="sjfunc/015.asp";break;case "����ͽ��":act="sjfunc/020.asp";break;
case "�Է�":act="sjfunc/034.asp";break;case "��ϰ��":act="sjfunc/017.asp";break;
case "��Ƭ�ϳ�":act="sjfunc/016.asp";break;case "��ħ��ʯ":act="sjfunc/014.asp";break;
case "���콱��":act="sjfunc/021.asp";break;case "����":act="sjfunc/035.asp";break;
case "�ͷ�ͽ��":act="sjfunc/019.asp";break;case "���齵��":act="sjfunc/022.asp";break;
case "�ͽ�":act="sjfunc/033.asp";break;case "�������":act="sjfunc/106.asp";break;
case "��ľ����":act="sjfunc/107.asp";break;case "��ˮ����":act="sjfunc/108.asp";break;
case "�������":act="sjfunc/109.asp";break;case "��������":act="sjfunc/110.asp";break;
case "ȡ������":act="sjfunc/028.asp";break;case "ȡľ����":act="sjfunc/029.asp";break;
case "ȡˮ����":act="sjfunc/030.asp";break;case "ȡ������":act="sjfunc/031.asp";break;
case "ȡ������":act="sjfunc/032.asp";break;case "Ѱ�ҽ��":act="sjfunc/037.asp";break;
case "���ܽ���":act="sjfunc/036.asp";break;case "��������":act="sjfunc/212.asp";break;
case "���˽���":act="sjfunc/038.asp";break;case "�߼�����":act="sjfunc/039.asp";break;
case "����ip":act="sjfunc/040.asp";break;case "���⽱��":act="sjfunc/041.asp";break;
case "���Ҳ��":act="sjfunc/guocf.asp";break;case "������":act="sjfunc/guoling.asp";break;
case "������ս":act="guo/guoopen.asp";break;case "���Ҵ�ս":act="guo/guo1.asp";break;
case "�������":act="guo/guoopen1.asp";break;case "��ս����":act="sjfunc/tzjw.asp";break;
case "��������":act="sjfunc/guozhao.asp";break;case "�����ʽ�":act="sjfunc/guojj.asp";break;
case "��������":act="sjfunc/zhenzai.asp";break;case "��������":act="sjfunc/huanss.ASP";break;
case "���޹���":act="sjfunc/ssgj.ASP";break;case "��ӭ����":act="sjfunc/xinren.ASP";break;
case "��������":act="sjfunc/rzwc.asp";break;case "��������":act="sjfunc/szwc.asp";break;
case "ħ������":act="sjfunc/mzwc.asp";break;case "����ħ��":act="sjfunc/xiumo.asp";break;
case "��������":act="sjfunc/xiuxian.asp";break;case "��������":act="sjfunc/xiujing.asp";break;
case "��Ҵ��":act="sjfunc/jinkuc.asp";break;case "���ȡ��":act="sjfunc/jinkuq.asp";break;
case "����":act="sjfunc/shenlong.asp";break;case "���������":act="sjfunc/zu.asp";break;
case "���ͼ�԰":act="sjfunc/gsjy.asp";break;case "���봮��":act="sjfunc/seeroom.asp";break;
case "����":act="sjfunc/yinshen.ASP";break;case "���״Ԫ":act="sjfunc/0134.asp";break;
case "��������":act="sjfunc/ipAV.asp";break;
default:alert('�������,�޷�ִ��')
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=saystemp;
this.f2.document.af.oldsays.value=saystemp;
this.f2.document.af.addsign.options[0].selected=true;
this.f2.document.af.tu.options[0].selected=true;
this.f2.document.af.sytemp.focus();
this.f2.document.af.sytemp.value='';return false;}}
else
if (nowmsg=="//"){act="say.asp";}else{act="say.asp"}
this.f2.document.af.action=act;
//this.f2.document.af.aqjhaction.value=act;
{this.f2.document.af.sy.value='';if(saystemp!=''){if((this.f2.document.af.oldsays.value==saystemp)&&(this.f2.document.af.oldtowho.value==this.f2.document.af.towho.value)){alert('���ݲ����ظ���');
this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.select();return false;}this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=saystemp;this.f2.document.af.oldsays.value=saystemp;this.f2.document.af.addsign.options[0].selected=true;
this.f2.document.af.tu.options[0].selected=true;this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.value='';
ty=new Date();var nh=ty.getHours();var nm=ty.getMinutes();var ns=ty.getSeconds();var ct=(nh*3600)+(nm*60)+ns;
if(((ct-lst)<2.5)&&(ct>lst)){this.f2.af.sytemp.value=this.f2.af.oldsays.value;this.f2.af.oldsays.value='';return false;}
else{lst=ct;}this.f2.addOne(this.f2.document.af.sy.value);this.f2.startnosay();
this.f2.document.af.subsay.disabled=1;setTimeout('this.f2.document.af.subsay.disabled=0',3000);return true;}if ((this.f2.document.af.addsign.options[this.f2.document.af.addsign.selectedIndex].value=='0')||(this.f2.document.af.addsign.options[this.f2.document.af.addsign.selectedIndex].value=='')){alert('�����뷢�Ի�ѡ������');
this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.select();return false;}}}
function wdmess(b){this.mess.document.writeln(b);}
function mytitle(title){
this.title.document.open();
this.title.document.writeln("<html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=gb2312\"><\/head>");
this.title.document.writeln("<style type=\"text/css\">body{font-family:\"����\";color:blue;font-size:10.5pt;line-height:15pt;text-align:center}</style>");
this.title.document.writeln("<body bgcolor='E0E0F0' text=#FFFFFF TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0 oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false><div align=center style='font-size: 9pt;)'><font color=green>");
this.title.document.writeln(title+"</font></div></body></html>");
this.title.document.close();}
function DB(){this.mess.document.open();
wdmess("<html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=gb2312\"><\/head>");
wdmess("<style type=\"text/css\">body{background-image: URL(bbg1.jpg);font-family:\"����\";color:white;font-size:9pt;line-height:15pt;text-align:center}a{color:white;text-decoration:none;}a:hover{color:white;text-decoration:underline;}</style>");
wdmess("<body bgcolor=#A90005 TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0 oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false >");
var i=0;i=Math.ceil(Math.random()*(Banner.length-1));
wdmess("<marquee scrollamount=3 scrolldelay=10 onmouseover=this.stop(); onmouseout=this.start();>���:" + Banner[i] + "<\/marquee><\Script Language=JavaScript>setTimeout('parent.DB()',100000);<\/script></body></html>");
this.mess.document.close();}
function tbclutch(){if(this.f2.document.af.tbclutch.value=='ȫ��'){this.f2.document.af.tbclutch.value='��ֱ';this.msgfrm.rows = "*";this.msgfrm.cols = "*";tbclu=false;}else{if(this.f2.document.af.tbclutch.value=='��ֱ'){this.f2.document.af.tbclutch.value='ˮƽ';this.msgfrm.cols = "*,*";this.msgfrm.rows = "*";tbclu=true;}else{this.f2.document.af.tbclutch.value='ȫ��';this.msgfrm.cols = "*";this.msgfrm.rows = "*,*";tbclu=true;}}this.f2.document.af.sytemp.focus();}
self.onerror=null;
var nullframe = '<HTML><BODY BGCOLOR=#000000 text=#ffffff style="font-size:10pt"><center><font color=yellow>��ӭ����<%=Application("aqjh_chatroomname")%></font><br><br>�������ڴӷ�������ȡ����, ���Ժ�^_^</center></BODY></HTML>';
</script>
</head>
<frameset cols="*,160" name=tbgn1 rows="*" framespacing="0" frameborder="no" border="3"  noresize>
<frameset rows="60,0,*,20,0,83,0" cols="*">
<frame src="toptupian.asp" scrolling="NO"  name="gg" marginwidth="3" marginheight="3" noresize>
<frame src="about:blank" name="mess" scrolling="no">
<frameset name=msgfrm rows="*,*" cols="*">
<frame src="javascript:parent.nullframe" name="f0" scrolling="AUTO" framespacing="1" marginheight="3" marginwidth="5" frameborder="yes" noresize>
<frame src="javascript:parent.nullframe" name="f1" scrolling="AUTO" framespacing="1" marginheight="3" marginwidth="5" frameborder="yes">
</frameset>
<frame src="about:blank" name="title" scrolling="no" noresize>
<frame src="about:blank" name="t" marginwidth="5" marginheight="5"  scrolling="NO" noresize>
<frame src="F2.asp" name="f2" scrolling="NO" marginwidth="3" marginheight="8" noresize>
<frame src="about:blank" name="d" scrolling="NO" noresize>
</frameset>
<frameset rows="0,0,22,*,0,0,104" cols="*" name="tbymd">
<frame src="about:blank" name="m">
<frame src="about:blank" name="ps">
<frame src="chang_room.asp" marginwidth="5" marginheight="5" scrolling="no" name="r" noresize>
<frame src="about:blank" marginwidth="5" marginheight="5" scrolling="auto" name="f3" noresize>
<frame src="menu.asp" scrolling="NO"  name="menu" marginwidth="3" marginheight="3" noresize>
<frame src="npc.asp" scrolling="NO"  name="npc" marginwidth="3" marginheight="3" noresize>
<frame src="F4.asp" scrolling="NO" name="CW_MENU" marginwidth="3" marginheight="3" noresize>
</frameset></frameset></html>