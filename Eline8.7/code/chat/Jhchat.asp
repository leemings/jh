<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("sjjh_useronlinename"&i))," "&LCase(sjjh_name)&" ")=0 then ydl=0
	if ydl=1 and clng(nowinroom)<>i then 
		Session.Abandon
		Response.Redirect "../error.asp?id=140"
		Response.End 
	end if
next
'����
chatroomname=trim(Application("sjjh_chatroomname"&session("nowinroom")))
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
sjjh_sid=trim(request.cookies("yxjh")("sjjh_sid")) 
if (sjjh_sid="" or sjjh_sid<>session.sessionid) and Session("sjjh_grade")<6 then
 Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('��һ̨��������˶���ʺţ���ϵͳ�����');}</script>"
 Response.End 
end if
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then 
	Session.Abandon
	Response.Write "<script language=JavaScript>{alert('��ʾ����ֹʹ�ô���');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
Dim SplitReflashPage
Dim DoReflashPage
dim shuaxin_time
DoReflashPage=true
shuaxin_time=10
ReflashTime=Now()
if (not isnull(session("ReflashTime"))) and cint(shuaxin_time)>0 and DoReflashPage then
	if DateDiff("s",session("ReflashTime"),Now())<cint(shuaxin_time) then
   	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��<b><font color=ff0000>"&shuaxin_time&"</font></b>��������ˢ�±�ҳ��<BR>���ڴ�ҳ�棬���Ժ򡭡�"
	response.end
	else
	session("ReflashTime")=Now()
	end if
elseif isnull(session("ReflashTime")) and cint(shuaxin_time)>0 and DoReflashPage then
	Session("ReflashTime")=Now()
end if
'#############�����ҹ������'#############
sjjh_dongtai=0					'�Ƿ���������������ʾ��̬���������chat/gg.js�ļ����ݣ�
'#############�����������'#############
'���û����ϴ���
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='���'",conn,2,2
sjjh_gghigh=clng(rs("b"))
rs.close
rs.open "SELECT * FROM sm where a='����'",conn,2,2
mybanner=rs("c")
jhauto=rs("d")
rs.close
rs.open "select k from r where a='"&chatroomname&"'",conn,3,3
if not(rs.eof and rs.bof) then
	jrht=rs("k")
else
	jrht="��ӭ������"&chatroomname&"��ף���ĵĿ���~~"
end if
rs.close
rs.open "select times,id,��Ա�ȼ�,�ȼ�,���,����,����ͷ��,�Ա�,��������,��ż,ͨ�� from �û� where ����='" & sjjh_name &"'",conn,2,2
tjrf=rs("ͨ��")
newuser=rs("times")-1
if newuser=3 then
newuser=2
end if
sjjh_id=rs("id")
hydj=rs("��Ա�ȼ�")
jhsf=rs("���")
if tjrf=True then
	jhmp="ͨ����"
else
	jhmp=rs("����")
end if
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
hymd=rs("��������")
mywife=rs("��ż")
mmp=rs("����")
dj=rs("�ȼ�")
rs.close
rs.open "select id,a,d,g,h,i,l from vh where b='"&sjjh_name&"' and c=true and j=false",conn
if rs.bof and rs.eof then
	randomize timer
    chatdj=int(rnd*5)+1
if newuser=1 then
	conn.Execute "update �û� set ����=����+50000000,���=���+10,����=����+20,��=��+100,ľ=ľ+100,ˮ=ˮ+100,��=��+100,��=��+100,times=times+1 where ����='"&sjjh_name&"'"
	Response.Write "<script Language=Javascript>alert('���������˷ѡ��\n\n��ӭ����"&sjjh_name&"��\n�������������������ȷ������վ�����͵����˷�\n����5000�򡢽��10������20����ľˮ������100��\n\nֻҪ��ϲ����������������ǹ�ͬ�ļ�԰��\n��������᲻�Ͻ�ʶ��������˷�������\n�����кܶ����ʵ��������Ȥ��������\nף���ڱ�������Զ���ġ���죡\n\n�������Ľ������İ� ���� ����');</script>"
        ii=int(rnd()*20)+1
        sjjh_userinto="<img src=xinrf.gif>##��ȡ��վ���͵����˷�:����<font color=#cc0000><b>5000��</b></font>/���<font color=#cc0000><b>10��</b></font>/����<font color=#cc0000><b>20��</b></font>/��ľˮ������<font color=#cc0000><b>100��</b></font>�������������ˣ���Ҷ�������Ҫ����չ˰���"
        chatdj=7
        sj=30
end if
if sjjh_grade=5 then
         ii=int(rnd()*20)+1
         sjjh_userinto="##<font color=blue>["&mmp&"����]</font>������"&Application("sjjh_chatroomname")&"�������Ǻ�...���ŵĸо����Ǻã���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
         chatdj=7
         sj=30
end if 
if sjjh_grade=10 then
         ii=int(rnd()*20)+1
         sjjh_userinto="##[<font color=red>վ��</font>]�����ˡ�"&Application("sjjh_chatroomname")&"������ϣ����Ҷ�֧�ֱ���������������Ե���̳���ۣ���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
         chatdj=7
         sj=30 
end if

if sjjh_grade>5 and sjjh_grade<10 and sjjh_name<>"��Ȼ" then
         ii=int(rnd()*20)+1
         sjjh_userinto="##[<font color=red>�ٸ�</font>"&jhsf&"]�����ˡ�"&Application("sjjh_chatroomname")&"���������ظ�λ���ҵ�ְ���е��ҵ��𣿡�<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
         chatdj=7
         sj=30 
end if
if sjjh_name="��Ȼ" then
randomize
         ii=int(rnd()*20)+1
         sjjh_userinto="##[<font color=red>վ��</font>]�����ˡ�"&Application("sjjh_chatroomname")&"������ϣ����Ҷ�֧�ֱ���������������Ե���̳���ۣ���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
         chatdj=7
         sj=30 
end if
	Select Case chatdj
case 1
    randomize
	ii=int(rnd()*20)+1
if sex="��" then
	sjjh_userinto="##["&mmp&""&jhsf&"]ƮƮȻ�ģ������ˡ�"&Application("sjjh_chatroomname")&"������С�ܳ����ɾ�����·�������񣡡�<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
else
        sjjh_userinto="##["&mmp&""&jhsf&"]ƮƮȻ�ģ������ˡ�"&Application("sjjh_chatroomname")&"������С�ó����ɾ�����·�������񣡡�<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
end if
case 2
    randomize
	ii=int(rnd()*20)+1
if sex="��" then
    sjjh_userinto="##["&mmp&""&jhsf&"]ֱ����"&Application("sjjh_chatroomname")&"��������˵��:���������ˣ��õط���ô���������ң���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
else
    sjjh_userinto="##["&mmp&""&jhsf&"]ֱ����"&Application("sjjh_chatroomname")&"��������˵��:���������ˣ��õط���ô���������ң���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
end if
case 3
    randomize
	ii=int(rnd()*20)+1
if sex="��" then
    sjjh_userinto="##["&mmp&""&jhsf&"]ƣ���������"&Application("sjjh_chatroomname")&"�������ҵĵ�������Ϊ��Ĵ��ڣ���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
else
    sjjh_userinto="##["&mmp&""&jhsf&"]ƣ���������"&Application("sjjh_chatroomname")&"�������ҵĵ�������Ϊ��Ĵ��ڣ���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
end if
case 4
    randomize
	ii=int(rnd()*2)+1
if sex="��" then
    sjjh_userinto="##["&mmp&""&jhsf&"]�����ˡ�"&Application("sjjh_chatroomname")&"��������Զ֧��--+"&Application("sjjh_chatroomname")&"+--<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
else
    sjjh_userinto="##["&mmp&""&jhsf&"]�����ˡ�"&Application("sjjh_chatroomname")&"��������Զ֧��--+"&Application("sjjh_chatroomname")&"+--<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
end if
case 5
    randomize
	ii=int(rnd()*5)+1
if sex="��" then
    sjjh_userinto="##["&mmp&""&jhsf&"]�Ѿ��ӵ�ɱ���ˡ�"&Application("sjjh_chatroomname")&"��...������λ���Ǹ��ŵ�������˰ɣ���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
else
   sjjh_userinto="##["&mmp&""&jhsf&"]�Ѿ��ӵ�ɱ���ˡ�"&Application("sjjh_chatroomname")&"��...������λ���Ǹ��ŵ�������˰ɣ���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
end if
end select
else
	sjjh_userinto=rs("d")
	vhwj=rs("h")
	vhname=rs("g")
	id=rs("id")
	vhnj=rs("i")
	sj=DateDiff("d",rs("l"),now())
	randomize timer
	ii=int(rnd()*10)

        if sjjh_grade=10 then
	sjjh_userinto="##[<font color=red>վ��</font>]�ڸ�·���ɵĴ�ӵ�£�������ʹ⻷������"&Application("sjjh_chatroomname")&"�������һ������⣡<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font><br><font color=red>��������Ϣ��</font>վ��["&sjjh_name&"]����"&Application("sjjh_chatroomname")&"�ˣ����·Ӣ�ۻ�ӭ��������["&sjjh_name&"]��������1000000�㣡��������1000000��"
	conn.execute "update �û� set ����=����+1000000,����=����+1000000 where ����='"&sjjh_name&"'"
	end if
        if sjjh_grade=5 then
	sjjh_userinto="##<font color=blue>["&mmp&"����]</font>�ڵ��ӵ���ͬ�£��������������ˡ�"&Application("sjjh_chatroomname")&"�������һ����С���ǣ������ˣ���<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font><br><font color=red>��������Ϣ��</font>["&mmp&"]����["&sjjh_name&"]����"&Application("sjjh_chatroomname")&"�ˣ�����е��ӻ�ӭ������["&sjjh_name&"]��������10000�㣡��������10000��"
	conn.execute "update �û� set ����=����+10000,����=����+10000 where ����='"&sjjh_name&"'"
	end if
        if sjjh_grade>5 and sjjh_grade<10 then
	sjjh_userinto="##[<font color=red>�ٸ�</font>"&jhsf&"]�ڵ��������У�ɱ�����ڵ������ˡ�"&Application("sjjh_chatroomname")&"����ŭ�������������Ҳ����<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font><br><font color=red>��������Ϣ��</font>[�ٸ�"&jhsf&"]["&sjjh_name&"]����"&Application("sjjh_chatroomname")&"�ˣ�����С�Ĵ��¡�������["&sjjh_name&"]��������50000�㣡��������50000��"
	conn.execute "update �û� set ����=����+50000,����=����+50000 where ����='"&sjjh_name&"'"
	end if
     if sjjh_grade>4 then
	sj=30 
	vhnj=50
     else
	if sj>42 or vhnj<1 then
		conn.execute "update vh set j=true where id="&id
		if sex="��" then
	           sjjh_userinto="##ƮƮȻ�ģ������ˡ�"&Application("sjjh_chatroomname")&"������С�ܳ����ɾ�����·�������񣡡�"
		else
                   sjjh_userinto="##ƮƮȻ�ģ������ˡ�"&Application("sjjh_chatroomname")&"������С�ó����ɾ�����·�������񣡡�"
		end if
		
	else
		if sj>35 or vhnj<10 then
			if ii<3 then
				conn.execute "update vh set j=true where id="&id
				
				sjjh_userinto="##["&mmp&""&jhsf&"]�������Լҵ�"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">Ҫ����"&Application("sjjh_chatroomname")&"���������İ������ݳ���Щ����--���϶���ʹ�ù��Ȼ����ʧ�ޣ����ˣ�<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font><br><font color=red>��������Ϣ��</font>["&sjjh_name&"]���������ˣ�ֻ���ܲ�����"&Application("sjjh_chatroomname")&"����������졣������"
			else
				if Isnull(vhname) or vhname="" or vhname="��" then vhname=rs("a")
				sjjh_userinto=replace(sjjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				sjjh_userinto=sjjh_userinto&"<br><font color=red>��������Ϣ��</font>��СϺ�����ؿ���["&sjjh_name&"]����¶��һ����ɷĽɷ�ı��飬��ʱ�Լ�Ҳ���С�����["&sjjh_name&"]��������28�㣡��������1000��"
				conn.execute "update �û� set ����=����+28,����=����+1000 where ����='"&sjjh_name&"'"
				conn.execute "update vh set i=i-1 where id="&id
			end if
		else
			if ii<2 then
				sjjh_userinto="##["&mmp&""&jhsf&"]�������Լҵ�"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">Ҫ����"&Application("sjjh_chatroomname")&"�����������Ͻ�ͨ������ֻ��һ·�ܲ�ǰ��������"&Application("sjjh_chatroomname")&"ʱ��������졣<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font><br><font color=red>��������Ϣ��</font>��СϺ��������������["&sjjh_name&"]����¶��һ������ı��顣����["&sjjh_name&"]��������18�㣡"
				conn.execute "update �û� set ����=����+18 where ����='"&sjjh_name&"'"
			else
				if Isnull(vhname) or vhname="" or vhname="��" then vhname=rs("a")
				sjjh_userinto=replace(sjjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				sjjh_userinto=sjjh_userinto&"<br><font color=red>��������Ϣ��</font>��СϺ�����ؿ���["&sjjh_name&"]����¶��һ����ɷĽɷ�ı��飬��ʱ�Լ�Ҳ���С�����["&sjjh_name&"]��������28�㣡��������1000��"
				conn.execute "update �û� set ����=����+28,����=����+1000 where ����='"&sjjh_name&"'"
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
if Application("sjjh_closedoor")="1" then Response.Redirect "../error.asp?id=100"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("sjjh_disproxy")="1" and (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
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
'Session("sjjh_lastsaytime")=sj

if Session("sjjh_inthechat")<>"1" then
if InStr(LCase(Application("sjjh_useronlinename"&nowinroom))," " & LCase(sjjh_name) & " ")<>0 then Response.Redirect "../error.asp?id=300"
Session("sjjh_inthechat")="1"
Session("sjjh_savetime")=now()
Session("sjjh_lasttime")=sj
myzanli=0
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0 then myzanli=1
'����0,�Ա�1,����2,���3,ͷ��4,�ȼ�5,id6,��Ա�ȼ�7,����8,����9
myonline = sjjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & sjjh_jhdj& "|" & sjjh_id& "|" & hydj&"|"&myzanli&"|"&"����"
Application.Lock
dim newonlinelist()
js=1
onlinelist=Application("sjjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")

'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'���������ֺ���ƴ������
'if yjl=0 and len(onuser(2))<len(jhmp) then			'���������ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(2))>len(jhmp) then			'���������ֳ���,�̵���ǰ
if yjl=0 and StrComp(onuser(0),sjjh_name,1)=1 then	'�����ֺ���ƴ������
'if yjl=0 and len(onuser(0))<len(sjjh_name) then	'�����ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(0))>len(sjjh_name) then	'�����ֳ���,�̵���ǰ
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

Application("sjjh_onlinelist"&nowinroom)=newonlinelist
Application("sjjh_useronlinename"&nowinroom)=Application("sjjh_useronlinename"&nowinroom)& " "&sjjh_name & " "
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
'sjjh_userinto="�ڽ����Ĵ����,##����һ���µϣ����������������һ��ȫ�Զ���������"&Application("sjjh_chatroomname")&",���ǡ���������ϴ󣡡�������Ի:����λ�ֵܣ�С������!��<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
if newuser=1 then sjjh_userinto=sjjh_userinto&"<br>[<font color=ff0000>"&sjjh_name&"</font>]��һ���������ǽ������Ҷ���չˡ���"
if tjrf=True then sjjh_userinto="����:##����������ɱ�����飬���°ܻ����ֱ��ٸ�ͨ����ɱ��ͨ���˷����ٸ�������100��Ľ�������λӢ���ɲ��ܷŹ�����������ֵĺû��ᡭ��<font color=#000000>����</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>�ε�½��</font>"
act="����"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
says="<font color=#cc0000>�����ߡ�</font><font color=008800>" & Replace(sjjh_userinto,"##","<img src=../ico/"& jhtx &"-2.gif><a href=javascript:parent.sw(\'[" & sjjh_name & "]\'); target=f2>" & sjjh_name &"</a><font color=#000000>{</font>ID:<font color=red><b>"& sjjh_id &"</b></font>�ȼ�:<font color=#ff00ff><b>"& dj &"</b></font><font color=#000000>}</font>") & "</font><font class=t>(" & t & ")</font><bgsound src=readonly/okok.wav loop=1>"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
'Response.Write saystr
'Response.End
end if
online=Application("sjjh_onlinelist"&nowinroom)
onlineno=ubound(online)
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
for i=0 to chatroomnum
	uonline=split(trim(Application("sjjh_useronlinename"&i)),"  ")
	onlinenum=ubound(uonline)+1
	sj_chat_info=split(sjjh_roominfo(i),"|")
roomtemp=roomtemp&"<img src='img/room.gif'><a href=\"&chr(34)&"changeroom.asp?roomsn="&i&"&chatroomname="&sj_chat_info(0)&"\"&chr(34)&" target=\"&chr(34)&"d\"&chr(34)&" title=\"&chr(34)&"���ƣ�"&sj_chat_info(3)&"\"&chr(34)&">"&sj_chat_info(0)&"</a>"
next
hiddenadmin="|"&Application("hidden_admin")&"|"
hiddenadmin=Replace(hiddenadmin,",","|")
randomize()
pic=int(rnd*3)
pcinfo = Request.ServerVariables("HTTP_USER_AGENT")
If InStr(pcinfo, "NT 5") <> 0 Then
 gdsd = 4
Else
 gdsd = 7
End If
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html;charset=gb2312">
<style>BODY{background-image: URL(bj.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}\n");</style>
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
if(window.name!="sjjh"){ var i=1;while (i<=50){window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ɣ�������50�Σ�");i=i+1;}top.location.href="../exit.asp"}
var crm="<%=chatroomname%>",bgc="<%=Application("sjjh_chatcolor")%>",systitle="<%=Application("sjjh_tltie")%>";
var figo='<%=Application("figo")%>';
var myn="<%=sjjh_name%>",mywife="<%=mywife%>",chatbgcolor="<%=chatbgcolor%>",chatimage="<%=chatimage%>";
var hiddenadmin="<%=hiddenadmin%>",cs=<%=sjjh_grade%>,automan="<%=Application("sjjh_automanname")%>";
var myroom=<%=session("nowinroom")%>,slbox=0,lst=0,tbclu=true,mdcls=true,listfaces=false;
var showhao=false,myxq=1,mymp=1,jhtx="<%=jhtx%>",showpy="<%=hymd%>",showmp="<%=jhmp%>",headhigh=24;
var showsex="<%=sex%>",showseek="",bc = 1,clsok=0,hang = 0 ;showtype = 0,bgimg=""
var gdsd = <%=gdsd%>;
<%if sjjh_dongtai=1 then%>jsjsstr="<\script src=gg.js></\script>";<%else%>jsjsstr="";<%end if%>
var badword = new Array("������վ","�ƽ���","���Ľ���","ȥ��Ľ���","ȥ�ҵĽ���","���ҵĽ���","���ҵĽ���","ȥ��������","ȥ��������","����������","����������","�ý���","�񽭺�","�Ľ���","�侫","��","�Ҹ�","����","���������","���������","���������","������","ma de","kan ni","fa","quan","ni ma","ni mu","niba","ni ba","nima","nimu","ni nai nai","ni jie","ni mei","nai","ye","ma","ba","zu","zhu","ɵ��","����","�ҿ�","����","������","������","ͱ��","ͱ����","��Ů","��Ѽ","��Ѽ","sai","����Ա","����","kao","cao","����ĸ","��ĸ","��ľ","ȥ��","��ʺ","��","��","��","��","����","��","үү","����","��ү","����","��ү","����","����","�Ҳ�","������","����","��","ɵ","��","������","���","����","���","����","�","����","����","��","����","����","����","����","����","����","غ��","��Ƥ","��ͷ","��","�P","��","�H","����","��","��","����","����","����","����","����","��Ů","����","����","��","�Ҷ�","���","���","��","����","����","�ڹ�","ʺ","ƨ","���","���˹�","http","www","com","cn","org","net","://","WWW","HTTP","Http","HTTp","HttP","hTTP","htTp","httP","COM","CN","update","UPDATE","grade","allvalue","set","SET","Set","sEt","seT","add","Add","AdD","ADd","ADD","table","����=","ORG","NET","Www","wWw","WWw","wwW","wWW","Net","NEt","nET","neT","Com","COm","cOM","Org","ORg","oRG","orG","coM","fuck","cao","gan","sai","��","����","����qq","����QQ","���ң��","���ңѣ�","����qq","����QQ","���ң��","���ңѣ�","��QQ","��qq","�ӣ��","�ӣѣ�","Q Q","q q","�� ��","�� ��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","123456789","000000000000","1111111111","2222222222","3333333333","4444444444","5555555555","6666666666","7777777777","8888888888","9999999999","aaaaaaaaaa","bbbbbbbbbb","cccccccccc","dddddddddd","eeeeeeeeee","ffffffffff","gggggggggg","hhhhhhhhhh","iiiiiiiiii","jjjjjjjjjj","kkkkkkkkkk","llllllllll","mmmmmmmmmm","nnnnnnnnnn","oooooooooo","pppppppppp","qqqqqqqqqq","rrrrrrrrrr","ssssssssss","tttttttttt","uuuuuuuuuu","vvvvvvvvvv","wwwwwwwwww","xxxxxxxxxx","yyyyyyyyyy","zzzzzzzzzz","<s","<S","��������","����������","ŶŶŶŶ","��������","��������","��������","��������");
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/����������������������������������������?������������������" ;
var writejs=false;
var Maxwrite=100;
var askjs="��ϵͳ���Ѵ��ڽű����ã�����ָ���Զ��������ʹ������ָ��";
var askqp2="<br><font size=2><font color=red>����ʾ��</font>�Ի������أ�10���Ӻ��Զ�������</font><br>";
var writeNUM=0;
var askqp="<br><font size=2><font color=red>����ʾ��</font>����<a href='javascript:parent.qp()'>[��]</a>����Լ������Դ����ʾ���κ��Զ�������</font><br>";
document.write("<title>��ӭ������<%=Application("sjjh_chatroomname")%>��,ף���ĵĿ��ģ���վ���ù������� - 51eline.com</title>");
function write(cls){var fsize,lheight;
//if(cls==1){fsize=this.f2.document.af.fs.value;lheight=this.f2.document.af.lh.value;}else
{fsize='10.5';lheight='125';}
this.f1.document.open();
this.f1.document.writeln("<html><head><title>�Ի���</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f1.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"����\";font-size:" + fsize + "pt;CURSOR: url('15.ani');scrollbar-face-color:#effaff;scrollbar-shadow-color:#eeeeee;scrollbar-highlight-color:#ffffff;scrollbar-3dlight-color:#eeeeee;scrollbar-darkshadow-color:#ffffff;scrollbar-track-color:#ffffff;scrollbar-arrow-color:#dddddd;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL(bj.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f1.document.writeln("<\Script Language=\"JavaScript1.1\">var autoScrollOn=1;var scrollOnFunction;var scrollOffFunction;");
this.f1.document.writeln("function scrollit(){if(!parent.f2.document.af.as.checked){autoScrollOn=0;return true;}else{autoScrollOn=1;StartUp();return true;}}");
this.f1.document.writeln("function scrollWindow(){if(autoScrollOn==1){this.scroll(0,65000);parent.f0.scroll(0,65000);setTimeout('scrollWindow()',200);}}");
this.f1.document.writeln("function scrollOn(){autoScrollOn=1;scrollWindow();}");
this.f1.document.writeln("function scrollOff(){autoScrollOn=0;}");
this.f1.document.writeln("function StartUp(){this.onblur=scrollOnFunction;this.onfocus=scrollOffFunction;scrollWindow();}");
this.f1.document.writeln("scrollOnFunction=new Function('scrollOn()');");
this.f1.document.writeln("scrollOffFunction=new Function('scrollOff()');");
this.f1.document.writeln("StartUp();</\script>");
this.f1.document.writeln("<body background='"+bgimg+"' oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f1.document.writeln("<span class=l><font color=red>�������ˢ�¡�</font>���һ�ӭ<font color=red>��"+myn+"��</font>������"+crm+"��- 51eline.com<font class=t>(<%=t%>)</font></span><br>");
this.f1.document.writeln("<%=roomtemp%><br>");
this.f0.document.open();
this.f0.document.writeln("<html><head><title>������ʾ</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f0.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"����\";font-size:" + fsize + "pt;CURSOR: url('aixin.ani');scrollbar-face-color:#effaff;scrollbar-shadow-color:#eeeeee;scrollbar-highlight-color:#ffffff;scrollbar-3dlight-color:#eeeeee;scrollbar-darkshadow-color:#ffffff;scrollbar-track-color:#ffffff;scrollbar-arrow-color:#dddddd;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL(jjchatbg.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f0.document.writeln("<body background='"+bgimg+"'oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f0.document.writeln("<span class=l><font color=red>�������ˢ�¡�</font>���һ�ӭ<font color=red>��"+myn+"��</font>������"+crm+"��- 51eline.com<font class=t>(<%=t%>)</font></span><br><b>���������桿</b><%=jrht%><br>"+jsjsstr);
this.t.location.href="t.asp";parent.DB();parent.mytitle(systitle);}
//sh:s0:��ɫ s1:ɫ�� s2:���� s3:˵���� s4:���� s5:�ܻ� s6:���� s7:˽�� s8:����
function sh(s0,s1,s2,s3,s4,s5,s6,s7,s8){
var show = "",ss="";if(s2=="����"){parent.mytitle(s6);return;}
if (myroom != s8){return;}//�Է��䴦��
if (s3!=myn && s5!=myn && s7 ==1 && slbox==0){return;}//��˽�Ĵ���
hang=hang+1;
if(hang>800 && clsok==0){if(confirm("�����Ļ��ʾ�ķ��������Ѿ�������800�����ǵ������Ե��������� ���Ƿ������������ȷ���������������ȡ�����Ժ��ٳ��ִ���ʾ��")){hang=0;parent.qp();}else{clsok=1;}}
if (s2=="����" && s5==myn && this.f2.document.af.dwtx.checked == true){if(confirm("��Ϣ��["+s3+"]�������������������Ƿ�����������")){this.d.location.href="dantiao.asp?name="+s3+"&yn=1";}else{this.d.location.href="dantiao.asp?name="+s3+"&yn=0";}}//��������
if(s2=="����"){
if (hiddenadmin.indexOf("|"+s3+"|")  !=-1){return;}
parent.m.location.reload();
if(s3==mywife && this.f2.document.af.py.checked==true){alert("�����ż["+mywife+"]�����ˡ�����");}
else{if(showpy.indexOf(s3+"|")!=-1 && this.f2.document.af.py.checked==true && s3!=""){alert("��ĺ���["+s3+"]�����ˡ�����");}}
}
if (s2=="��ը����" && s5==myn){top.location.href="readonly/bomb1.htm";}
if (s2=="ը��" && s5==myn){top.location.href="readonly/bomb.htm";}
if (s2=="��ըӲ��" && s5==myn){top.location.href="readonly/bomb2.htm";}
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
if(s2=="����"&&slbox!=0){show=show+"[�Է�Ϊ��"+s5+"]"}
if(s2!="����"){show=show+s6+"</font><br>"}//�Զ������ж�
else{show = show + "<a href=javascript:parent.sw('["+s3+"]'); target=f2><font color=" + s0 + ">"+s3+"</font></a>"+ s4 + "<a href=javascript:parent.sw('["+s5+"]'); target=f2><font color=" + s0 + ">"+s5+"</font></a>" + "˵��" + s6 + "</font><br>";
if ((s3==myn || s5==myn) && tbclu==false){show="<div  style='background:#66CCCC;'>"+show+"</div>"}//����ȫ����ɫ��
}
if (s5==automan){//���������
dz = "��˽�ġ�"+automan+"��"+s3+"˵��"+huida[Math.floor(Math.random()*huida.length)]+"<br>";
for(i=0;i<wen.length;i++){if(s6.search(wen[i]) != -1){dz="��˽�ġ�"+automan+"��"+s3+"˵��"+huida[i]+"<br>"}}
show=show+dz;}
if(tbclu && (myn == s3 || myn == s5)){this.f1.document.writeln(show);}
else{this.f0.document.writeln(show);}
writeNUM=writeNUM+1;
if(writeNUM>Maxwrite){writqp();writeNUM=0}
}

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

function qp()
{
writejs=false;
Maxwrite=100;
this.t.location.href="about:blank";
this.f0.location.href="about:blank";
this.f1.location.href="about:blank";
setTimeout('parent.write(1)',500);}
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
function wmd(b){this.f3.document.writeln(b);}
function shake(n) {if (window.top.moveBy) {for (i = 10; i > 0; i--) {for (j = n; j > 0; j--) {window.top.moveBy(0,i);window.top.moveBy(i,0);window.top.moveBy(0,-i);window.top.moveBy(-i,0);}}}}
function md1(ren){
	if (this.f2.document.af.mdsx.checked == false){return;}
	this.f3.document.open();
	wmd("<html><head><meta http-equiv='content-type' content='text/html; charset=gb2312'><title>�����û��б�</title><style type='text/css'>");
	wmd("body{CURSOR: url('aixin.ani');scrollbar-face-color:\"#effaff\";scrollbar-shadow-color:\"#eeeeee\";scrollbar-highlight-color:\"#ffffff\";scrollbar-3dlight-color:\"#eeeeee\";scrollbar-darkshadow-color:\"#ffffff\";scrollbar-track-color:\"#ffffff\";scrollbar-arrow-color:\"#dddddd\";font-family:\"����\";font-size:12pt;}td{font-family:\"����\";font-size:10.5pt;line-height:125%;}A{color:#ffffff;text-decoration:none;}A:Hover{color: #FF0000; font-family: \"����\"; position: relative; left: 2px; top: 1px; clip:  rect(   )}A:Active {color:#ffffff}.b{color:#ffff99;}.g{color:#00FF00;}.hb{color:#b7d4f1;}.hg{color:#00FFFF;}.d{font-family:\"����\";font-size:10pt;color:#9999cc;}.z{font-family:\"����\";font-size:10pt;color:orange;}.xinxi{font-family:\"����\";font-size:9pt;}.banq{font-family:\"����\";font-size:10pt;color:ffff00; filter: DropShadow(Color=000000, OffX=1, OffY=1, Positive=1)}.gf{font-family:\"����\";font-size:10.5pt;color:ff6600;}.gfm{font-size:10.5pt;color:ff0000;}.zl{color:999999;text-decoration: line-through;}</style>");
	wmd("<div align=\"center\"><font color=\"#b3d4ff\"<b>��"+crm+"��</b></font><hr size=1 color=b3d4ff>");
	wmd("<font color=\"#ffff00\" style=\"font-size:10.5pt\";font-family:\"����\">˫�������Ҽ�������</font><br><font color=\"#ffff00\" style=\"font-size:12pt\">--+</font> <font color=red><b>"+ren+"</b></font><font color=\"#00FF00\" style=\"font-size:11pt\">������</font> <font color=\"#ffff00\" style=\"font-size:12pt\">+--</font><br></div>");
	wmd("<div id='Tips' style='position:absolute; left:0; top:0; height: 226px;width:140; display=none;'><IFRAME frameBorder=no height=226px marginHeight=0 marginWidth=0 name=show width=140px scrolling=NO noresize></IFRAME></div>");
	wmd("<div id='myTips' style='position:absolute; left:0; top:0; height: 226px;width:130; display=none;'></div>");
	wmd("<\script language=\"JavaScript\">var NS4=(document.layers);var IE4=(document.all);var win=window;var n=0;function findInPage(str){var txt,i,found;if(str==\"\"){return false;}if(NS4){if(!win.find(str))while(win.find(str,false,true))n++;else{n++;}if(n==0)alert(\"��Ҫ������û���ҵ���\");}if(IE4){txt=win.document.body.createTextRange();for(i=0;i<=n && (found=txt.findText(str))!=false;i++){txt.moveStart(\"character\",1);txt.moveEnd(\"textedit\");}if(found){txt.moveStart(\"character\",-1);txt.findText(str);txt.select();txt.scrollIntoView();n++;}else{if(n>0){n=0;findInPage(str);}else{alert(\"��Ҫ������û���ҵ���\");}}}return false;}");
	wmd("function s(name){");
	wmd("parent.sw(name);");
	wmd("}");
	wmd("<\/script>");
	
	wmd("</head><body oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=\"006699\" background=\"f2.gif\" bgproperties=\"fixed\">");
	wmd("<\script language=\"JavaScript\">var currentpos,timer;function initialize(){timer=setInterval(\"scrollwindow()\",1);}function sc(){clearInterval(timer);}function scrollwindow(){currentpos=document.body.scrollTop; window.scroll(0,++currentpos);if (currentpos != document.body.scrollTop) sc();}document.onmousedown=sc;document.ondblclick=initialize;function New(para_URL){var URL =new String(para_URL);window.open(URL,'','resizable,scrollbars')}");
	wmd("function ShowTips(uurl,pThis){Hiddenmy();window.open(uurl,'show');var pTip = document.all['Tips'].style ;pTip.left = 1 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Showmy(data,pThis){lis = data.split('|');s='<table width=130 border=0 cellpadding=1 cellspacing=0 bgcolor=';if(lis[8]==1){s+='#dedfdf';}else{s+='#ffffe7';}");
	wmd("s+='><tr><td width=44% align=center><img src=../ico/'+lis[4]+'-2.gif width=32 height=32></td>';");
	wmd("s+='<td width=56% align=center valign=top><font color=blue>'+lis[0]+'</font><br>';");
	wmd("s+='<div align=left><font class=xinxi>�Ա�:';if(lis[1]=='��'){s+='����';}else{s+='Ů��';}s+='<br>�ȼ�:'+lis[5]+'</font></div></td></tr>';");
	wmd("s+='<tr><td colspan=2 align=center><font class=xinxi>����:'+lis[2]+'&nbsp;'+lis[3];if(lis[2]!='NPC'){s+='<br>עID:'+lis[6]+'&nbsp;&nbsp;';if(lis[7]==0){s+='�ǻ�Ա';}else{s+=lis[7]+'����Ա';}}s+='</font></td></tr></table>';");
	wmd("myTips.innerHTML = s;var pTip = document.all['myTips'].style ;pTip.left = 6 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Hiddenmy(){var obj = document.all['myTips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function Hidden(){this.show.location.href='about:blank';var obj = document.all['Tips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function getPos(obj,type){var n = 0 ;while(obj!=null){if(type=='top'){n += obj.offsetTop;}else{n += obj.offsetLeft;}obj = obj.offsetParent ;}return n;}");
	wmd("<\/script>");

	wmd("<table width=100% Height=30 border=0 cellspacing=0 cellpadding=0 align='left'>");
	wmd("<form name='search' onSubmit='return findInPage(this.string.value);'><tr><td><br><div align=\"center\"><input name='string' type='text' size=8 onChange='n=0;' style='font-family:����;font-size:9pt;background-color:008800;color:FFFFFF;border: 1 double'> <input type='submit' value='����' style='font-size:9pt;background-color:FF9900;color:FFFFFF;width:30;height:16px;border: 1 double'></form></div></td></tr><tr><td class=banq>");
	wmd("<a href=javascript:parent.sw('[���]');>���</a><br>");

	}
function md2(data){
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
    	
    	
    	//'0����,1�Ա�,2����,3���,4ͷ��,5�ȼ�,6id,7״̬,8���뿪,9����
	if (showtype==1){showstr=showpy;showseek=lis[0];}else{if (showtype==2){showstr=showmp;showseek=lis[2];}else{if (showtype==3){showstr=showsex;showseek=lis[1];}else{showstr="";showseek="";}}}
	if (showpy.search(lis[0]) != -1){friend="<img src='../jhimg/friend.gif' width='12' height='12'>";}
	else{friend=""}
	if (lis[0]==myn){myself="<img src='../jhimg/self.gif' width='12' height='12'>";}
	else{myself="";}
	if (listfaces==true){
	faces="<img src='../ico/"+lis[4]+"-2.gif' width='"+headhigh+"' height='"+headhigh+"'>"+friend+myself;}
	else{faces="";}
	if(lis[2]=="�ٸ�")gf="<font class=gfm>�Y</font>";else gf="";
	ss=faces+"<a href=\"JavaScript:parent.sw(\'[" + lis[0] + "]\');\"";
	ss=ss+" onmouseover=\"Showmy('" + data + "'," + "this" + ");\" onmouseout=\"Hidden();Hiddenmy();\""
	if(lis[2]!="NPC"){ss=ss+"oncontextmenu=\"ShowTips('../show/userface.asp?username=" + lis[0] + "&sex="+lis[1]+"',this);\""}
	else{ss=ss+"oncontextmenu=\"JavaScript:alert('NPC��֧��Show����!');\""}
	ss=ss+"><font class=\"" + cls + "\">"+lis[0]+"</font></a>"+gf
	//�Ƿ���ʾ����
	if (mymp==1){ss=ss+"&nbsp;<font class=\"" + mp + "\">" +lis[2]+ "</font>";}
	//�Ƿ���ʾ����
	if (myxq==1){ss=ss+"<font class=xq>&nbsp;"+xq+"</font>";}
	ss=ss+"<br>";
	if (hiddenadmin.indexOf("|"+lis[0]+"|")  ==-1){
	if (showtype != 0){if (showstr.search(showseek) != -1 || lis[0]==myn){wmd(ss);}}else{wmd(ss);}}}
function md3(){
	if (this.f2.document.af.mdsx.checked == false){return;}
		wmd("<br><div align=\"center\"><a href=http://51eline.com target=_blank><img src=../logo.gif width=88 height=31 border=0 alt=ȫ�����쾫�ʽ�������̳></a><HR size=1 color=b3d4ff><font class=banq>��Ȩ:���߽�����վ<br>�汾:ELINE 8.7.0<br>����//����:һ����<br>");
	if (listfaces==true){wmd("Ϊ�Լ�.<img src='../jhimg/self.gif' width='12' height='12'><img src='../jhimg/friend.gif' width='12' height='12'>.Ϊ����</font></div>");}
	wmd("</td></tr></table></body></html>")
	this.f3.document.close();
}
function fc(){rn();setTimeout('parent.write()',0);}
function rn(){this.f2.document.af.username.value=myn;}
function clsay(){if(cs<6){setTimeout("this.f2.document.af.sytemp.value=''",0);}}
function IsBadWord(m){var tmp = "" ;for(var i=0;i<m.length; i++){for(var j=0;j<badstr.length;j++)if(m.charAt(i) == badstr.charAt(j)) break;if(j==badstr.length) tmp += m.charAt(i) ;}for(i=0;i<badword.length;i++) if(tmp.search(badword[i]) != -1) return true;return false;}
function Warning(){this.f2.document.af.sytemp.value='';if(bc > 12) d.location.href="autokick.asp";else{bc++;alert("�벻Ҫ����������ʹ�ý���,�����߳���");}}
function checksays()
{if(this.f2.document.af.addvalues.checked)
{alert("��Ŀǰ�����Զ��ݵ㹦�ܣ��޷����ԣ�")
return false;}
if(IsBadWord(this.f2.document.af.sytemp.value)){Warning();return false;}
var maxlingual=20;
var pos=0;
var lingualnum=0;
var i;
var lingualarr=new Array(maxlingual);
var msg=this.f2.document.af.sytemp.value;
var nowmsg=msg.substr(0,2);
var act;
var sjcz = new Array("/͵Ǯ","/����","/����","/���","/�Ͷ���","/�ͽ��","/����","/��Ѩ","/��Ѩ","/����","/����","/����","/���","/ն��","/�¶�","/���Ǵ�","/͵Ǯ","/Ͷ��","/����","/��ԯ","/����","/��Ǯ","/���","/����","/����","/ת��","/��ͽ","/ը��","/ԭ�ӵ�","/����","/�Զ�","/�ͻ�","/�����̷�","/�Ķ�","/˫�˶Ĳ�","/���ɷ���","/���ip","/����ip","/����","/�������","/�Զ�����","/���ɴ�ս","/����","/���崺��","/Ǭ��һ��","/��ƭ��Ů","/��ƭ����","/ͬ���ھ�") ; 
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
alert("���ܶԴ�ҡ������˻����Լ�ʹ�ñ����ܣ�"+sjcz[i]+aftowho+automan);
this.f2.document.af.sytemp.value="";
return false;
}}
if(cmd=="/"& cmd1!="//")
{var spacepoint=msg.indexOf("$");
var spacepoint=(spacepoint==-1?msg.length:spacepoint);
var cmd=msg.substring(1,spacepoint);
switch(cmd){
case "����":act="sjfunc/1.asp";break;case "���һ��":act="sjfunc/900.asp";break;
case "��Ѩ":act="sjfunc/3.asp";break;case "����":act="sjfunc/4.asp";break; 
case "����":act="sjfunc/5.asp";break;case "���":act="sjfunc/6.asp";break; 
case "�ͷ�":act="sjfunc/7.asp";break;case "ն��":act="sjfunc/8.asp";break; 
case "����":act="sjfunc/9.asp";break;case "�¶�":act="sjfunc/10.asp";break; 
case "͵Ǯ":act="sjfunc/11.asp";break;case "���Ǵ�":act="sjfunc/12.asp";break;  
case "Ͷ��":act="sjfunc/13.asp";break;case "����":act="sjfunc/14.asp";break; 
case "����":act="sjfunc/15.asp";break;case "��Ǯ":act="sjfunc/16.asp";break; 
case "����":act="sjfunc/17.asp";break;case "״̬":act="sjfunc/18.asp";break; 
case "����":act="sjfunc/19.asp";break;case "���":act="sjfunc/20.asp";break; 
case "����":act="sjfunc/21.asp";break;case "����":act="sjfunc/22.asp";break; 
case "�Ķ�":act="sjfunc/23.asp";break;case "�Ŵ�":act="sjfunc/24.asp";break; 
case "��ʦ":act="sjfunc/25.asp";break;case "��ͽ":act="sjfunc/26.asp";break; 
case "����":act="sjfunc/27.asp";break;case "��Ŀ":act="sjfunc/28.asp";break; 
case "����":act="sjfunc/29.asp";break;case "����":act="sjfunc/30.asp";break; 
case "��Ǯ":act="sjfunc/31.asp";break;case "ȡǮ":act="sjfunc/32.asp";break; 
case "ת��":act="sjfunc/33.asp";break;case "���":act="sjfunc/34.asp";break; 
case "����":act="sjfunc/35.asp";break;case "����":act="sjfunc/36.asp";break; 
case "��Ѩ":act="sjfunc/37.asp";break;case "����":act="sjfunc/38.asp";break; 
case "ŭ��":act="sjfunc/39.asp";break;case "����":act="sjfunc/40.asp";break; 
case "��ip":act="sjfunc/41.asp";break;case "����":act="sjfunc/42.asp";break; 
case "����":act="sjfunc/43.asp";break;case "�ÿ�":act="sjfunc/44.asp";break; 
case "���յ���":act="sjfunc/46.asp";break; case "��ԯ":act="sjfunc/14xx.asp";break;
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
case "��������":act="sjfunc/77.asp";break;case "����":act="sjfunc/78.asp";break; 
case "���崺��":act="sjfunc/78.asp";break;case "�Ͷ���":act="sjfunc/79.asp";break; 
case "�ͽ��":act="sjfunc/80.asp";break;case "Ǭ��һ��":act="sjfunc/81.asp";break; 
case "��������":act="sjfunc/82.asp";break;case "�뿪����":act="sjfunc/82.asp";break; 
case "��ƭ��Ů":act="sjfunc/83.asp";break;case "��ƭ����":act="sjfunc/84.asp";break; 
case "�����书":act="sjfunc/87.asp";break;case "��������":act="sjfunc/88.asp";break;
case "�����Ա�":act="sjfunc/90.asp";break;case "Ѱ�ҽ��":act="sjfunc/91.asp";break;
case "ͬ���ھ�":act="sjfunc/92.asp";break;case "͵���":act="sjfunc/93.asp";break;
case "������":act="sjfunc/94.asp";break;case "Ѱ������":act="sjfunc/95.asp";break;
case "������":act="sjfunc/96.asp";case "��������":act="sjfunc/45001.asp";break;
case "���ͷ���":act="sjfunc/45002.asp";break;case "�淨��":act="sjfunc/450020.asp";break;
case "ȡ����":act="sjfunc/450021.asp";break;case "�β�":act="sjfunc/501.asp";break;
case "����":act="sjfunc/502.asp";break;case "��ǩ":act="sjfunc/503.asp";break;
case "����":act="sjfunc/601.asp";break;case "�ٱ���":act="sjfunc/45007.asp";break;
case "��ʩ��":act="sjfunc/45008.asp";break;case "ħ������":act="sjfunc/45009.asp";break;
case "�Ȼ��˼�":act="sjfunc/45011.asp";break;case "����׶":act="sjfunc/45013.asp";break;
case "�Ի��":act="sjfunc/450010.asp";break;case "����":act="sjfunc/85.asp";break;
case "�ߵ�":act="sjfunc/86.asp";break;case "�ƶ�":act="npc/a1.asp";break;
case "Ѱˮ����":act="sjfunc/45003.asp";break;case "ħ��ˮ��":act="sjfunc/45004.asp";break;
case "Ѱ�ҷ���":act="sjfunc/45005.asp";break;case "ִ��":act="sjfunc/602.asp";break;
case "û�շ���":act="sjfunc/450151.asp";break;case "��������":act="sjfunc/45015.asp";break;
case "������":act="sjfunc/45012.asp";break;case "������":act="sjfunc/45013.asp";break;
case "ʹ��":act="sjfunc/45019.asp";break;case "�����Ṧ":act="npc/1.asp";break;
case "�Ṧ�ݴ�":act="npc/2.asp";break;case "����Ṧ":act="npc/3.asp";break;
case "��ȡ�Ṧ":act="npc/4.asp";break;case "Ѱ������":act="npc/5.asp";break;
case "������":act="sjfunc/2001.asp";break;case "�����":act="sjfunc/2003.asp";break;
case "�ⶾ��":act="sjfunc/2004.asp";break;case "ƽ����":act="sjfunc/2005.asp";break;
case "͵����":act="sjfunc/2006.asp";break;case "ҡǮ��":act="sjfunc/2008.asp";break;
case "�����":act="sjfunc/2009.asp";break;case "˧����":act="sjfunc/2010.asp";break;
case "������":act="sjfunc/2011.asp";break;case "���黷":act="sjfunc/2012.asp";break;
case "������":act="sjfunc/2014.asp";break;case "�ݴ�����":act="sjfunc/1001.asp";break;
case "��ȡ����":act="sjfunc/1002.asp";break;case "������":act="sjfunc/1003.asp";break;
case "������":act="sjfunc/1004.asp";break;case "��ť��":act="sjfunc/1005.asp";break;
case "������ť":act="sjfunc/1006.asp";break;case "���°�ť":act="sjfunc/1007.asp";break; 
case "˵��":act="ico/02.asp";break;case "��ҹ��":act="sjfunc/505.asp";break;
case "�ķ���":act="sjfunc/1012.asp";break;case "���Ṧ":act="sjfunc/1013.asp";break;
case "��������":act="sjfunc/1014.asp";break;case "�鿴����":act="ico/09.asp";break;
case "С��":act="sjfunc/98.asp";break;case "����":act="sjfunc/97.asp";break;
case "ԭ�ӵ�":act="sjfunc/100.asp";break;case "��������":act="sjfunc/101.asp";break;
case "��ľ����":act="sjfunc/102.asp";break;case "��ˮ����":act="sjfunc/103.asp";break;
case "��������":act="sjfunc/104.asp";break;case "��������":act="sjfunc/105.asp";break;
case "�ͽ�����":act="sjfunc/106.asp";break;case "��ľ����":act="sjfunc/107.asp";break;
case "��ˮ����":act="sjfunc/108.asp";break;case "�ͻ�����":act="sjfunc/109.asp";break;
case "��������":act="sjfunc/110.asp";break;case "�Ľ�����":act="sjfunc/111.asp";break;
case "��ľ����":act="sjfunc/112.asp";break;case "��ˮ����":act="sjfunc/113.asp";break;
case "�Ļ�����":act="sjfunc/114.asp";break;case "��������":act="sjfunc/115.asp";break;
case "���˷�":act="sjfunc/116.asp";break;case "��������":act="sjfunc/117.asp";break;
case "�����Ṧ":act="sjfunc/118.asp";break;case "���ڵ���":act="sjfunc/119.asp";break;
case "��������":act="sjfunc/120.asp";break;case "��������":act="sjfunc/202.asp";break;
case "��������":act="sjfunc/203.asp";break;case "���ŷ���":act="sjfunc/206.asp";break;
case "���ŷ���":act="sjfunc/207.asp";break;case "������":act="sjfunc/99.asp";break;
case "���˿�":act="f2/dpk-ask.asp";break;case "����":act="f2/dpkfp.asp";break;
case "���齫":act="f2/dmj-ask.asp";break;case "����":act="f2/DMJFP.ASP?action=1";break;
case "����":act="f2/DMJFP.ASP?action=2";break;case "����":act="f2/DMJFP.ASP?action=3";break;
case "����":act="f2/DMJFP.ASP?action=4";break;case "����":act="f2/DMJFP.ASP?action=5";break;
case "����":act="f2/DMJFP.ASP?action=6";break;case "���չٸ�":act="sjfunc/5042.asp";break;
case "͵���":act="sjfunc/93.asp";break;case "�Զ�����":act="sjfunc/zjn_add_fun_doub.asp";break;
case "�������":act="sjfunc/tw.asp";break;case "���ɴ�ս":act="empdz/mp1.asp";break;
case "�ᱦ��ս":act="sjfunc/db14.asp";break;case "�ᱦ�¶�":act="sjfunc/db10.asp";break;
case "�ᱦͶ��":act="sjfunc/db13.asp";break;case "�ᱦС������":act="sjfunc/dbxhgj.asp";break;
case "����ᱦ":act="sjfunc/db73.asp";break;case "�ᱦ�����Ա�":act="sjfunc/db85.asp";break;
case "�ᱦʤ��":act="sjfunc/dbsl.asp";break;case "�Ͻ��«":act="sjfunc/db35.asp";break;
case "����":act="sjfunc/02.asp";break;case "�����ӵ�":act="sjfunc/450017.asp";break;
case "�ػ�ھ�":act="sjfunc/hhkj.asp";break;case "���ֻش�":act="sjfunc/mshc.asp";break;
case "�вƽ���":act="sjfunc/zcjb.asp";break;case "���Ĵ�":act="sjfunc/yxdf.asp";break;
case "��ʯ�ɽ�":act="sjfunc/dscj.asp";break;case "�����Ϸ�":act="sjfunc/xynf.asp";break;
case "������ɽ":act="sjfunc/xyns.asp";break;case "�ᱦ���﹥��":act="sjfunc/db73.asp";break;
case "������ս":act="empdz/mpopen.asp";break;case "�������":act="empdz/mpopen1.asp";break;
case "Ѫ����":act="sjfunc/45014.asp";break;case "���鵶":act="sjfunc/450019.ASP";break;
case "������":act="sjfunc/45016.asp";break;case "û��ħ��":act="sjfunc/450152.asp";break;
case "Ѱ��ħ��":act="sjfunc/45017.asp";break;case "���Ʊ�ʯ":act="sjfunc/45018.asp";break;
case "���յ���":act="sjfunc/450018.ASP";break;case "����ħ��":act="sjfunc/1008.asp";break;
case "�ƶ�ħ��":act="sjfunc/1008.asp";break;case "��ťħ��":act="sjfunc/1008.asp";break;
case "ħ����ʯ":act="sjfunc/45006.asp";break;case "�ٱ���ͨ":act="sjfunc/45007.asp";break;
case "�ᱦ��ԯ":act="sjfunc/14dbxx.asp";break;case "�����ͻ�":act="sjfunc/yqsh.asp";break;
case "��ը����":act="sjfunc/ckzd.asp";break;case "��ըӲ��":act="sjfunc/gpzd.asp";break;case "���ջ���":act="sjfunc/45.asp";break;
default:alert('��������,�޷�ִ��')
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
//this.f2.document.af.sjjhaction.value=act;
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
this.title.document.writeln("<body bgcolor='#b7d4f1' oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0><div align='center'>");
this.title.document.writeln(title+"</div></body></html>");
this.title.document.close();}
function DB(){this.mess.document.open();
wdmess("<html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=gb2312\"><\/head>");
wdmess("<style type=\"text/css\">body{font-family:\"����\";color:blue;font-size:10.5pt;line-height:15pt;text-align:center}a{color:blue;text-decoration:none;}a:hover{color:blue;text-decoration:underline;}</style>");
wdmess("<body bgcolor='#eaeaff' TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0>");
var i=0;i=Math.ceil(Math.random()*(Banner.length-1));
wdmess("<marquee scrollamount=3 scrolldelay=10 onmouseover=this.stop(); onmouseout=this.start();>" + Banner[i] + "<\/marquee><\Script Language=JavaScript>setTimeout('parent.DB()',100000);<\/script></body></html>");
this.mess.document.close();}

function tbclutch(){if(this.f2.document.af.tbclutch.value=='ȫ��'){this.f2.document.af.tbclutch.value='��ֱ';this.msgfrm.rows = "*";this.msgfrm.cols = "*";tbclu=false;}else{if(this.f2.document.af.tbclutch.value=='��ֱ'){this.f2.document.af.tbclutch.value='ˮƽ';this.msgfrm.cols = "*,*";this.msgfrm.rows = "*";tbclu=true;}else{this.f2.document.af.tbclutch.value='ȫ��';this.msgfrm.cols = "*";this.msgfrm.rows = "*,*";tbclu=true;}}this.f2.document.af.sytemp.focus();}
function tbmd(){if(this.f2.document.af.tbmd.value=='��'){this.f2.document.af.tbmd.value='��';this.tbgn1.cols="1*,0";}else{this.f2.document.af.tbmd.value='��';this.tbgn1.cols="1*,160";}this.f2.document.af.sytemp.focus();}
function tbygn(){if(this.f2.document.af.tbygn.value=='��'){this.f2.document.af.tbygn.value='��';this.tbymd.rows="0,*,0,0" ;}else{this.f2.document.af.tbygn.value='��';this.tbymd.rows="0,*,0,105" ;}this.f2.document.af.sytemp.focus();}
function tbgn(){if(this.f2.document.af.tbgn.value=='���ܿ�'){this.f2.document.af.tbgn.value='���ܹ�';this.mainfrm.cols="34,*";}else{this.f2.document.af.tbgn.value='���ܿ�';this.mainfrm.cols="0,*";}this.f2.document.af.sytemp.focus();}
self.onerror=null;
var nullframe = '<HTML><BODY BGCOLOR=#000000 text=#ffffff><center><H3 color=yellow><br><font color=yellow>��ӭ������<%=Application("sjjh_chatroomname")%>��վ��</font></h3><br><font size=4>�������ڴӷ�������ȡ����, ���Ժ� ......</font></center></BODY></HTML>';
</script>
</head>
<frameset cols="1,*,1,160,1" name=tbgn1 rows="*" border="0" framespacing="0" frameborder="NO">
<frameset rows="47,*,72" cols="*" frameborder="no" name=tbymd7>
<frame src="ico/wfy24.htm" name="llmbt1" scrolling="no">
<frame src="ico/wfy26.htm" name="llmbt2" scrolling="no">
<frame src="ico/wfy25.htm" name="llmbt3" scrolling="no">
</frameset>

<frameset rows="24,*,0,0,0,25,76,0,1" cols="*">
<frameset name=msgfrm cols="*">
<frame src="ico/eline_time.htm" name="yt" scrolling="no" marginwidth="0" marginheight="0">
<frameset name=msgfrm rows="*,*" cols="*">
//<frame src="about:blank" name="f0" scrolling="AUTO" marginheight="3" marginwidth="5" frameborder="yes">
//<frame src="about:blank" name="f1" scrolling="AUTO" marginheight="3" marginwidth="5" frameborder="no">
<frame src="javascript:parent.nullframe" name="f0" scrolling="AUTO" framespacing="0" marginheight="3" marginwidth="5" frameborder="yes">
<frame src="javascript:parent.nullframe" name="f1" scrolling="AUTO" framespacing="0" marginheight="3" marginwidth="5" frameborder="no">
</frameset>

<frame src="guanggao.asp" scrolling="NO"  name="gg" marginwidth="3" marginheight="3">
<frame src="about:blank" name="mess" scrolling="no" >
<frame src="about:blank" name="t" marginwidth="5" marginheight="5"  scrolling="NO">
<frame src="about:blank" name="title" scrolling="no" >
<frame src="f22.asp" name="f2" scrolling="NO" marginwidth="3" marginheight="8">
<frame src="about:blank" name="d" scrolling="NO">
<frame src="ico/wfy15.htm" name="zt" scrolling="no" marginwidth="0" marginheight="0">
</frameset>
<frameset rows="47,*,72" cols="*" frameborder="no" name=tbymd>
<frame src="ico/wfy20.htm" name="lmbt1" scrolling="no">
<frame src="ico/wfy19.htm" name="lmbt2" scrolling="no">
<frame src="ico/wfy18.htm" name="lmbt3" scrolling="no">
</frameset>
<frameset rows="1,0,*,0,182" name=tbymd>
<frame src="ico/wfy16.htm" name="st" scrolling="no">
<frameset cols="*" name=tbymd>
<frame src="about:blank" name="m">
<frameset cols="*" name=tbymd>
<frame src="about:blank" marginwidth="5" marginheight="5" name="f3">
<frameset cols="*" name=tbymd>
<frame src="about:blank" marginwidth="5" marginheight="5" name="ps">
<frameset cols="*" name=tbymd>
<frame src="F4.asp" name="F4" scrolling="no">
</frameset>
<frameset rows="47,*,72" cols="*" frameborder="no" name=tbymd>
<frame src="ico/wfy23.htm" name="rmbt1" scrolling="no">
<frame src="ico/wfy22.htm" name="rmbt2" scrolling="no">
<frame src="ico/wfy21.htm" name="rmbt3" scrolling="no">
</frameset>
</frameset>

</noframes>
</html>