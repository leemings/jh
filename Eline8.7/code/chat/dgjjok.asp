<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="sjfunc/chatconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if

chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');}</script>" 
	Response.End 
end if 
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');</script>"
	Response.End
end if
to1=trim(request.form("to1"))
dgjjzs=trim(request.form("dgjjzs"))
tl=int(abs(clng(request.form("tl"))))
nl=int(abs(clng(request.form("nl"))))
maxtl=int(abs(clng(request.form("maxtl"))))
maxnl=int(abs(clng(request.form("maxnl"))))
money=int(abs(clng(request.form("money"))))
zsdj=int(abs(clng(request.form("dj"))))
if Instr(dgjjzs,chr(39))<>0 or Instr(dgjjzs,chr(34))<>0 then
	Response.Write "<script Language=Javascript>alert('��Ҫ��ʲô����Զ�㣡');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('�����ڡ����롱״̬�У���ʹ�á��һ����������ܽ����');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("sjjh_zanli")),LCase("!"&to1&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('"&to1&"���ڡ����롱״̬�У��벻Ҫ������');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if tl<maxtl or nl<maxnl then
	Response.Write "<script Language=Javascript>alert('��ʾ��Ҫʹ��["&dgjjzs&"]������Ҫʹ��"&maxtl&"������"&maxnl&"������');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������˲��ڽ����У�');</script>"
	Response.End
end if
if to1="���" or to1=Application("sjjh_automanname") then
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻���ԶԴ�һ򽭺������˷��С�');</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)<>"���ַ���"  then
Response.Write "<script Language=Javascript>alert('��ʾ:������ֻ�������ͻ��������ϣ����ŵȽ������ɴ�ս�������˵��ڳ����������ɼ�ǿ�����ܵ�[���ַ���]����ȥ�ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select * from �û� where ����='" & sjjh_name & "'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ոձ�����ɱ���������������ɣ�');}</script>"
	Response.End
end if
mp=rs("����")
if rs("grade")>=6 and rs("grade")<10 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ���Ա������ʹ��["&dgjjzs&"]��');</script>"
	myclose()
end if
if rs("����")<money then
	Response.Write "<script Language=Javascript>alert('��ʾ����û��"&money&"�����ӣ�����ʹ��["&dgjjzs&"]��');</script>"
	myclose()
end if

if rs("����")<nl or rs("����")<tl then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������������㣬���ܷ����У�');</script>"
	myclose()
end if
if rs("����")=True and sjjh_grade<10 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������������У������Դ�ܣ�');</script>"
	myclose()
end if
if rs("�ȼ�")<zsdj then
	Response.Write "<script Language=Javascript>alert('��ʾ����ĵȼ�����["&zsdj&"],������ʹ�ô��У�');</script>"
	myclose()
end if
if rs("ɱ����")>=int(Application("sjjh_killman")) then
	Response.Write "<script Language=Javascript>alert('��ʾ�������ɱ���˹����ˣ�������ɱ����');</script>"
	myclose()
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<8 then
	s=8-sj
	Response.Write "<script Language=Javascript>alert('���棺���"& s &"���ٷ���,�ɱ����ţ�');</script>"
	myclose()
end if
lxjwg1=rs("�书")
lxjgj1=rs("����")
mydj=rs("�ȼ�")
nn=int(mydj/10)+1
tlsx=rs("�ȼ�")*sjjh_tlsx+5260+rs("������")
nlsx=rs("�ȼ�")*sjjh_nlsx+2000+rs("������")
wgsx=rs("�ȼ�")*sjjh_wgsx+3800+rs("�书��")
tlbf=int(tl/tlsx)
nlbf=int(nl/nlsx)
wgbf=int(lxjwg1/wgsx)
bf=int((nlbf+tlbf+wgbf)/3)
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<8 and rs("����")="��" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ոձ�����ɱ���������ȷŹ����ɣ�');}</script>"
	Response.End
end if
jhhy=rs("��Ա�ȼ�")
ntnt=rs("�ȼ�")
tjf=rs("ͨ��")
if rs("grade")>=6 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ���Ա��������ʲô����������վ����Ӧ��');</script>"
	myclose()
end if
if rs("�ȼ�")<15 then
	Response.Write "<script Language=Javascript>alert('��ʾ����������15��������������ˣ��㲻�ɹ�������');</script>"
	myclose()
end if
if rs("����")=true and dgjjzs<>"����ʽ" then
	Response.Write "<script Language=Javascript>alert('����������������Ĺ�����Ч���Ժ���˵�ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	myclose()
end if
tomp=rs("����")
lgbh=rs("����")
lxjwg2=rs("�书")
lxjgj2=rs("����")
yinliang2=rs("����")
youid=rs("id")
randomize timer
sss=int(rnd*9)+1
qq1=lxjwg1+lxjgj1+22000-lxjwg2+lxjgj2
qq2=(tl+nl)*sss
qq3=(qq1+qq2)*bf
qq4=sqr(qq3)
killer=int((qq4*nn)/2)
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
killer=killer+(money/50)
randomize timer
m1 = Int(100 * Rnd)+100
gjtl=int(fn1/m1)
killer=killer+gjtl
shengtili=rs("����")-killer
conn.execute "update �û� set ����=����-"  & killer & " where ����='" & to1 & "'"
conn.execute "update �û� set ����=����-100,����=����-"&money&",����=����-" & tl & ",����=����-" & nl & ",����ʱ��=now() where ����='" & sjjh_name & "'"
e=""
if shengtili<-100 then
	conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & sjjh_name & "'"
	if rs("����")=Application("sjjh_baowuname") then
		conn.execute "update �û� set ��������=0,����='"& Application("sjjh_baowuname") &"' where id="&sjjh_id
		conn.execute "update �û� set ��������=0,����='��',����=100,����=2000 where ����='"& to1 &"'"
		stunt=sjjh_name & "��"& to1 &"�ı���:"& Application("sjjh_baowuname") &"���ߡ�����������Ҫ���������ſ��Եõ�����Ķ�����"
	else
	if dgjjzs="����ʽ" then
		if lgbh=true then
			conn.execute "update �û� set ����=false,����=����+" & killer & ",����ʱ��=now() where ����='" & to1 & "'"
			lgbh=false
			e="������Ҫ����ù�ˡ�"
			stunt=sjjh_name & "<bgsound src=wav/bsj.wav loop=1>����" & nl & "��������" & tl & "����������<img src='img/41.gif'><font color=blue>" & to1 & "</font>ʹ���˽����Ͼ��Ѵ��Ķ��¾Ž�֮<font color=008000>��"&dgjjzs&"��</font>��ȥ����<font color=blue>" & to1 & "</font>��<font color=red>��������<font/>��" & e
		else
			conn.execute "update �û� set ״̬='��',����=0 where ����='" & to1 & "'"
			conn.execute "update �û� set ����=����+" & yinliang2 & " where ����='" & sjjh_name & "'"
			
			e="�㣬" & to1 & "<bgsound src=wav/daipu.wav loop=1>�����ĵ�����ȥ�����Ӵ˽�����������һֻ��Ϻ," & sjjh_name & "�õ�" & to1 & "�������ֽ�" & yinliang2 & "����" & to1 & "��������Ʒ�Ӵ�ɢ�佭��!"
			fn1=dgjjzs
			call boot(to1,"���¾Ž��������ߣ�"&sjjh_name&",["&mp&"]"&fn1)
			conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','���¾Ž�֮"&dgjjzs&"','����')"
		end if
	else
		conn.execute "update �û� set ״̬='��',����=0,�¼�ԭ��='"&sjjh_name&"|���¾Ž�֮"&dgjjzs&"' where ����='" & to1 & "'"
		if tjf=True then
			conn.execute "update �û� set ����=0,���=int(���/2),����=0,����=0 where ����='" & to1 & "'"
		end if
		conn.execute "update �û� set ����=����+" & yinliang2 & " where ����='" & sjjh_name & "'"
		e="�㣬" & to1 & "<bgsound src=wav/daipu.wav loop=1>�����ĵ�����ȥ�����Ӵ˽�����������һֻ��Ϻ," & sjjh_name & "�õ�" & to1 & "�������ֽ�" & yinliang2 & "����"
		fn1=dgjjzs
		call boot(to1,"���¾Ž��������ߣ�"&sjjh_name&"��["&mp&"]"&fn1)
		conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','���¾Ž�֮"&dgjjzs&"','����')"
	end if
	end if
end if
if lgbh=true then
	conn.execute "update �û� set ����=����+100,����=����+11002000,����=����+" & tl & ",����=����+" & nl & ",����ʱ��=now() where ����='" & sjjh_name & "'"
    Response.Write "<script Language=Javascript>alert('����������������Ĺ������޷��Ƴ���������������������������˵�ɣ�');</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if e="" then
	e="�㡣"
end if
if stunt="" then
	stunt=sjjh_name & "<bgsound src=wav/bsj.wav loop=1>����" & nl & "��������" & tl & "����������<img src='img/021.gif'><font color=blue>" & to1 & "</font>ʹ���˽����Ͼ��Ѵ��Ķ��¾Ž�֮<font color=008000>��"&dgjjzs&"��</font>��ɱ��" & killer & e
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>�����¾Ž�֮"&dgjjzs&"��</b>"& stunt &"</font>"& t
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
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
Response.Write "<script Language=Javascript>alert('��ϲ������"&dgjjzs&"�Ѿ�������ɣ�');</script>"
function myclose()
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end function
%>