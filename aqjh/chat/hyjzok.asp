<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
 Response.Write "<script language=javascript>{alert(''''��ʾ����["&chatinfo(0)&"]���䲻���Է��У�'''');}</script>"
 Response.End
end if
f=Minute(time()) 
%>
<div align="center"> 
  <p><font color="#FFFFFF"><span style='font-size:9pt'><font size="3">��</font></span></font><a href="hyjz.asp" target="f3"><font size="3" color="#FF0000">��Ա����</font></a><font size="3" color="#FFFFFF">��</font>
  </p>
</div>
<%
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
 Response.Write "<script language=javascript>{alert(''''�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�'''');}</script>" 
 Response.End 
end if 
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
 Response.Write "<script Language=javascript>alert(''''�㲻�ܽ��в��������д˲���������������ң�'''');</script>"
 Response.End
end if
to1=trim(request.form("to1"))
hyjzzs=trim(request.form("hyjzzs"))
tl=int(abs(clng(request.form("tl"))))
nl=int(abs(clng(request.form("nl"))))
maxtl=int(abs(clng(request.form("maxtl"))))
maxnl=int(abs(clng(request.form("maxnl"))))
money=int(abs(clng(request.form("money"))))
zsdj=int(abs(clng(request.form("dj"))))
if Instr(dgjjzs,chr(39))<>0 or Instr(dgjjzs,chr(34))<>0 then
	Response.Write "<script Language=Javascript>alert('��Ҫ��ʲô��Զ�㣡');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('���ڲ��ǻ�Ա����ʱ�䣬��Ա�ؼ�Ϊ1-25���ӣ�');window.close();}</script>"
	Response.End 
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('�����ڡ����롱״̬�У���ʹ�á��һ����������ܽ����');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('"&to1&"���ڡ����롱״̬�У��벻Ҫ������');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if tl<maxtl or nl<maxnl then
	Response.Write "<script Language=Javascript>alert('��ʾ��Ҫʹ��["&dgjjzs&"]������Ҫʹ��"&maxtl&"������"&maxnl&"������');</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������˲��ڽ����У�');</script>"
	Response.End
end if
if to1="���" or to1=Application("aqjh_automanname") then
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻���ԶԴ�һ򽭺������˷��С�');</script>"
	Response.End
end if
if InStr(";" & Application("aqjh_npc"), ";" & to1 & "|")<>0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܶ�NPCʹ�ô˲�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name & "'",conn,2,2
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ոձ�����ɱ���������������ɣ�');}</script>"
	Response.End
end if
hy=rs("��Ա�ȼ�")
pdhy=rs("��Ա")
if hy=0 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ѽ��["&aqjh_name &"]��ɲ��ǻ�Ա������ʹ�û�Ա���У�');window.close();</script>"
 response.end
end if
mp=rs("����")
if rs("grade")>=6 and rs("grade")<10 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ���Ա������ʹ��["&hyjzzs&"]��');</script>"
	myclose()
end if
if rs("����")<money then
	Response.Write "<script Language=Javascript>alert('��ʾ����û��"&money&"�����ӣ�����ʹ��["&hyjzzs&"]��');</script>"
	myclose()
end if
if rs("֪��")<15 then
	Response.Write "<script Language=Javascript>alert('��ʾ����û��15��֪�ʣ�����ʹ��["&dgjjzs&"]��');</script>"
	myclose()
end if
if rs("����")<nl or rs("����")<tl then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������������㣬���ܷ����У�');</script>"
	myclose()
end if
if rs("����")=True and aqjh_grade<10 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������������У������Դ�ܣ�');</script>"
	myclose()
end if
if rs("�ȼ�")<50 then
	Response.Write "<script Language=Javascript>alert('��ʾ����ĵȼ�����[50],������ʹ�ô��У�');</script>"
	myclose()
end if
if rs("ɱ����")>=int(Application("aqjh_killman")) then
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
tlsx=rs("�ȼ�")*aqjh_tlsx+5260+rs("������")
nlsx=rs("�ȼ�")*aqjh_nlsx+2000+rs("������")
wgsx=rs("�ȼ�")*aqjh_wgsx+3800+rs("�书��")
tlbf=int(tl/tlsx)
nlbf=int(nl/nlsx)
wgbf=int(lxjwg1/wgsx)
bf=int((nlbf+tlbf+wgbf)/3)
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn,2,2
zstt=rs("ת��")
if zstt>=8 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է���ת��8���ˣ������˲��˶Է���');}</script>"
	Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
dlsj=DateDiff("n",rs("ɱ��ʱ��"),now())
if dlsj<30 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ոձ�ɱ�������30���Ӻ�ɣ�');</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<8 and rs("����")="��" and aqjh_grade<10 then
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
if rs("�ȼ�")<40 then
	Response.Write "<script Language=Javascript>alert('��ʾ����������40�����㲻Ҫ��ô�ݣ�');</script>"
	myclose()
end if
if rs("����")=true and hyjzzs<>"�������" then
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
qq1=lxjwg1+lxjgj1+100000-lxjwg2+lxjgj2
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
conn.execute "update �û� set ֪��=֪��-15,����=����-100,����=����-"&money&",����=����-" & tl & ",����=����-" & nl & ",����ʱ��=now() where ����='" & aqjh_name & "'"
e=""
if shengtili<-100 then
	conn.execute "update �û� set ֪��=֪��-15,ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & aqjh_name & "'"
	if rs("����")=Application("aqjh_baowuname") then
		conn.execute "update �û� set ��������=0,����='"& Application("aqjh_baowuname") &"' where id="&aqjh_id
		conn.execute "update �û� set ��������=0,����='��',����=100,����=2000 where ����='"& to1 &"'"
		stunt=aqjh_name & "��"& to1 &"�ı���:"& Application("aqjh_baowuname") &"���ߡ�����������Ҫ���������ſ��Եõ�����Ķ�����"
	else
	if hyjzzs="�������" then
		if lgbh=true then
			conn.execute "update �û� set ����=false,����=����+" & killer & ",����ʱ��=now() where ����='" & to1 & "'"
			lgbh=false
			e="������Ҫ����ù�ˡ�"
			stunt=aqjh_name & "<bgsound src=wav/dgjj0.wav loop=1>����" & nl & "��������" & tl & "����������<img src='img/41.gif'><font color=blue>" & to1 & "</font>ʹ���˽����Ͼ��Ѵ��Ļ�Ա����֮<font color=008000>��"&hyjzzs&"��</font>��ȥ����<font color=blue>" & to1 & "</font>��<font color=red>��������<font/>��" & e
		else
			conn.execute "update �û� set ɱ��ʱ��=now(),״̬='��',����=0 where ����='" & to1 & "'"
			conn.execute "update �û� set ֪��=֪��-15,allvalue=allvalue+50,����=����+" & yinliang2 & " where ����='" & aqjh_name & "'"
			
			e="�㣬" & to1 & "<bgsound src=wav/si.wav loop=1>������<img src=xx/gif/WG8.GIF>������ȥ�����Ӵ˽�����������һֻ��Ϻ," & aqjh_name & "�õ�" & to1 & "�������ֽ�" & yinliang2 & "�������[50]�㣬" & to1 & "��������Ʒ�Ӵ�ɢ�佭��!"
			fn1=hyjzzs
			call boot(to1,"��Ա���У������ߣ�"&aqjh_name&",["&mp&"]"&fn1)
			conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','��Ա����֮"&hyjzzs&"','����')"
		end if
	else
		conn.execute "update �û� set ɱ��ʱ��=now(),״̬='��',����=0,�¼�ԭ��='"&aqjh_name&"|��Ա����֮"&hyjzzs&"' where ����='" & to1 & "'"
		if tjf=True then
			conn.execute "update �û� set ����=0,���=int(���/2),����=0,����=0 where ����='" & to1 & "'"
		end if
		conn.execute "update �û� set ֪��=֪��-15,allvalue=allvalue+500,����=����+" & yinliang2 & " where ����='" & aqjh_name & "'"
		e="�㣬" & to1 & "<bgsound src=wav/si.wav loop=1>������<img src=xx/gif/WG8.GIF>������ȥ�����Ӵ˽�����������һֻ��Ϻ," & aqjh_name & "�õ�" & to1 & "�������ֽ�" & yinliang2 & "�������[500]�㣡"
		fn1=hyjzzs
		call boot(to1,"��Ա���У������ߣ�"&aqjh_name&"��["&mp&"]"&fn1)
		conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','��Ա����֮"&hyjzzs&"','����')"
	end if
	end if
end if
if lgbh=true then
	conn.execute "update �û� set ֪��=֪��-15,����=����+100,����=����+11002000,����=����+" & tl & ",����=����+" & nl & ",����ʱ��=now() where ����='" & aqjh_name & "'"
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
	stunt=aqjh_name & "<bgsound src=wav/dgjj0.wav loop=2>����" & nl & "��������" & tl & "����������<img src='img/021.gif'><font color=blue>" & to1 & "</font>ʹ���˽����Ͼ��Ѵ��Ļ�Ա����֮<font color=008000>��"&hyjzzs&"��</font>��ɱ��" & killer & e
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>����Ա����"&hyjzzs&"��</b>"& stunt &"</font>"& t
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('��ϲ������"&hyjzzs&"�Ѿ�������ɣ�');</script>"
function myclose()
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end function
%>