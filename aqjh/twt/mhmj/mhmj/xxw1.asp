<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then 
Response.Write "<script language=javascript>{alert('��ʾ����û�е�½���Ѿ���ʱ�Ͽ����ӣ������µ�½��');parent.history.go(-1);}</script>" 
Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"mhmj/xxw.asp")=0 then 
Response.Write "<script language=javascript>{alert('�ύ���ݷǷ�������ܾ�ִ�У�');parent.history.go(-1);}</script>" 
Response.End 
end if
if Minute(time())<40 or Minute(time())>=60 then
Response.Write "<script Language=Javascript>{alert('��ʾ���λ�ħ��ֻ��ÿ��Сʱ�ĺ�20���ӲŻῪ��');parent.history.go(-1);}</script>" 
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gwm="������Ѫ��"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
tl=rs("����")
nl=rs("����")
fl=rs("����")
dj=rs("�ȼ�")
hydj=rs("��Ա�ȼ�")
gj=rs("����")
fy=rs("����")
wg=rs("�书")
allss=int(gj+fy+wg)
select case hydj
	case 0
		jgsj=320
	case 1
		jgsj=260
	case 2
		jgsj=220
	case 3
		jgsj=180
	case 4
		jgsj=130
	case 5
		jgsj=100
	case 6
		jgsj=60
end select
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<jgsj then
	s=jgsj-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������"&hydj&"����Ա��"&jgsj&"��ɽ���һ�Σ����"&s&"��������Գ���');parent.history.go(-1);}</script>" 
	Response.End
end if
if tl<0 then 
	Response.Write "<script Language=Javascript>alert('��ʾ�����������Ѿ�׳���ˡ����´�ע���ˣ�');location.href = '../../../exit.asp';</script>"
	Response.End
end if
if nl<18000 then 
	Response.Write "<script Language=Javascript>alert('��ʾ������С��18000���ܽ��룡');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if fl<60000 then 
	Response.Write "<script Language=Javascript>alert('��ʾ������С��60000���ܽ��룡');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
dim js(21)
js(0) ="��ͷ"
js(1) ="�̱�ʯ"
js(2) ="ˮ��ʯ"
js(3) ="����"
js(4) ="����ʯ"
js(5) ="ľ̿"
js(6) ="����"
js(7) ="���"
js(8) ="�챦ʯ"
js(9) ="��ʬ��"
js(10) ="��ʬѪ"
js(11) ="����"
js(12) ="�챦ʯ"
js(13) ="��ʬ��"
js(14) ="����"
js(15) ="�����"
js(16) ="������"
js(17) ="�̱�ʯ"
js(18) ="��ˮ"
js(19) ="ľͷ"
js(20) ="��ʯ"
randomize()
myxy = Int(Rnd*21)
s1=Int(Rnd*20000)
s2=Int(Rnd*10000)
s3=Int(Rnd*150)
s4=Int(Rnd*4)
s5=Int(Rnd*2)
s6=Int(Rnd*300)
zstemp=add(rs("w6"),js(myxy),7)
dim js1(16)
js1(0) ="���Ƶ���"
js1(1) ="��ҩ"
js1(2) ="�׷�"
js1(3) ="��ײ�"
js1(4) ="�ƾ�"
js1(5) ="����"
js1(6) ="��֥��"
js1(7) ="�����"
js1(8) ="ֹѪ��"
js1(9) ="Ů����"
js1(10) ="�׾�"
js1(11) ="����"
js1(12) ="����"
js1(13) ="����"
js1(14) ="�����"
js1(15) ="���Ǿ�"
randomize()
myxx = Int(Rnd*16)
zstemp1=add(rs("w1"),js1(myxx),30)
rs.close
rs.open "select * from ���� where ����='"&gwm&"'",conn
gwdj=rs("�ȼ�")
zt=rs("״̬")
gwtl=rs("����")
lsz=rs("��ɱ��")
gjto=rs("����")
fyto=rs("����")
ssto=rs("ɱ��")
allssto=int(gjto+fyto+ssto)
if dj<gwdj then 
	Response.Write "<script Language=Javascript>alert('��ʾ����ĵȼ�С�ڹ���ȼ����޷���ɱ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if zt="����" then 
	Response.Write "<script Language=Javascript>alert('��ʾ��"&gwm&"�ѱ���ɱ��"&lsz&"��ɱ������������Ƚ��份�ѣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
if allss<allssto then
       ss=0
       ss1=int(allssto-allss)
else
       ss=int(allss-allssto)
       ss1=ss/4
end if
conn.execute "update ���� set ����=����-"&ss&",ɱ��=ɱ��-"&s6&" where ����='"&gwm&"'"
conn.execute "update �û� set ����=����-"&ss1&" where ����='"&aqjh_name&"'"
rs.open "select * from ���� where ����='"&gwm&"'",conn
zt=rs("״̬")
gwtl=rs("����")
lsz=rs("��ɱ��")
if zt="����" then 
	Response.Write "<script Language=Javascript>alert('��ʾ��"&gwm&"�ѱ���ɱ��"&lsz&"��ɱ������������Ƚ��份�ѣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if tl<0 then
conn.execute "update �û� set ״̬='��',�¼�ԭ��='����|"&fn1&"' where ����='" & aqjh_name &"'"
conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "','����','����',now(),'�λ�ħ��')"
end if
if gwtl<0 then
conn.execute "update ���� set ״̬='����',����=0,����=0,ɱ��=0,����=0,��ɱ��='"&aqjh_name&"' where ����='"&gwm&"'"
conn.execute "update �û� set ���=���+25,allvalue=allvalue+"&s3&",w1='"&zstemp1&"',w6='"&zstemp&"',����ʱ��=now() where ����='"&aqjh_name&"'"
end if
if gwtl<0 then
	Response.Write "<script Language=Javascript>alert('��ʾ��"&aqjh_name&"��ϲ�㣡"&gwm&"�Ѿ�����ɹ����𣬻��"&js(myxy)&"7����"&js1(myxx)&"30�������25�������֣���㣩"&s3&"�㣡��');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('����������"&gwm&"�����½�"&ss&"�㣬ɱ���½�"&s6&"�㣬��������½�"&ss1&"�㣡');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=RED><b>���λ�ħ����Ϣ��</b></font><font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]����һ����ս���ѽ�[<font color=red><b>"&gwm&"</b></font>]���𣬲��Ҵӹ������ϻ��[<font color=red><b>"&js(myxy)&"</b></font>]7����[<font color=red><b>"&js1(myxx)&"</b></font>]30���ʹ��[<font color=red><b>"&s3&"</b></font>]�㣬���[<font color=red><b>25</b></font>]����<font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]��<font color=red><b>����Ա�ȼ���</b></font>Ϊ[<font color=red><b>"&hydj&"</b></font>]�����������λ�ħ���<font color=red><b>��ʱ�����ơ�</b></font>Ϊ��[<font color=red><b>"&jgsj&"</b></font>]��</font>"			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
%>
