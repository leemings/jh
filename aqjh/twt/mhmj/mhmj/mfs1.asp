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
if InStr(http,"mhmj/mfs.asp")=0 then 
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
gwm="�쳾����"
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
if nl<20000 then 
	Response.Write "<script Language=Javascript>alert('��ʾ������С��20000���ܽ��룡');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if fl<120000 then 
	Response.Write "<script Language=Javascript>alert('��ʾ������С��120000���ܽ��룡');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
dim js(32)
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
js(10) ="ľ̿"
js(11) ="ľ̿"
js(12) ="��ʬѪ"
js(13) ="�̱�ʯ"
js(14) ="�̱�ʯ"
js(15) ="����"
js(16) ="��ͷ"
js(17) ="��ʬѪ"
js(18) ="�챦ʯ"
js(19) ="��ʬѪ"
js(20) ="����ʯ"
js(21) ="ľ̿"
js(22) ="�̱�ʯ"
js(23) ="���"
js(24) ="���"
js(25) ="����ʯ"
js(26) ="ľ̿"
js(27) ="ľ̿"
js(28) ="��ʬ��"
js(29) ="����"
js(30) ="��ʬ��"
js(31) ="����"
randomize()
myxy = Int(Rnd*32)
s1=Int(Rnd*20000)
s2=Int(Rnd*10000)
s3=Int(Rnd*150)
s4=Int(Rnd*4)
s5=Int(Rnd*2)
s6=Int(Rnd*300)
zstemp=add(rs("w6"),js(myxy),10)
dim js1(24)
js1(0) ="���Ƶ���"
js1(1) ="��ҩ"
js1(2) ="�׷�"
js1(3) ="��ײ�"
js1(4) ="�ƾ�"
js1(5) ="����"
js1(6) ="��ҩ"
js1(7) ="��ҩ"
js1(8) ="��֥��"
js1(9) ="�ƾ�"
js1(10) ="�׷�"
js1(11) ="�ƾ�"
js1(12) ="���Ƶ���"
js1(13) ="��ҩ"
js1(14) ="��ҩ"
js1(15) ="�ƾ�"
js1(16) ="��ײ�"
js1(17) ="����"
js1(18) ="��֥��"
js1(19) ="��ײ�"
js1(20) ="��ҩ"
js1(21) ="�����"
js1(22) ="�׷�"
js1(23) ="��ײ�"
randomize()
myxx = Int(Rnd*24)
zstemp1=add(rs("w1"),js1(myxx),22)
dim js2(40)
js2(0) ="̫������"
js2(1) ="�¹����"
js2(2) ="������"
js2(3) ="������"
js2(4) ="̫����"
js2(5) ="���佣"
js2(6) ="������"
js2(7) ="��Ů��"
js2(8) ="̫����ָ"
js2(9) ="̫������"
js2(10) ="�¹��ָ"
js2(11) ="�¹�����"
js2(12) ="̫������"
js2(13) ="�¹Ᵽ��"
js2(14) ="��ѥ"
js2(15) ="�ƽ�ѥ"
js2(16) ="��ñ"
js2(17) ="�ƽ�ñ"
js2(18) ="���齣"
js2(19) ="ʤа��"
js2(20) ="����"
js2(21) ="��΢��"
js2(22) ="���ƽ�"
js2(23) ="��ֽ�"
js2(24) ="���߽�"
js2(25) ="���潣"
js2(26) ="��ˮ��"
js2(27) ="ԽŮ��"
js2(28) ="���ӽ�"
js2(29) ="������"
js2(30) ="ʥȪ��"
js2(31) ="ˮ������"
js2(32) ="�ƽ����"
js2(33) ="��������"
js2(34) ="��ͭ����"
js2(35) ="������"
js2(36) ="��ͭ��"
js2(37) ="ˮ��ñ"
js2(38) ="̫��ñ"
js2(39) ="����ָ"
randomize()
myxa = Int(Rnd*41)
zstemp2=add(rs("w3"),js2(myxa),1)
dim js3(41)
js3(0) ="��Ѫ��"
js3(1) ="������"
js3(2) ="������"
js3(3) ="������"
js3(4) ="��Ѫ��"
js3(5) ="�Զ���"
js3(6) ="���ѿ�"
js3(7) ="��Ѫ��"
js3(8) ="������"
js3(9) ="������"
js3(10) ="��鿨"
js3(11) ="�Զ���"
js3(12) ="���ѿ�"
js3(13) ="���׿�"
js3(14) ="�氮��"
js3(15) ="�崿��"
js3(16) ="���￨"
js3(17) ="����"
js3(18) ="����"
js3(19) ="����"
js3(20) ="��Ǯ��"
js3(21) ="��Ѫ��"
js3(22) ="������"
js3(23) ="������"
js3(24) ="����"
js3(25) ="�Զ���"
js3(26) ="���ѿ�"
js3(27) ="��Ѫ��"
js3(28) ="������"
js3(29) ="������"
js3(30) ="���ҿ�"
js3(31) ="�Զ���"
js3(32) ="���ѿ�"
js3(33) ="���׿�"
js3(34) ="��˰��"
js3(35) ="�崿��"
js3(36) ="���￨"
js3(37) ="����"
js3(38) ="����"
js3(39) ="����"
js3(40) ="�ӵ㿨"
randomize()
myxb = Int(Rnd*41)
zstemp3=add(rs("w5"),js3(myxb),1)
dim js4(28)
js4(0) ="�Ƴ�"
js4(1) ="������"
js4(2) ="������"
js4(3) ="������"
js4(4) ="������"
js4(5) ="�Ż���¶��"
js4(6) ="�����۶�"
js4(7) ="��̥�׽���"
js4(8) ="������"
js4(9) ="����˪"
js4(10) ="���ҩ"
js4(11) ="ըҩ"
js4(12) ="������"
js4(13) ="��ʬˮ"
js4(14) ="�Ƴ�"
js4(15) ="����ɢ"
js4(16) ="������"
js4(17) ="������"
js4(18) ="��Ѫ��"
js4(19) ="�Ż���¶��"
js4(20) ="�����۶�"
js4(21) ="��̥�׽���"
js4(22) ="������"
js4(23) ="����˪"
js4(24) ="���ҩ"
js4(25) ="ըҩ"
js4(26) ="������"
js4(27) ="��ʬˮ"
randomize()
myxc = Int(Rnd*28)
zstemp4=add(rs("w8"),js4(myxc),1)
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
conn.execute "update �û� set ״̬='��',�¼�ԭ��='����|׳��"' where ����='" & aqjh_name &"'"
end if
if gwtl<0 then
conn.execute "update ���� set ״̬='����',����=0,����=0,ɱ��=0,����=0,��ɱ��='"&aqjh_name&"' where ����='"&gwm&"'"
conn.execute "update �û� set ���=���+70,allvalue=allvalue+"&s3&",w1='"&zstemp1&"',w6='"&zstemp&"',w3='"&zstemp2&"',w5='"&zstemp3&"',w8='"&zstemp4&"',����ʱ��=now() where ����='"&aqjh_name&"'"
end if
if gwtl<0 then
	Response.Write "<script Language=Javascript>alert('��ʾ��"&aqjh_name&"��ϲ�㣡"&gwm&"�Ѿ�����ɹ����𣬻��"&js2(myxa)&"1����"&js(myxy)&"10����"&js1(myxx)&"22������Ƭ"&js3(myxb)&"1�ţ���ҩ"&js4(myxc)&"1�������70�������֣���㣩"&s3&"�㣡��');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('����������"&gwm&"�����½�"&ss&"�㣬ɱ���½�"&s6&"�㣬��������½�"&ss1&"�㣡');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=RED><b>���λ�ħ����Ϣ��</b></font><font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]����һ����ս���ѽ�[<font color=red><b>"&gwm&"</b></font>]���𣬲��Ҵӹ������ϻ��[<font color=red><b>"&js(myxy)&"</b></font>]10����[<font color=red><b>"&js1(myxx)&"</b></font>]22����[<font color=red><b>"&js2(myxa)&"</b></font>]1����[<font color=red><b>"&js3(myxb)&"</b></font>]1�š���ҩ[<font color=red><b>"&js4(myxc)&"</b></font>]1���ʹ��[<font color=red><b>"&s3&"</b></font>]�㣬���[<font color=red><b>70</b></font>]����<font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]��<font color=red><b>����Ա�ȼ���</b></font>Ϊ[<font color=red><b>"&hydj&"</b></font>]�����������λ�ħ���<font color=red><b>��ʱ�����ơ�</b></font>Ϊ��[<font color=red><b>"&jgsj&"</b></font>]��</font>"			'��������
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
