<%@ LANGUAGE=VBScript codepage ="936" %>

<!--#include file="sjfunc/sjfunc.asp"-->

<!--#include file="sjfunc/func.asp"-->
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Application("sjjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����С���ǲ���û������ѽ�������й���');</script>"
	response.end
end if
id=LCase(trim(Request.QueryString("id")))
tsk=int(abs(clng(Request.QueryString("tsk"))))
tempjs=int(abs(clng(Application("sjjh_kl"))))
if tempjs<>tsk then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� where ״̬='��ʬ��' and ���='��ʬ��' order by ���� ",conn
sf=rs("���")
attackname=rs("����")
tl=rs("����")
if sjjh_jhdj>30 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ս��30�����������ʸ��ʬ��');window.close();}</script>"
	Response.End 
end if
Application.Lock
Application("sjjh_kl")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.execute "update �û� set ����=����-"&tempjs&" where ����='" & sjjh_name &"'"
rs.open "select ����,w6 FROM �û� WHERE ����='" & sjjh_name &"'",conn
if rs("����")<-100000 then
	conn.execute "update �û� set ״̬='��ʬ��',���='��ʬ��',�¼�ԭ��='��ʬ��|"&fn1&"' where ����='" & sjjh_name &"'"
    conn.execute "update �û� set ����=5000,����=5000,�¼�ԭ��='��',״̬='����',���='����' where ����='" & attackname &"'"
  	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'��ʬ��','���´�ʬ������','����')"
	kl="<img src='pic/kl.gif'>["&sjjh_name&"]̫˥�ˣ�Ϊ�˸���ʬ��" & attackname & "��������Լ���������ʬ��"&id&"ҧ��,["&sjjh_name&"]����˽�ʬ��." & attackname & "��ʽ�����λ��Ϻ��ϲ����" & attackname & "�����ڽ�ʬ��þ���..��������"
	call boot(sjjh_name,"��ʬ���������ߣ���ʬ��������ѽ������")
else
dim js(10)
js(0) ="��ͷ"
js(1) ="����ʯ"
js(2) ="�챦ʯ"
js(3) ="��ʬѪ"
js(4) ="��ʬ��"
js(5) ="ˮ��ʯ"
js(6) ="�̱�ʯ"
js(7) ="���"
js(8) ="����"
js(9)="ľ̿"
randomize()
myxy = Int(Rnd*10)
zstemp=add(rs("w6"),js(myxy),2)
conn.execute "update �û� set ����=����+"& tempjs*1 &",w6='"&zstemp&"' where ����='" & sjjh_name & "'"
kl="��call,Ӣ�ۣ�����Ӣ�ۣ�["&sjjh_name&"]�뽩ʬ��" & attackname & "�����֡�����ʬ�����ڵ����ˣ������Լ�Ҳ�˵Ĳ��ᣬ�ٸ�����<img src='img/251.gif'>"&tempjs*1&"�����ڴ�ܽ�ʬ����"&sjjh_name&"�õ���[<b><font color=red>"&js(myxy)&"</font></b>]2������"

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>����ʬ����</b></font>"&kl

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
