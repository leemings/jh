<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/func.asp"-->
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Application("aqjh_zqyb")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����С�������ˣ��±��类�˳Ե��ˣ�');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_zqyb"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	response.end
end if
Application.Lock
Application("aqjh_zqyb")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update �û� set ����=����+"&tempjs&" where ����='" & aqjh_name &"'"
rs.open "select ����,w5 FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<-10000 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='���±�|"&fn1&"' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'�±�','���±����ɱ��˲���','����')"
	kl="<img src='img/zqyb.gif'>������["&aqjh_name&"]�������±������ǲ�С�ģ�������˸������ˡ�����"
	call boot(aqjh_name,"�������±������ǲ�С�ģ�������˸������ˣ�")
else
dim js(10)
js(0) ="���ѿ�"
js(1) ="��Ѫ��"
js(2) ="���߿�"
js(3) ="�ݺ���"
js(4) ="�崿��"
js(5) ="���˿�"
js(6) ="����"
js(7) ="�Զ���"
js(8) ="������"
js(9)="������"
randomize()
myxy = Int(Rnd*10)
zstemp=add(rs("w5"),js(myxy),1)
conn.execute "update �û� set ����=����+"& tempjs*30 &",w5='"&zstemp&"' where ����='" & aqjh_name &"'"
kl="վ�����±�����["&aqjh_name&"]�ܵ���죬���ڴ��ǰ����±�������"&aqjh_name&"�õ�վ���ͳ���[<b><font color=red>"&js(myxy)&"</font></b>]1�š���"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>��������Ϣ��</b></font>"&kl	
says=replace(says,chr(39),"\'")
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