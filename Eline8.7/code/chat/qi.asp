<%@ LANGUAGE=VBScript codepage ="936" %>
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
if Application("sjjh_qi")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����С���ǲ���û������ѽ��������������');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("sjjh_qi"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	response.end
end if
Application.Lock
Application("sjjh_qi")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.execute "update �û� set ����=����-"&tempjs&" where ����='" & sjjh_name &"'"
rs.open "select ����,w9 FROM �û� WHERE ����='" & sjjh_name &"'",conn
if rs("����")<-100 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='��ǹ|"&fn1&"' where ����='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'��ǹ','���´�����ǹ������','����')"
	kl="<img src='img/tank.gif'>̫�����ˣ�["&sjjh_name&"]Ϊ�˱�����ҵİ�ȫ������ȡ�壬����ǹɨ���������䣬�������ˡ�����"
	call boot(sjjh_name,"��ǹ�������ߣ���ǹ������ѽ������")
else
dim js(13)
js(0) ="�ӵ�"
js(1) ="ħ��ˮ��"
js(2) ="������"
js(3) ="����׶"
js(4) ="������"
js(5) ="Ѫ����"
js(6) ="ħ����ʯ"
js(7) ="���յ���"
js(8) ="������"
js(9)="���λ�"
js(10)="������ǹ"
js(11)="��ҩ"
js(12)="��Ӱˮ��"
randomize()
myxy = Int(Rnd*13)
zstemp=add(rs("w9"),js(myxy),1)
conn.execute "update �û� set ����=����+"& tempjs*10 &",w9='"&zstemp&"' where ����='" & sjjh_name & "'"
kl="��call,Ӣ�ۣ�����Ӣ�ۣ�["&sjjh_name&"]����ǹΪ�п�����ǹ������ǹ���ڵв�����ԭ�书�������Լ�Ҳ�˵Ĳ��ᣬ�ٸ�����<img src='img/251.gif'>"&tempjs*15&"�����ڴ����ǹ��"&sjjh_name&"�õ���[<b><font color=red>"&js(myxy)&"</font></b>]1������"

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
