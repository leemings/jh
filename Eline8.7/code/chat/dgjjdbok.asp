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
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"http://127.0.0.1/chat")=0 and InStr(http,"http://127.0.0.1/chat")=0then
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');}</script>" 
	Response.End 
end if 
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');</script>"
	Response.End
end if
nowinroom=session("nowinroom")
call roompd("�ᱦ���¾Ž�")
to1=trim(request.form("to1"))
dgjjzs=trim(request.form("dgjjzs"))
tl=int(abs(clng(request.form("tl"))))
nl=int(abs(clng(request.form("nl"))))
maxtl=int(abs(clng(request.form("maxtl"))))
maxnl=int(abs(clng(request.form("maxnl"))))
money=int(abs(clng(request.form("money"))))
zsdj=int(abs(clng(request.form("dj"))))
if Instr(at,chr(39))<>0 or Instr(at,chr(34))<>0 then
	Response.Write "<script Language=Javascript>alert('��Ҫ��ʲô����Զ�㣡');location.href = 'javascript:history.go(-1)'</script>"
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
if to1="���" or to1=Application("sjjh_automanname") or to1=sjjh_name then
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻���ԶԴ�ҡ����ѻ򽭺������˷��С�');</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='" & sjjh_name & "'",conn,2,2
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
if rs("�ȼ�")<zsdj then
	Response.Write "<script Language=Javascript>alert('��ʾ����ĵȼ�����["&zsdj&"],������ʹ�ô��У�');</script>"
	myclose()
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<15 then
	s=15-sj
	Response.Write "<script Language=Javascript>alert('���棺���"& s &"���ٷ���,�ɱ����ţ�');location.href = 'javascript:history.go(-1)';</script>"
	myclose()
end if
mynameid=rs("id")
lxjwg1=rs("�书")
lxjgj1=rs("����")
mydj=rs("�ȼ�")
nn=int(mydj/10)+1
tlsx=rs("�ȼ�")*sjjh_tlsx+5260+rs("������")
nlsx=rs("�ȼ�")*sjjh_nlsx+2000+rs("������")
wgsx=rs("�ȼ�")*sjjh_wgsx+3800+rs("�书��")
tlbf=tl/tlsx
nlbf=nl/nlsx
wgbf=lxjwg1/wgsx
bf=(nlbf+tlbf+wgbf)/3
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn,2,2
jhhy=rs("��Ա�ȼ�")
ntnt=rs("�ȼ�")
tjf=rs("ͨ��")
if rs("grade")>=6 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ���Ա��������ʲô����������վ����Ӧ��');</script>"
	myclose()
end if
tomp=rs("����")
lxjwg2=rs("�书")
lxjgj2=rs("����")
yinliang2=rs("����")
youid=rs("id")
select case dgjjzs
	case "�ܾ�ʽ"
		ssl=5000
	case "�Ƶ�ʽ"
		ssl=7000
	case "�ƽ�ʽ"
		ssl=9000
	case "��ǹʽ"
		ssl=11000
	case "�Ʊ�ʽ"
		ssl=13000
	case "����ʽ"
		ssl=15000
	case "�Ƽ�ʽ"
		ssl=17000
	case "����ʽ"
		ssl=19000
	case "����ʽ"
		ssl=21000
end select
randomize timer
sss=int(rnd*99)+1
y=int(rnd*999)+1
qq1=lxjwg1+lxjgj1+ssl-(lxjwg2+lxjgj2)
qq2=(tl+nl)*sss
qq3=(qq1+qq2)*bf
qq4=sqr(qq3)
killer=int(qq4*nn+money/y)
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
shengtili=rs("����")-killer
conn.execute "update �û� set ����=����-"  & killer & " where ����='" & to1 & "'"
conn.execute "update �û� set ����=����-100,����=����-"&money&",����=����-" & tl & ",����=����-" & nl & ",����ʱ��=now() where ����='" & sjjh_name & "'"
e=""
if shengtili<-100 then
	conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1,����=����+" & yinliang2 & " where ����='" & sjjh_name & "'"
	conn.execute "update �û� set ����="&shengtili&",״̬='��',����=0,�¼�ԭ��='"&sjjh_name&"|���¾Ž�֮"&dgjjzs&"' where ����='" & to1 & "'"
	e="�㣬<font color=blue>" & to1 & "</font><bgsound src=wav/daipu.wav loop=1>�����ĵ�����ȥ�����Ӵ˽�����������һֻ��Ϻ," & sjjh_name & "�õ�" & to1 & "�������ֽ�<font color=blue><b>" & yinliang2 & "</b></font>��!"
	call boot(to1,"���¾Ž��������ߣ�"&sjjh_name&",["&mp&"]"&dgjjzs)
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','���¾Ž�֮"&dgjjzs&"','����')"
end if
if e="" then
	e="�㡣"
end if
stunt="<font color=blue>"& sjjh_name & "</font><bgsound src=wav/bsj.wav loop=1>����" & nl & "��������" & tl & "����������<img src='img/021.gif'><font color=blue>" & to1 & "</font>ʹ���˽����Ͼ��Ѵ��Ķ��¾Ž�֮<font color=008000><b>��"&dgjjzs&"��</b></font>��ɱ��" & killer & e
if shengtili<-100 then
	bbb=zxrsbd(sjjh_name,mynameid)
	stunt=stunt&"<BR><BR>"&bbb
	newstunt="Ϊ��ȡ���"&stunt
	call dbxx(newstunt)
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ϲ������"&dgjjzs&"�Ѿ�������ɣ�');</script>"
says="<font color=red><b>�����¾Ž�֮"&dgjjzs&"��</b>"& stunt &"</font>"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho=to1
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
function myclose()
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end function
%>