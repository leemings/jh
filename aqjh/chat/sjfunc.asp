<!--#include file="czdj.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
function bdsays(says)
says=replace(says,"\","")
says=replace(says,"..","")
says=Replace(says,"<"," ")
says=Replace(says,">"," ")
says=Replace(says,"\x3c"," ")
says=Replace(says,"\x3e"," ")
says=Replace(says,"\074"," ")
says=Replace(says,"\74"," ")
says=Replace(says,"\75"," ")
says=Replace(says,"\76"," ")
says=Replace(says,"&lt"," ")
says=Replace(says,"&gt"," ")
says=Replace(says,"\076"," ")
says=Replace(LCase(says),LCase("con"),"f2uck")
says=Replace(LCase(says),LCase("or"),"���")
says=Replace(LCase(says),LCase("java"),"f2uck")
says=Replace(LCase(says),LCase("www"),"f2uck")
says=Replace(LCase(says),LCase("com"),"f2uck")
says=Replace(LCase(says),LCase("net"),"f2uck")
says=Replace(says,"��","f2uck")
says=Replace(says,"��","f2uck")
says=Replace(says,"asp","f2uck")
says=Replace(says,"jh","f2uck")
says=Replace(says,"ȥ�˿��Ե�","f2uck")
says=Replace(says,"ȥ�˿�����","f2uck")
says=Replace(says,"��һ���ý���","f2uck")
says=Replace(says,"ȥ�ҵĽ���","f2uck")
says=Replace(says,"���潭��","f2uck")
says=Replace(says,"���ҽ���","f2uck")
says=Replace(LCase(says),LCase("http"),"f2uck")
maren=instr(says,"f2uck")
	if maren<>0 then
		Response.Write "<script Language=Javascript>top.location.href='../autokick.asp';</script>"
		Response.end
	end if
bdsays=says
end function
'��Ѩ������
sub dianzan(towho)
'��Ѩ
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');}</script>" 
	Response.End 
end if
if towho="" then towho="���"
gw=left(towho,1)
if IsNumeric(gw)=true then
zz=split(gw,"��")
gw=1
else 
gw=0
end if 
if towho<>Application("aqjh_automanname") and towho<>"���" and Instr(Application("aqjh_useronlinename"&nowinroom)," "&towho&" ")=0 and towho<>Application("figo") and gw=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����" & towho & "�������������У����ܶ��䷢�ԡ�');parent.f2.document.af.towho.value='���';parent.f2.document.af.towho.text='���';parent.m.location.reload();</script>"
	Response.end
end if
Application.Lock
if Instr(LCase(application("aqjh_dianxuename")),LCase(aqjh_name))>0 then
dianxue=split(application("aqjh_dianxuename"),";")
for x=0 to UBound(dianxue)-1
	dxname=split(dianxue(x),"|")
if dxname(0)=aqjh_name then
	if DateDiff("s",dxname(1),now())<60 then
		Response.Write "<script Language=Javascript>alert('���ѱ���Ѩ���������κ��°�����ʣ" & 60-DateDiff("s",dxname(1),now()) & "���Զ��⿪!');</script>"
		erase dianxue
		erase dxname
		response.end
	else
		dxcz=dianxue(x)&";"
		Application.Lock()
		application("aqjh_dianxuename")=replace(application("aqjh_dianxuename"),dxcz,"")
		Application.UnLock()
		Response.Write "<script Language=Javascript>alert('ʱ�䵽�ˣ������Ѩ�Զ��⿪��!');</script>"
		erase dianxue
		erase dxname
		response.end
	end if
end if
erase dxname
next
erase dianxue
end if
Application.UnLock
'���뿪
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 and left(says,4)<>"/����$" and aqjh_grade<10 then
	Response.Write "<script Language=Javascript>alert('�����ڡ����롱״̬�У���ʹ�á��һ����������ܽ����');</script>"
	response.end
end if
'�ж϶Է�����
if Instr(LCase(application("aqjh_zanli")),LCase("!"&towho&"!"))>0 and left(says,4)<>"/����$" then
	aqjh_zanli=split(application("aqjh_zanli"),";")
	for x=0 to UBound(aqjh_zanli)
		dxname=split(aqjh_zanli(x),"|")
	if dxname(0)="!"&towho&"!" then
		Response.Write "<script Language=Javascript>alert('"&towho&"˵��"&dxname(1)&"');</script>"
		erase dxname
		erase aqjh_zanli
		response.end
	end if
	erase dxname
	next
	erase aqjh_zanli
end if
end sub
'����ģ��
sub chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
zj="<a href=javascript:parent.sw(\'[" & aqjh_name & "]\'); target=f2><font color="&addwordcolor&">"& aqjh_name & "</font></a>"
br="<a href=javascript:parent.sw(\'[" & towho & "]\'); target=f2><font color="&addwordcolor&">"& towho & "</font></a>"
says=Replace(says,"##",zj)
says=Replace(says,"%%",br)
says=Replace(says,"$$redb","<font color=red><b>")
says=Replace(says,"$$blueb","<font color=blue><b>")
says=Replace(says,"$$yellowb","<font color=yellow><b>")
says=Replace(says,"$$greenb","<font color=green><b>")
says=Replace(says,"$$b","</b></font>")
says=Replace(says,"$$red","<font color=red>")
says=Replace(says,"$$blue","<font color=blue>")
says=Replace(says,"$$yellow","<font color=yellow>")
says=Replace(says,"$$green","<font color=green>")
says=Replace(says,"$$","</font>")
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway & "," & nowinroom & ");<"&"/script>"
AddMsg SayStr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
Response.Write "<script Language=JavaScript>"
Response.Write "parent.f1.scrollWindow();"
Response.Write "</script>"
end sub
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
sub myclose(fs)
select case fs
case 0
	 if IsObject(rs) then
		set rs=nothing
	 end if
	 if IsObject(conn) then
	 	set conn=nothing
	 end if
case 1
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
case 2
	rs.close
	set rs=nothing
case 3
	conn.close
	set conn=nothing
end select
end sub
sub mess(mymess,fs)
Response.Write "<script language=JavaScript>{alert('"&mymess&"');}</script>"
Response.End
end sub
function ErrSays(from1,strContent)
ErrSays="[<a href=javascript:parent.sw('["&from1&"]'); target=f2> "&from1 & " </a>]" & strContent
end function
Function ErrALT2(strContent)
ErrALT2 = "<script Language=Javascript>alert('" & strContent & "');</script>"
End Function
Sub ErrALT(strContent)
Response.Write "<script Language=Javascript>alert('" & strContent & "');window.close();</script>"
Response.end
End Sub
%>