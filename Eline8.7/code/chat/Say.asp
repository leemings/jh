<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sayRandom.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ�����\n     ��ȷ�����أ�');}</script>" 
	Response.End 
end if 
sjjh_sid=trim(request.cookies("yxjh")("sjjh_sid")) 
'if (sjjh_sid="" or  sjjh_sid<>session.sessionid) and Session("sjjh_grade")<6 then
'Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('��һ̨��������˶���ʺţ���ϵͳ�����');}</script>"
'Response.End 
'end if
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then 
	Session.Abandon
	Response.Write "<script language=JavaScript>{alert('��ʾ����ֹʹ�ô���');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
'#####���䴦��#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
useronlinename=Application("sjjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
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
sj=s & ":" & f & ":" & m
if DateDiff("n",Session("sjjh_savetime"),now())>=15 then
	Response.Write "<script Language=Javascript>parent.f3.location.href='savevalue.asp';alert('"&sjjh_name&"����ϲ�����Զ����ɹ�����㲻���룡');</script>"
end if
sj2=n & "-" & y & "-" & r & " " & sj
if DateDiff("s",Session("sjjh_lasttime"),sj2)<2 then
Response.Write "<script language=JavaScript>{alert('�л�����˵����ҭ��Ŷ��');}</script>"
Response.End
end if
t="<font class=t>(" & sj & ")</font>"
'�Ƿ񱻵�Ѩ
Application.Lock
if Instr(LCase(application("sjjh_dianxuename")),LCase(sjjh_name))>0 then
dianxue=split(application("sjjh_dianxuename"),";")
dxdx=0
for x=0 to UBound(dianxue)-1
dxname=split(dianxue(x),"|")
if dxname(0)=sjjh_name then
	if DateDiff("s",dxname(1),sj2)<60 then
		Response.Write "<script Language=Javascript>alert('���ѱ���Ѩ���������κ��°�����ʣ[" & 60-DateDiff("s",dxname(1),sj2) & "��]�Զ��⿪��');</script>"
		erase dianxue
		erase dxname
		response.end
	else
		application("sjjh_dianxuename")=replace(application("sjjh_dianxuename"),dianxue(x)&";","")
		Response.Write "<script Language=Javascript>alert('ʱ�䵽�ˣ������Ѩ�Զ��⿪�ˣ�');</script>"
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
towho=Trim(Request.Form("towho"))
says=Trim(Request.Form("sy"))
if Instr("/����$/����$/Ͷ��$/�¶�$/�ÿ�$",left(says,4))>=0 and sjjh_grade<6 then
 if Instr(LCase(application("qzld_pingbi")),LCase(";"&sjjh_name&"|"&towho&";"))>0 then
   Response.Write "<script Language=javascript>alert('���Ѿ����˼������ˣ�Ҫ���˼�˵�����Ȱ��˼Ҵ�����������ɾ����');</script>"
  response.end
 elseif Instr(LCase(application("qzld_pingbi")),LCase(";"&towho&"|"&sjjh_name&";"))>0 then
   Response.Write "<script Language=javascript>alert('���˵�������Ѿ����������ˣ��㲻���ٶ��˼�˵���ˣ�');</script>"
  response.end
 end if
end if
'�Ƿ�����
'�ж��Լ�˵��
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0 and sjjh_grade<10 then
	Response.Write "<script Language=Javascript>alert('������[����]״̬�У���ʹ��[�һ�����]���ܽ����');</script>"
	response.end
end if
'�ж϶Է�����
if Instr(LCase(application("sjjh_zanli")),LCase("!"&towho&"!"))>0 then
	sjjh_zanli=split(application("sjjh_zanli"),";")
	for x=0 to UBound(sjjh_zanli)
		dxname=split(sjjh_zanli(x),"|")
	if dxname(0)="!"&towho&"!" then
		Response.Write "<script Language=Javascript>alert('"&towho&"˵��"&dxname(1)&"');</script>"
		erase dxname
		erase sjjh_zanli
		response.end
	end if
	erase dxname
	next
	erase sjjh_zanli
end if
if towho="" then towho="���"
if len(towho)>10 then towho=Left(towho,10)
if towho<>Application("sjjh_automanname") and towho<>"���" and Instr(Application("sjjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
Response.Write "<script Language=Javascript>alert('[" & towho & "]�����������У����ܶ��䷢�ԣ�');parent.f2.document.af.towho.value='���';parent.f2.document.af.towho.text='���';parent.m.location.reload();</script>"
Response.end
end if
act=0
towhoway=Request.Form("towhoway")
if towhoway<>1 then towhoway=0
addwordcolor=Request.Form("addwordcolor")
saycolor=Request.Form("saycolor")
addsays=Request.Form("addsays")

'�������з�����ĵȼ�72����
if towho="���" or towho=sjjh_name then towhoway=0
if instr(Application("sjjh_admin"),towho)<>0 or towho=Application("sjjh_automanname") or towho=Application("figo") then towhoway=1
if len(says)>150 then says=Left(says,150)
if (says="" or says="//") then Response.End
if InStr(says,"��")<>0 or InStr(says,"&#63736")<>0 then Response.End
says=Replace(says,"&amp;","")
says=lcase(says)
says=Replace(says,"&amp;","")
'˽�ĵȼ�Ĭ��Ϊ7
says=lcase(says)
if sjjh_grade<9 then
says=replace(says,"..","")
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
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
says=Replace(LCase(says),LCase(""),"f2uck")
says=Replace(LCase(says),LCase("or"),"���")
says=Replace(LCase(says),LCase(""),"f2uck")
says=Replace(LCase(says),LCase(""),"f2uck")
says=Replace(LCase(says),LCase(""),"f2uck")
says=Replace(LCase(says),LCase(""),"f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(says,"","f2uck")
says=Replace(LCase(says),LCase(""),"f2uck")
maren=instr(says,"f2uck")
if maren<>0 then
Response.Write "<script Language=Javascript>location.href='autokick.asp';</script>"
Response.end
end if
end if

if sjjh_jhdj<10 then
if InStr(says,"q")<>0 then Response.End
if InStr(says,"Q")<>0 then Response.End
if InStr(says,"��")<>0 then Response.End
if InStr(says,"����")<>0 then Response.End
end if

'�Ƿ�ս���ȼ�����5�ſ�����ͼ
if sjjh_jhdj>2 and Instr(says," ")=0 and Instr(says,"[img]")<>0 then
if Instr(says,"[img]")<>0 then
if instr(says,"width")<>0 or instr(says,"height")<>0 then Response.End	
gsz=0
Do While Instr(says,"[img]")<>0
says=Replace(says,"[img]","<img src=pic/",1,1)
gsz=gsz+1
Loop
gsy=0
Do While Instr(says,"[/img]")<>0
says=Replace(says,"[/img]",">",1,1)
gsy=gsy+1
Loop
if gsz>gsy then says=says&">"
end if
end if
zj="<a href=javascript:parent.sw('[" & sjjh_name & "]'); target=f2><font color="&addwordcolor&">"& sjjh_name & "</font></a>"
br="<a href=javascript:parent.sw('[" & towho & "]'); target=f2><font color="&addwordcolor&">"& towho & "</font></a>"
if Left(says,2)="//" then
	says=Replace(says,"##",zj,1,3,1)
	says=Replace(says,"%%",br,1,3,1)
	says="<font color=" & sayscolor & ">" & zj & Trim(mid(says,3,len(says)-2)) & "</font>"
	act=1
end if
'ɾ����ԭ���ı���
addsays=addsays
Session("sjjh_lasttime")=sj2
says=Replace(says,"\","\\")
says=Replace(says,"/","\/")
says=Replace(says,chr(34),"\"&chr(34))
act="����"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
says=says&t
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway & "," & nowinroom & ");<"&"/script>"
AddMsg SayStr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next

session("SayCount")=Application("SayCount")
Response.Write "<script Language=JavaScript>"
Response.Write "parent.f1.scrollWindow();"
Response.Write "</script>"

RandomSay

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
