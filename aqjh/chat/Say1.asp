<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="npc_chat.asp"-->
<!--#include file="../const2.asp"-->
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
ysjh_name=Session("ysjh_name")
ysjh_grade=Session("ysjh_grade")
ysjh_jhdj=Session("ysjh_jhdj")
nowinroom=session("nowinroom")
ysjh_roominfo=split(Application("ysjh_room"),";")
chatinfo=split(ysjh_roominfo(nowinroom),"|")
if ysjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');}</script>" 
	Response.End 
end if
if Application("ysjh_yjdh")=1 then
ysjh_sid=trim(request.cookies("ysjh")("ysjh_sid")) 
if (ysjh_sid="" or  ysjh_sid<>session.sessionid) and Session("ysjh_grade")<6 then
Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('��һ̨��������˶���ʺţ���ϵͳ�����');}</script>"
Response.End 
end if
end if
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("ysjh_disproxy")=1 then
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then 
	Session.Abandon
	Response.Write "<script language=JavaScript>{alert('��ʾ����ֹʹ�ô���');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
end if
'#####���䴦��#####
ysjh_roominfo=split(Application("ysjh_room"),";")
chatinfo=split(ysjh_roominfo(nowinroom),"|")
useronlinename=Application("ysjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&ysjh_name&" "))=0 then 
	Session("ysjh_inthechat")="0" 
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
if DateDiff("n",Session("ysjh_savetime"),now())>=15 then
	Response.Write "<script Language=Javascript>parent.f3.location.href='savevalue.asp';</script>"
end if
sj2=n & "-" & y & "-" & r & " " & sj
if DateDiff("s",Session("ysjh_lasttime"),sj2)<2 then
Response.Write "<script language=JavaScript>{alert('�л�����˵����ҭ��Ŷ��');}</script>"
Response.End
end if
t="<font class=t>(" & sj & ")</font>"
'�Ƿ񱻵�Ѩ
Application.Lock
if Instr(LCase(application("ysjh_dianxuename")),LCase(ysjh_name))>0 then
dianxue=split(application("ysjh_dianxuename"),";")
dxdx=0
for x=0 to UBound(dianxue)-1
dxname=split(dianxue(x),"|")
if dxname(0)=ysjh_name then
	if DateDiff("s",dxname(1),sj2)<60 then
		Response.Write "<script Language=Javascript>alert('���ѱ���Ѩ���������κ��°�����ʣ" & 60-DateDiff("s",dxname(1),sj2) & "���Զ��⿪!');</script>"
		erase dianxue
		erase dxname
		response.end
	else
		application("ysjh_dianxuename")=replace(application("ysjh_dianxuename"),dianxue(x)&";","")
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
towho=Trim(Request.Form("towho"))
says=Trim(Request.Form("sy"))
if Instr("/����$/����$/Ͷ��$/�¶�$/�ÿ�$",left(says,4))=0 and ysjh_grade<6 then
 if Instr(LCase(application("qzld_pingbi")),LCase(";"&ysjh_name&"|"&towho&";"))>0 then
   Response.Write "<script Language=javascript>alert('���Ѿ����˼������ˣ�Ҫ���˼�˵�����Ȱ��˼Ҵ�����������ɾ����');</script>"
  response.end
 elseif Instr(LCase(application("qzld_pingbi")),LCase(";"&towho&"|"&ysjh_name&";"))>0 then
   Response.Write "<script Language=javascript>alert('���˵�������Ѿ����������ˣ��㲻���ٶ��˼�˵���ˣ�');</script>"
  response.end
 end if
end if
'�Ƿ�����
'�ж��Լ�˵��
if Instr(LCase(application("ysjh_zanli")),LCase("!"&ysjh_name&"!"))>0 and ysjh_grade<9 then
	Response.Write "<script Language=Javascript>alert('�����ڡ����롱״̬�У���ʹ�á��һ����������ܽ����');</script>"
	response.end
end if
'�ж϶Է�����
if Instr(LCase(application("ysjh_zanli")),LCase("!"&towho&"!"))>0 and left(says,4)<>"/�ÿ�$" and ysjh_grade<9 then
	ysjh_zanli=split(application("ysjh_zanli"),";")
	for x=0 to UBound(ysjh_zanli)
		dxname=split(ysjh_zanli(x),"|")
	if dxname(0)="!"&towho&"!" then
		Response.Write "<script Language=Javascript>alert('"&towho&"˵��"&dxname(1)&"');</script>"
		erase dxname
		erase ysjh_zanli
		response.end
	end if
	erase dxname
	next
	erase ysjh_zanli
end if
if towho="" then towho="���"
gw=left(towho,1)
if IsNumeric(gw)=true then
zz=split(gw,"��")
gw=1
else 
gw=0
end if  

if len(towho)>10 then towho=Left(towho,10)
if Not(towho=Application("ysjh_automanname") or towho=Application("figo") or towho="���" or Instr(useronlinename," "&towho&" ")<>0) and gw=0 then
Response.Write "<script Language=Javascript>alert('��" & towho & "�������������У����ܶ��䷢�ԡ�');parent.f2.document.af.towho.value='���';parent.f2.document.af.towho.text='���';parent.m.location.reload();</script>"
Response.end
end if
act=0
towhoway=Request.Form("towhoway")
if towhoway<>1 then towhoway=0
addwordcolor=Request.Form("addwordcolor")
saycolor=Request.Form("saycolor")
addsays=Request.Form("addsays")
'�������з�����ĵȼ�72����
if towho="���" or towho=ysjh_name then towhoway=0
if instr("," & Application("ysjh_admin") & ",","," & towho & ",")<>0 or towho=Application("ysjh_automanname") or towho=Application("figo") then towhoway=1
if len(says)>150 then says=Left(says,150)
if (says="" or says="//") then Response.End
if InStr(says,"��")<>0 or InStr(says,"&#63736")<>0 then Response.End
says=Replace(says,"&amp;","&")
says=lcase(says)
says=Replace(says,"&amp;","&")
'˽�ĵȼ�Ĭ��Ϊ7
says=lcase(says)
says=replace(says,"\","")
if ysjh_grade<9 then
says=replace(says,"&#","")
says=replace(says,"..","")
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
says=replace(says,chr(32),"")
says=replace(says,"��","")
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
says=Replace(LCase(says),LCase("org"),"f2uck")
says=Replace(LCase(says),LCase("com"),"f2uck")
says=Replace(LCase(says),LCase("net"),"f2uck")
says=Replace(says,"QQ","f2uck")
says=Replace(says,"pp","f2uck")
says=Replace(says,"200.","f2uck")
says=Replace(says,"�ףף�","f2uck")
says=Replace(says,"������","f2uck")
says=Replace(says,"asp","f2uck")
says=Replace(says,"jh","f2uck")
says=Replace(says,"ȥ�˿��Ե�","f2uck")
says=Replace(says,"ȥ�˿�����","f2uck")
says=Replace(says,"��һ���ý���","f2uck")
says=Replace(says,"������Ѷ","f2uck")
says=Replace(says,"����qq����ô�","f2uck")
says=Replace(says,"����qq����ô�","f2uck")
says=Replace(says,"����q����ô�","f2uck")
says=Replace(says,"����q����ô�","f2uck")
says=Replace(says,"����","f2uck")
says=Replace(says,"����","f2uck")
says=Replace(says,"����վ��","f2uck")
says=Replace(LCase(says),LCase("http"),"f2uck")
maren=instr(says,"f2uck")
if maren<>0 then
Response.Write "<script Language=Javascript>location.href='autokick.asp';</script>"
Response.end
end if
end if
if ysjh_jhdj<10 then
if InStr(says,"q")<>0 then Response.End
if InStr(says,"Q")<>0 then Response.End
if InStr(says,"��")<>0 then Response.End
if InStr(says,"����")<>0 then Response.End
end if
'������
Zshow=request("Zshow")
if Zshow="1" then
show=ysjh_Zshow
says=Plus_ChkPicWords(says)
Function Plus_ChkPicWords(strText)
	If IsNull(StrText) Then Exit Function
        StrText=Ucase(StrText)
        StrText=replace(StrText,chr(32),"")
        if left(StrText,5)="[IMG]" then
           Response.Write "<script Language=Javascript>alert('ʹ����ͼ���ȹر������㹦�ܣ�');parent.f2.document.af.sytemp.value='';</script>"
           Response.end
        end if
        if len(StrText)>35 then
           Response.Write "<script Language=Javascript>alert('Ϊ�˽��������ٶȣ������㲻�ó���50����');parent.f2.document.af.sytemp.value='';</script>"
           Response.end
        end if
	Dim PicWords,i
	PicWords = Split(show,"|")
	For i = 0 To Ubound(PicWords)
	   StrText = Replace(trim(StrText),PicWords(i),"[img]../picwords/"&picwords(i)&".gif[/img]")
	Next
	Plus_ChkPicWords = StrText
End Function
end if
'������
'�Ƿ�ս���ȼ�����18�ſ�����ͼ
v0=says

If InStr(v0, "[flash]") <> 0 Then
if ysjh_jhdj<1 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ�ȼ�1�����ϣ�');}</script>"
	Response.End
end if
if Instr(LCase(Application("ysjh_useronlinename"&nowinroom)),LCase(" "&ysjh_name&" "))=0 then 
	Session("ysjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
        gsz = 0
        Do While InStr(v0, "[flash]") <> 0
			v0 = Replace(v0, "[flash]", "<object codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' height='250' width='250' onMouseOver='Play();' classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'><param name='movie' value='CW_swf/", 1, 1)
			gsz = gsz + 1
        Loop
        gsy = 0
        Do While InStr(v0, "[/flash]") <> 0
		v0 = Replace(v0, "[/flash]", "'><param name='quality' value='high'><param name='wmode' value='transparent'></object>", 1, 1)
		gsy = gsy + 1
        Loop
        If gsz > gsy Then v0 = v0 & ">"
End If
says=v0
dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
're.Pattern="\[afaimg@*([0-9]*)\]"
'says=re.Replace(says,"&nbsp;<img src=QQ/gif/$1.gif border=0 alt='������ҳ������' onload=ShowQQ('$1');  onmouseover=ShowQQ('$1'); style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;'>")
re.Pattern="\[afaimg@*([0-9]*)\]"
says=re.Replace(says,"&nbsp;<img src=QQ/gif/$1.gif border=0 alt='������ҳ������' onload=ShowQQ('$1');  onmouseover=ShowQQ('$1'); style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;'>")
'says=re.Replace(says,"&nbsp;<img src=QQ/gif/$1.gif border=0 alt=������ҳ������ onload=setTimeout(""ShowQQ('$1')"";,3000);  onmouseover=ShowQQ('$1'); style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;'>")
'v1=says
'If InStr(v1, "[afaimg@]") <> 0 Then

 '       gsz1 = 0
  '      Do While InStr(v1, "[afaimg@]") <> 0
			'v1 = Replace(v1, "[afaimg@]", "<img src=QQ/gif/", 1, 1)
			'gsz1 = gsz1 + 1
        'Loop
        'gsy1 = 0
        'Do While InStr(v1, "[/qq]") <> 0
		'v1 = Replace(v1, "[/qq]", ".gif border=0 alt=���ҳ������ style='cursor: hand;border:1px border-collapse: collapse; border-style: dotted;' onmouseover=ShowQQ('001');>&nbsp;", 1, 1)
		'gsy1 = gsy1 + 1
        'Loop
        'If gsz1 > gsy1 Then v1 = v1 & ">"
'End If
'says=v1
if ysjh_jhdj>=18 and Instr(says," ")=0 and Instr(says,"[img]")<>0 then
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
zj="<a href=javascript:parent.sw('[" & ysjh_name & "]'); target=f2><font color="&addwordcolor&">"& ysjh_name & "</font></a>"
br="<a href=javascript:parent.sw('[" & towho & "]'); target=f2><font color="&addwordcolor&">"& towho & "</font></a>"
if Left(says,2)="//" then
	says=Replace(says,"##",zj,1,3,1)
	says=Replace(says,"%%",br,1,3,1)
	says="<font color=" & sayscolor & ">" & zj & Trim(mid(says,3,len(says)-2)) & "</font>"
	act=1
end if
'ɾ����ԭ���ı���
addsays=addsays
Session("ysjh_lasttime")=sj2
says=Replace(says,"\","\\")
says=Replace(says,"/","\/")
says=Replace(says,chr(34),"\"&chr(34))
if chatinfo(6)=0 then
'��ʼ����¼�
randomize()
rnd1=int(rnd*400)+1
sayyq=""
'ÿСʱ15�ֿ��Է����
if Minute(time())=1 then
if Application("ysjh_zzfafang")<>""  then
      zzfff=Application("ysjh_zzfafang")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
      sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>�����鷢�Ž�ҡ�ÿСʱ��һ���ӷ���5����ң�����Ҫ�������ã�ʱЧ60�롣</font><input  type=button value='��ȡ��ң�' onClick=javascript:zsdzff"&rnd1&".disabled=1;window.open('sjfunc/zz_fafang.asp','d') name=zsdzff"&rnd1&">"
      Application.Lock
      Application("ysjh_zzfafang")="5|"&now()
      application("ysjh_zzffjl")=""
      Application.UnLock
 end if
end if
if Minute(time())=15  then
if Application("ysjh_zzfafang1")<>""  then
      zzfff=Application("ysjh_zzfafang1")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
   sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>�����鷢���书��ÿСʱ��15���ӷ����书������Ҫ�������ã�ʱЧ60�룬����ȡ�������Զ�������</font><input  type=button value='��ȡ�书��' onClick=javascript:zsdzff"&rnd1&".disabled=1;window.open('sjfunc/zdff1.asp','d') name=zsdzff"&rnd1&">"
      Application.Lock
      Application("ysjh_zzfafang1")="1|"&now()
      application("ysjh_zzffjl1")=""
      Application.UnLock
 end if
end if

if Minute(time())=30  then
if Application("ysjh_zzfafang2")<>""  then
      zzfff=Application("ysjh_zzfafang2")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
   sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>�����鷢�Ŵ�㡿ÿСʱ��30���ӷ��Ŵ�㣬����Ҫ�������ã�ʱЧ60�룬����ȡ�������Զ�������</font><input  type=button value='��ȡ��㣡' onClick=javascript:zsdzff"&rnd1&".disabled=1;window.open('sjfunc/zdff2.asp','d') name=zsdzff"&rnd1&">"
      Application.Lock
      Application("ysjh_zzfafang2")="1|"&now()
      application("ysjh_zzffjl2")=""
      Application.UnLock
 end if
end if
if Minute(time())=45  then
if Application("ysjh_zzfafang3")<>""  then
      zzfff=Application("ysjh_zzfafang3")
       fffdata=split(zzfff,"|")
        ffsl=abs(int(clng(fffdata(0))))
        ffsj=fffdata(1)
       erase fffdata
    end if 
 if DateDiff("s",ffsj,now())>10 then
   sayyq="<bgsound src=wav/luo.wav loop=2><br><font color=red>�����鷢�����ӡ�ÿСʱ��45���ӷ�������������Ҫ�������ã�ʱЧ60�룬����ȡ�������Զ�������</font><input  type=button value='���ҷ��ƣ�' onClick=javascript:zsdzff"&rnd3&".disabled=1;window.open('sjfunc/zdff3.asp','d') name=zsdzff"&rnd3&">"
      Application.Lock
      Application("ysjh_zzfafang3")="1|"&now()
      application("ysjh_zzffjl3")=""
      Application.UnLock
 end if
end if
'��������������
if rnd1<=16 then
	chance=array("����","����","����","����","�书","����","����","����","����","����","����","����","�书","����","����","����")
	chancetxtgood=array("�����ִ�Ħ��ʦ��̳������##��������","�䵱���������������Ღ�����##��������","��Ѫ���������ػ��˼䣬ͻ������##��������","����������϶����䣬�����������##��������","##�����˾ƣ�ȴ����������͵����Ʒ����ʹ���������࣬�书����","##���ڽ������ߣ�����MM��ȴ��������Ӳ��������","##���ڻ�ɽ������򾣬����Ҳ�벻���������˵�еĽ��������������˶�����","�ڷ粨ͤ����������Ĺ��##��������","�����ִ�Ħ��ʦ��̳������##��������","�䵱���������������Ღ�����##��������","��Ѫ���������ػ��˼䣬ͻ������##��������","����������϶����䣬�����������##��������","##�����˾ƣ�ȴ����������͵����Ʒ����ʹ���������࣬�书����","##���ڽ������ߣ�����MM��ȴ��������Ӳ��������","##���ڻ�ɽ������򾣬����Ҳ�벻���������˵�еĽ��������������˶�����","�ڷ粨ͤ����������Ĺ��##��������")
	chancetxtbad=array("##��ʱ�䰾ս����,���¾������ã���·Ҳˤ��,������˶��½���","##��·�ϼ�����һ��������С���������֮�ģ�����ȴ������͵͵�ػ�ȥ������","##�չؿ�����ʬ������ȴ����Ҫ�죬����ֻ�˶�����","��Ц##û����ѧѧ���������Բн��������ֻ��������˰̣�������˶�����","##�������¹��˹�仨¥����������֪���书�½���","##���Ѳ���,������ƭȥ��","##�����䳡��ʶ��Ѧ󴣬��̸�����������˶�������","##;���տն������İ�ʦѧ�գ����Ϲ���û��ѧ��������ȴ�½���","##��ʱ�䰾ս����,���¾������ã���·Ҳˤ��,������˶��½���","##��·�ϼ�����һ��������С���������֮�ģ�����ȴ������͵͵�ػ�ȥ������","##�չؿ�����ʬ������ȴ����Ҫ�죬����ֻ�˶�����","��Ц##û����ѧѧ���������Բн��������ֻ��������˰̣�������˶�����","##�������¹��˹�仨¥����������֪���书�½���","##���Ѳ���,������ƭȥ��","##�����䳡��ʶ��Ѧ󴣬��̸�����������˶�������","##;���տն������İ�ʦѧ�գ����Ϲ���û��ѧ��������ȴ�½���")
	rnd2=int(rnd*400)-100
	if rnd2>0 then
		rndtxt=chancetxtgood(rnd1)&"<font color=#FF8800><b>+"& abs(rnd2) & "</b></font>"
	else
		rnd2=rnd2-1
		rndtxt=chancetxtbad(rnd1)&"<font color=#FF8800><b>-"&abs(rnd2)&"</b></font>"
	end if	
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("ysjh_usermdb")
	conn.execute "update �û� set "&chance(rnd1)&"="&chance(rnd1)&"+"&rnd2&" where ����='" & ysjh_name &"'"
	conn.close
	set conn=nothing
	zj="<font color="&addwordcolor&">"& ysjh_name & "</font>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=FF0000>�����š�"&Replace(rndtxt,"##",zj,1,3,1)&"</font><font class=t>("&time&")</font>"
else
select case rnd1
case 18
banner="��л���֧�֡���ʫ��������.�λ����罫�������׽����ϸ����ĵ�����������.������κ��������ʱ�����.�������ĵ�ַ�� cys.hxjhw.com �����˽�������һ����Ŷ^_^; һ���ƽ�Ľ���,�����ǵĽ������Բ����ҵ�״̬,Ϊ��Ҵ���һ�����õ����л���; ��վ�ٸ���ְ��������в�����ĵط�����ȥ��̳վ�����Ͷ��򷢱������"
banners=Split(Trim(banner),";",-1)
total=UBound(banners)
randomize timer
x=int(rnd*(total+1))
sayyq="<bgsound src=wav/luo.wav loop=3><table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chat/f2/img/rightct3.gif></td><td background=../chat/f2/img/rightct4.gif></td><td><img  src=../chat/f2/img/rightct1.gif></td></tr><tr><td background=../chat/f2/img/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr><td valign=center align=center><font style='font-size:9pt;color:red'>========�ٸ�����========</font></td></tr><tr><td valign=center align=center><font style='font-size:9pt;color:green'>"&banners(x)&"</td></tr></table></td><td background=../chat/f2/img/rightct08.gif></td></tr><tr><td><img src=../chat/f2/img/rightct9.gif></td><td background=../chat/f2/img/rightct10.gif></td><td><img src=../chat/f2/img/rightct11.gif></td></tr></table></td></table>"
case 20,25
	useronlinename=Application("ysjh_useronlinename"&nowinroom)
	online=Split(Trim(useronlinename)," ",-1)
	x=UBound(online)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("ysjh_usermdb")
        randomize timer
        rnd22=int(rnd*4)+1
        select case rnd22
        case 1
			randomize timer
			s=int(rnd*10000)
			for i=0 to x
				conn.Execute "update �û� set ����=����+" & s & " where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>��ϵͳ��Ǯ��</font><B><font color=#ff0000>��Ǯඣ���Ǯඣ�<img src=img/251.gif><img src=img/251.gif><img src=img/251.gif> ���������ÿ���˷���"& s &"������������ַ��http://cys.hxjhw.com�����˽�������һ����Ŷ ^_^</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/47.gif><img src=img/47.gif><img src=img/47.gif><img src=img/47.gif></marquee>"
        case 2
			for i=0 to x
				conn.Execute "update �û� set ���=���+5 where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>��ϵͳ��ҡ�</font><B><font color=#ff0000>վ���λ������л���һֱ�����ԡ���ʫ����������֧�ֺͺ�,���и��ˣ���Ϊ�����ҵ�ÿ���˷��˽��5����</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif></marquee>"
        case 3
			mm=50
			for i=0 to x
				conn.Execute "update �û� set allvalue=allvalue+" & mm & " where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>��ϵͳ���顿</font><font color=#ff0000>վ��·���˵أ����̿�����λ��СϺ��˿�������Ϊ�����ҵ�ÿ���˷��ž���" & mm & "�㣡</font>"
        case 4
			randomize timer
			s=int(rnd*150)
			for i=0 to x
				conn.Execute "update �û� set �ݶ�����=�ݶ�����+" & s & " where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/zt.mid loop=1><font color=green>�����ӡ�</font><font color=#ff0000>��֪˭�Ĵ���©�ˣ���������һ�أ������ĸ�λÿ�˼���" & s & "�㣡</font>"
        end select
case 5
Application("yx8_mhjh_klt1")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ��¹��͵Ϯ<font color=#0000FF>" & ysjh_name & "</font>,ȴ��<font color=#0000FF>" & ysjh_name & "</font>��ȡ����"&Application("yx8_mhjh_klt1")*5000&"</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt1")*5000&" where ����='" & ysjh_name & "'")

case 6
Application("yx8_mhjh_klt3")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ�Ӵ�Ŀֲ����޴���������" & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt3")*2000&"��</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt3")*2000&" where ����='" & ysjh_name & "'")
case 7
Application("yx8_mhjh_klt4")=r
	sayyq="<font color=red>��������</font><font color=red>������ͻȻһ���𵴣�һֻ�޴��ʬ�����Ƕ�����<font color=#0000FF>" & ysjh_name & "</font>��ȡ������"&Application("yx8_mhjh_klt4")*1000&"��</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt4")*1000&" where ����='" & ysjh_name & "'")
case 8
Application("yx8_mhjh_klt4")=r
	sayyq="<font color=red>��������</font><font color=red>һ����͵ķ��������һֻ����ǧ��������Ȼ��������<font color=#0000FF>" & ysjh_name & "</font>���ӻ���"&Application("yx8_mhjh_klt4")*20&"��</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set allvalue=allvalue+"&Application("yx8_mhjh_klt4")*20&" where ����='" & ysjh_name & "'")
case 9
Application("yx8_mhjh_klt6")=r
	sayyq="<font color=red>��������</font><font color=red>һ�������ĺ����ɨ�������������Ⱥ�����Ű������<font color=#0000FF>" & ysjh_name & "</font>�������õ�"&Application("yx8_mhjh_klt6")*100&"�������</font>"
	Set conn=server.CreateObject("adodb.connection")
	conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt6")*100&" where ����='" & ysjh_name & "'")
case 10
Application("yx8_mhjh_klt6")=r
	sayyq="<font color=red>��������</font><font color=red>һ������ͻϮ��������������һȺ��Ů�ӵ��У����˾�ҧ��<font color=#0000FF>" & ysjh_name & "</font>�������õ�����"&Application("yx8_mhjh_klt6")*2000&"�㣡</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt6")*2000&" where ����='" & ysjh_name & "'")
case 11
Application("yx8_mhjh_klt7")=r
	sayyq="<font color=red>��������</font><font color=red>��¿�����ɽ��һ��Ұ��������ɽ���ε���"&Application("ysjh_chatroomname")&"��������ʳ��<font color=#0000FF>" & ysjh_name & "</font>�������õ�����"&Application("yx8_mhjh_klt7")*500&"�㣡</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt7")*500&" where ����='" & ysjh_name & "'")
case 12
Application("yx8_mhjh_klt8")=r
	sayyq="<font color=red>��������</font><font color=red>ħ�ý����ڸ���ɽ������������һ�ɽ��������Լ�<font color=#0000FF>" & ysjh_name & "</font>�������õ�����"&Application("yx8_mhjh_klt8")*100&"�㣡</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt8")*100&" where ����='" & ysjh_name & "'")
case 13
Application("yx8_mhjh_klt9")=r
	sayyq="<font color=red>��������</font><font color=red>��~~~�ÿ��һֻѼ!�������Ӵ��ˣ�ʲô���У�һֻ����Ѽ�ܽ�"&Application("ysjh_chatroomname")&"��<font color=#0000FF>" & ysjh_name & "</font>͵��������"&Application("yx8_mhjh_klt9")*5000&"�㣡</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt9")*5000&" where ����='" & ysjh_name & "'")
case 14
Application("yx8_mhjh_klt10")=r
	sayyq="<font color=red>��������</font><font color=red>������ͻȻ���"&Application("ysjh_chatroomname")&"��һ��������<font color=#0000FF>" & ysjh_name & "</font>��û���ں����ģ��õ�����"&Application("yx8_mhjh_klt10")*10&"�㣡</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set allvalue=allvalue+"&Application("yx8_mhjh_klt10")*10&" where ����='" & ysjh_name & "'")
case 15
Application("yx8_mhjh_klt11")=r
	sayyq="<font color=red>��������</font><font color=red>һֻ���겻���Ľ���֮���ξ���ͻȻ�������������ڣ��͸�<font color=#0000FF>" & ysjh_name & "</font>����"&Application("yx8_mhjh_klt11")*4000&"�㣡</font><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt11")*4000&" where ����='" & ysjh_name & "'")
case 16
Application("yx8_mhjh_klt12")=r
	sayyq="<font color=red>��������</font><font color=red>һֻ��ʷ��ķ�ı���������"&Application("ysjh_chatroomname")&"����<font color=#0000FF>" & ysjh_name & "</font>��ȡ������"&Application("yx8_mhjh_klt12")*1000&"�㣡</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt12")*1000&" where ����='" & ysjh_name & "'")
case 17
Application("yx8_mhjh_klt13")=r
	sayyq="<font color=red>����긽�塿</font><font color=red>����<font color=#0000FF>" & ysjh_name & "</font>�����ڹ�ʱ��С���߻���ħ�������򱦱��Ļ��鸽�ŵ����ϣ���������"&Application("yx8_mhjh_klt13")*4000&"�㣡</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt13")*4000&" where ����='" & ysjh_name & "'")
case 18
Application("yx8_mhjh_klt14")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ���ﴳ��������!" & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt14")*3000&"�㣡</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt14")*3000&" where ����='" & ysjh_name & "'")
case 19
Application("yx8_mhjh_klt15")=r
	sayyq="<font color=red>��������</font><font color=red>��Ư����~~~~һ����֪��ʲô�������������������,�˷��ж�ʧ�˶������,��" & ysjh_name & "�õ���</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ���=���+2 where ����='" & ysjh_name & "'")
case 20
Application("yx8_mhjh_klt16")=r
	sayyq="<font color=red>��������</font><font color=red>��Ư������һ����֪��ʲô�������������������,�˷��ж�ʧ��һ�����,��" & ysjh_name & "�õ���</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ���=���+1 where ����='" & ysjh_name & "'")
case 21
Application("yx8_mhjh_klt17")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt17")*500&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt17")*500&" where ����='" & ysjh_name & "'")
case 22
Application("yx8_mhjh_klt18")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt18")*10&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt18")*10&" where ����='" & ysjh_name & "'")
case 23
Application("yx8_mhjh_klt19")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt19")*10&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt19")*10&" where ����='" & ysjh_name & "'")
case 24
Application("yx8_mhjh_klt20")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt20")*1200&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt20")*1200&" where ����='" & ysjh_name & "'")
case 25
Application("yx8_mhjh_klt21")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt21")*5&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set allvalue=allvalue+"&Application("yx8_mhjh_klt21")*5&" where ����='" & ysjh_name & "'")
case 26
Application("yx8_mhjh_klt22")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt22")*500&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt22")*500&" where ����='" & ysjh_name & "'")
case 27
Application("yx8_mhjh_klt23")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt23")*1900&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt23")*1900&" where ����='" & ysjh_name & "'")
case 28
Application("yx8_mhjh_klt24")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt24")*2300&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt24")*2300&" where ����='" & ysjh_name & "'")
case 29
Application("yx8_mhjh_klt25")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt25")*2000&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt25")*2000&" where ����='" & ysjh_name & "'")
case 30
Application("yx8_mhjh_klt26")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt26")*1400&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt26")*1400&" where ����='" & ysjh_name & "'")
case 31
Application("yx8_mhjh_klt27")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt27")*1100&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt27")*1100&" where ����='" & ysjh_name & "'")
case 32
Application("yx8_mhjh_klt28")=r
	sayyq="<font color=red>��������</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������," & ysjh_name & "�������õ�����"&Application("yx8_mhjh_klt28")*3000&"��</font>"
Set conn=server.CreateObject("adodb.connection")
conn.open Application("ysjh_usermdb")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt28")*3000&" where ����='" & ysjh_name & "'")
case 33,36
	s=int(rnd*20000)
	Application.Lock
	Application("ysjh_blh")=s+10000
	Application.UnLock
	abc="<a href='fafang/blh.asp?tl="&Application("ysjh_blh")&"' target='d'><img src='img/blh.gif' border='0'></a>"
	sayyq="<bgsound src=wav/Gui.WAV loop=1>����Ϣ��ƽ����ͻȻ����һ���ȷ磬һֻ<font color=red>���ϻ�</font>������������Ĵ����ˡ��ٸ�����������������ϻ������ؽ������ϻ�������<font color=red>"&s+10000&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 35,39
	s=int(rnd*20000)
	Application.Lock
	Application("ysjh_jx")=s+17000
	Application.UnLock
	abc="<a href='fafang/jx.asp?tl="&Application("ysjh_jx")&"' target='d'><img src='img/jx.gif' border='0'></a>"
	sayyq="<bgsound src=wav/Gui.WAV loop=1>����Ϣ��һֻ�����Ы�����ɾ�������ɽ�����ˡ���ҿ��ѽ......��Ы������"&s+17000&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 41,43
	s=int(rnd*20000)
	Application.Lock
	Application("ysjh_hx")=s+10000
	Application.UnLock
	abc="<a href='fafang/hx.asp?tl="&Application("ysjh_hx")&"' target='d'><img src='img/hx.gif' border='0'></a>"
	sayyq="<bgsound src=wav/Gui.WAV loop=1>����Ϣ������԰����Ա���ְ�أ����˹���Ȧ��һֻ<font color=red>����</font>�ܳ�����Ȧ����ҿ�ץѽ��������������</font>"&s+10000&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 48,55
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=jstl
		Application.UnLock
		abc="<a href='fafang/kl.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/gui.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/Gui.WAV loop=1><font color=red>����Ϣ��ͻȻ������³¡�������ʬѽ����һŮ�Ӽ�У�һͷ��ʬ���������ң���ʬ������+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 60,66,70
		Application.Lock
		Application("ysjh_dalie")="�ϻ�"
		Application.UnLock
		sayyq="<bgsound src=wav/hu.WAV loop=1><font color=red>����Ϣ��</font><font color=red>ͻȻһֻҰ��<img src=img/laohu.gif>���뽭�������ˣ�������ǿ�ȥ���԰���</font>"
case 90,91,98,120
		jstl=int(rnd*50000)+100
		Application.Lock
		Application("ysjh_yb")=jstl
		Application.UnLock
		abc="<a href='fafang/yb.asp?tl="&Application("ysjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
		sayyq="<font color=red>����Ϣ��������ʲô���ӣ�ͻȻ������ϵ�����һ����Ԫ������ֵ��+"&jstl&"��!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee><bgsound src=wav/diaoxia.wav loop=1>"
case 130,135
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_yb")=jstl
		Application.UnLock
		abc="<a href='fafang/yb.asp?tl="&Application("ysjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
		sayyq="<font color=red>����Ϣ���������Ǻ����ӣ�"&Application("ysjh_user")&"���˽������ͷ�Ƚ���500��ѽ�����������ֵ���һ����������ֵ��+"&jstl&"��!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee><bgsound src=wav/diaoxia.wav loop=1>"
case 160,170
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=jstl
		Application.UnLock
		abc="<a href='fafang/boy.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
		sayyq="<bgsound src=wav/Gui.WAV loop=1><font color=red>����Ϣ��һ��˥����˵�����������Ů�ر�࣬�ܵ�����������Ů�����׼���ð��Ӵ�ѽ�������˹ٸ��н�����˥��������+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 180,182
		jstl=int(rnd*50000)+100000
		Application.Lock
		Application("ysjh_qi")=jstl
		Application.UnLock
		abc="<a href='fafang/qi.asp?tl="&Application("ysjh_qi")&"' target='d'><img src='img/tank.gif' border='0'></a>"
		sayyq="<bgsound src=wav/Bombs020.wav loop=1><font color=red>����Ϣ��һ�Ѿ�����ǹ����������������ʿ����Ŀ�ɿڴ���Ҳ��֪������ʲô���������ȴ�����˵����������ǹ������+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 190
		rftl=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=rftl
		Application.UnLock
		abc="<a href='fafang/rf.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/rf.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/by.wma loop=1><font color=#FF6600>����������׽���˷���:Ҷ����(ye sanniang)�����ɶ���˷��ӣ������ǿ����ĺ����յ�����ȥ����Ҫ�����Ҳ����Ļ����Ҿͼ�һ������ɱһ�������˷���������+"&rftl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 200
		piaoke=int(rnd*50000)+10000
		Application.Lock
		Application("ysjh_kl")=piaoke
		Application.UnLock
		abc="<a href='fafang/piaoke.asp?tl="&Application("ysjh_kl")&"' target='d'><img src='img/piaoke.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/xiaotou.wav loop=1><font color=#FF6600>����������ץС͵����ʼ������ͷ����Ҳ��ƽ��������Ȼ��λ��ͷ���Ե�С͵���뽭����˭������ס������л����</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 201
         r=int(rnd*30)+1
Application.Lock
		Application("ysjh_klt1")=r
Application.UnLock
	        sayyq="<font color=green>������͵Ϯ��</font><font color=red>ͻȻһֻ��¹��͵Ϯ<font color=#0000FF>" &ysjh_name& "</font>,��ȡ<font color=#0000FF>" &ysjh_name& "</font>����"&Application("ysjh_klt1")*500&"</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc1.asp?r="&Application("ysjh_klt1")&" target=optfrm><img src=img/tr003.gif  border=0></a></marquee><bgsound src=../mid/KAI.WAV                loop=3>"
               Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
               conn.execute("update �û� set ����=����-"&Application("ysjh_klt1")*50&" where ����='" &ysjh_name& "'")
case 203
         r=int(rnd*30)+1
Application.Lock
        Application("ysjh_klt3")=r
Application.UnLock
	sayyq="<font color=green>������͵Ϯ��</font><font color=red>ͻȻһֻ�Ӵ�Ŀֲ����޴���������ҧ��" &ysjh_name& "����"&Application("ysjh_klt3")*200&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc3.asp?r="&Application("ysjh_klt3")&" target=optfrm><img src=img/mon24.gif   border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt3")*200&" where ����='" &ysjh_name& "'")
case 204
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt6")=r
Application.UnLock
	sayyq="<font color=green>������͵Ϯ��</font><font color=red>һ�������ĺ����ɨ�������������Ⱥ�����Ű������<font color=#0000FF>" &ysjh_name& "</font>��ҧȥ"&Application("ysjh_klt6")*10&"�������</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc6.asp?r="&Application("ysjh_klt6")&" target=optfrm><img src=img/mon15.gif   border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt6")*10&" where ����='" &ysjh_name& "'")
case 205
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt5")=r
Application.UnLock
	sayyq="<font color=green>��������顿</font><font color=red>һ����͵ķ��������һֻ����ǧ��������Ȼ��������<font color=#0000FF>" & username & "</font>���ӻ���"&Application("ysjh_klt5")*20&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc5.asp?r="&Application("ysjh_klt5")&" target=optfrm><img src=img/mon20.gif   border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set allvalue=allvalue+"&Application("ysjh_klt5")*20&" where ����='" & username & "'")
case 206
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt7")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>һ������ͻϮ��������������һȺ��Ů�ӵ��У����˾�ҧ��<font color=#0000FF>" &ysjh_name& "</font>��ҧ������"&Application("ysjh_klt7")*200&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc7.asp?r="&Application("ysjh_klt7")&" target=optfrm><img src=img/mon23.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt7")*200&" where ����='" &ysjh_name& "'")
case 207
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt8")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>��¿�����ɽ��һ��Ұ��������ɽ���ε�����ʫ������������ʳ������<font color=#0000FF>" &ysjh_name& "</font>����"&Application("ysjh_klt8")*50&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc8.asp?r="&Application("ysjh_klt8")&" target=optfrm><img src=img/mon28.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt8")*50&" where ����='" &ysjh_name& "'")
case 208
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt9")=r
Application.UnLock
	sayyq="<font color=green>������͵Ϯ��</font><font color=red>ħ�ý����ڸ���ɽ������������һ�ɽ�������<font color=#0000FF>" &ysjh_name& "</font>����"&Application("ysjh_klt9")*10&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc9.asp?r="&Application("ysjh_klt9")&" target=optfrm><img src=img/mon27.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt9")*10&" where ����='" &ysjh_name& "'")
case 209
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt10")=r
Application.UnLock
	sayyq="<font color=green>������͵Ϯ��</font><font color=red>��~~~�ÿ��һֻѼ!�������Ӵ��ˣ�ʲô���У�һֻ����Ѽ�ܽ���ʫ������͵��<font color=#0000FF>" &ysjh_name& "</font>����"&Application("ysjh_klt10")*500&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc10.asp?r="&Application("ysjh_klt10")&" target=optfrm><img src=img/mon3.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt10")*500&" where ����='" &ysjh_name& "'")
case 210
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt11")=r
Application.UnLock
	sayyq="<font color=green>������͵Ϯ��</font><font color=red>������ͻȻ���ħ�ý�����һ��������<font color=#0000FF>" &ysjh_name& "</font>��û���ں����ģ�����ȡ����"&Application("ysjh_klt11")*10&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc11.asp?r="&Application("ysjh_klt11")&" target=optfrm><img src=img/mon4.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set allvalue=allvalue-"&Application("ysjh_klt11")*10&" where ����='" &ysjh_name& "'")
case 211
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt12")=r
Application.UnLock
	sayyq="<font color=green>��������顿</font><font color=red>һֻ���겻���Ľ���֮���ξ���ͻȻ�������������ڣ��͸�<font color=#0000FF>" &ysjh_name& "</font>����"&Application("ysjh_klt12")*400&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc12.asp?r="&Application("ysjh_klt12")&" target=optfrm><img src=img/mon10.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����+"&Application("ysjh_klt12")*400&" where ����='" &ysjh_name& "'")
case 212
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt13")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>һֻ��ʷ��ķ�ı�����������ʫ��������<font color=#0000FF>" &ysjh_name& "</font>�����Ƶ���ȡ����"&Application("ysjh_klt13")*100&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc13.asp?r="&Application("ysjh_klt13")&" target=optfrm><img src=img/mon9.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt13")*100&" where ����='" &ysjh_name& "'")
case 213
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt14")=r
Application.UnLock
	sayyq="<font color=green>����긽�塿</font><font color=red>����<font color=#0000FF>" &ysjh_name& "</font>�����ڹ�ʱ��С���߻���ħ�������򱦱��Ļ��鸽�ŵ����ϣ���ҧ��"&Application("ysjh_klt14")*400&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc14.asp?r="&Application("ysjh_klt14")&" target=optfrm><img src=img/mon8.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt14")*400&" where ����='" &ysjh_name& "'")
case 214
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt15")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,��ȡ" &ysjh_name& "����"&Application("ysjh_klt15")*50&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc15.asp?r="&Application("ysjh_klt15")&" target=optfrm><img src=img/Shenmo.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt15")*50&" where ����='" &ysjh_name& "'")
case 215
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt16")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,���" &ysjh_name& "����"&Application("ysjh_klt16")*10&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc16.asp?r="&Application("ysjh_klt16")&" target=optfrm><img src=img/Ying.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt16")*10&" where ����='" &ysjh_name& "'")
case 216
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt17")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,���" &ysjh_name& "����"&Application("ysjh_klt17")*10&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc17.asp?r="&Application("ysjh_klt17")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt17")*10&" where ����='" &ysjh_name& "'")
case 217
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt18")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,͵��" &ysjh_name& "����"&Application("ysjh_klt18")*5&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc18.asp?r="&Application("ysjh_klt18")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb") 
    conn.execute("update �û� set allvalue=allvalue-"&Application("ysjh_klt18")*5&" where ����='" &ysjh_name& "'")
case 218
r=int(rnd*30)+1
Application.Lock
Application("ysjh_klt19")=r
Application.UnLock
	sayyq="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,����" &ysjh_name& "����"&Application("ysjh_klt19")*50&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc19.asp?r="&Application("ysjh_klt19")&" target=optfrm><img src=img/tr003.gif border=0></a></marquee><bgsound src=../mid/KAI.WAV loop=3>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
    conn.execute("update �û� set ����=����-"&Application("ysjh_klt19")*50&" where ����='" &ysjh_name& "'")
case 230
if Application("ysjh_baowu")>0 and ysjh_grade<6 then 
   Set conn=Server.CreateObject("ADODB.CONNECTION") 
   Set rs=Server.CreateObject("ADODB.RecordSet") 
   conn.open Application("ysjh_usermdb") 
   rs.open "select id,����,����,�ȼ�,����,��� from �û� where ����='" & Application("ysjh_baowuname") & "'",conn 
   if rs.eof or rs.bof  then 
    rs.close 
    rs.open "select ���� from �û� where ����='"&ysjh_name&"'" 
    if rs("����")="����" then 
     sayyq1="[������Ϣ]��##�ڽ�������żȻ�����˽�������<img src=img/z1.gif><font color=red>"& Application("ysjh_baowuname")&"</font>���������Լ��Ǹ������ˣ������˽�̰��ֻ�ÿ���һ�۱����Ȼ��ȥ~~~��" 
     sayyq=Replace(sayyq1,"##",zj,1,3,1)  
    else 
     conn.execute  "update �û� set ����=false,��������=0,����='"& Application("ysjh_baowuname") &"' where ����='" & ysjh_name &"'" 
     sayyq1="[������Ϣ]��##����ӵ�н�������<img src=img/z1.gif><font color=red>"& Application("ysjh_baowuname")&"</font>" & Application("ysjh_baowusm") &"��λ��ʿ�������ᣡ" 
     sayyq=Replace(sayyq1,"##",zj,1,3,1)  
    end if 
   else 
    baouser=rs("����") 
    baowuid=rs("id") 
    if Instr(LCase(Application("ysjh_useronlinename"&nowinroom))," "&LCase(baouser)&" ")=0 then 
     conn.execute  "update �û� set ����='��',��������=0 where ����='"& baouser &"'" 
    else 
     sayyq1="[������Ϣ]����������<img src=img/z1.gif><font color=red>"& Application("ysjh_baowuname")&"</font>�ѱ�["&rs("����")&"]"&rs("���")&baouser&"���ߣ���λ��������ȥ��..." 
     sayyq=Replace(sayyq1,"##",zj)  
    end if 
   end if 
   rs.close 
   set rs=nothing 
   conn.close 
   set conn=nothing 
  end if
case 233,235,238,250
'ȡ����Ц��
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
'ȡЦ���ļ�¼����
sql="select count(id) from jokelist"
set rs=Conn.execute(sql)
jokecount=rs(0)
randomize
rand=Int(jokecount * Rnd +1 )
sql="select * from jokelist where id=" & rand & ""
set rs=Conn.execute(sql)
joke=rs("����")&":"&rs("��Ŀ")&":"&rs("����")
rs.close
set rs=nothing
conn.close
set conn=nothing
joke=replace(joke,chr(13),"")
joke=replace(joke,chr(10),"")
                Application.Lock
                Application("ysjh_joke")=joke
                Application.UnLock
                sayyq="<bgsound src=wav/haha.wav loop=1><img src=img/joke.GIF><font color=009933>������һЦ��</font><img src=img/jokepic.GIF><font color=666666>" & Application("ysjh_joke") & "</font>"			'��������
case 265,270
		jstl=int(rnd*2)+1
		Application.Lock
		Application("ysjh_jinbi")=jstl
		Application.UnLock
		abc="<a href='fafang/bxinshou.asp?tl="&Application("ysjh_jinbi")&"' target='d'><img src='img/xinshou.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/xinshou.wav loop=1><font color=#000000>���չ����֡�</font><b><font color=red>���������������ˣ�վ��˵���������������չ������߽�:+"&jstl&"�����!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 280,289,295,305
'�Զ����⿪ʼ
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
'ȡ����¼��
sql="select count(id) from QuestLib "
set rs=Conn.execute(sql)
askcount=rs(0)
randomize
rand=Int(askcount * Rnd +1 )
sql="select * from QuestLib where id=" & rand & ""
set rs=Conn.execute(sql)
ask=rs("����")&":"&rs("����")
reply=rs("�ش�")
rs.close
set rs=nothing
conn.close
set conn=nothing
ask=replace(ask,chr(13),"")
ask=replace(ask,chr(10),"")
                Application.Lock
                Application("ysjh_ask")=ask
                Application("ysjh_reply")=reply
                Application("ysjh_askuser")="������"
                Application("ysjh_asksilver")=int(rnd*9999)+100000
                Application.UnLock
                sayyq="<bgsound src=wav/zhanfa.wav loop=1>��ϵͳ���⡿<font color=balck>" & Application("ysjh_ask") & "�� </font>��ȷ����ʲô��[������]:<font color=red>["&Application("ysjh_askuser")&"]</font><font color=green>������"&Application("ysjh_asksilver")&"��!</font>"
end select			
end if
end if
if sayyq<>"" then
	sayyq=replace(sayyq,"'","\'")
	sayyq=replace(sayyq,chr(34),"\"&chr(34))
	act="����"
	saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & ysjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & sayyq & chr(39) &",0," & nowinroom & ");<"&"/script>"
	AddMsg SayStr
end if
act="����"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
says=says&t
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & ysjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway & "," & nowinroom & ");"&npccls&"<"&"/script>"
AddMsg SayStr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
Response.Write "<script Language=JavaScript>"
Response.Write "parent.f1.scrollWindow();"
Response.Write "</script>"
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