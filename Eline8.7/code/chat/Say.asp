<%@ LANGUAGE=VBScript codepage ="936" %>
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
weekdate=weekday(date())	'����ֵ
sjz=weekdate*10000+s*100+f	'�ᱦ����ʱ��
beginsj=weekdate*100000+s*1000+f*10+m '�ᱦ����׼ʱ��ʼʱ��620001  620005
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
if chatinfo(6)=0 then
'��ʼ����¼�
randomize()
rnd1=int(rnd*400)+1
sayyq=""
if beginsj>=621030 and beginsj<=621040  and application("sjjh_mpdz")<>1 then
   Application.Lock
   Application("sjjh_mpdz")=1              '1Ϊ��ʼ��0Ϊ����
   Application.UnLock
   Set conn=Server.CreateObject("ADODB.CONNECTION") 
   Set rs=Server.CreateObject("ADODB.RecordSet") 
   conn.open Application("sjjh_usermdb")
   conn.Execute ("update �û� set ����ɱ��=0  ")
   conn.close
   set conn=nothing
   sayyq="<bgsound src=wav/lunjian.mid loop=1><table width='85%' bgcolor='#CC6699' bordercolorlight='000000' bordercolordark='FFFFFF' border=1 cellspacing='0' cellpadding='2' align='center'><tr><td height='15'><div align='center'><font color='#FFFFFF' size='4'><font color=yellow face='Wingdings'>[</font><font color='yellow'><b>�� �� �� ս ��</b></font><font color=yellow face='Wingdings'>]</font></font></div></td></tr><tr><td bgcolor='#6699CC'><div align='center'><font color='#000000' size='3'>"&Application("sjjh_user")&"���������ɴ�ս���ڿ�ʼ...</font></div></td></tr></table>" 
end if

if beginsj>=720300 and beginsj<=720305 then
	yyyy="�ᱦ���������ѿ�ʼ���ᱦ��ս��ս���᲻���ڴ������б���..."
	sayyq="<bgsound src=wav/duobao.mid loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>�� �� �� ��</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&yyyy&"</font></div></td></tr></table>"
	rnd1=999999
end if

if rnd1>1 and (Minute(time())=30 and Second(time())<7 or Minute(time())=45 and Second(time())<7) then
sayyq="<bgsound src=wav/02.wav loop=1><div align=center><img src='epk.gif' border='0'></div>"
end if

if rnd1<=7 and (Minute(time())>=20 and Minute(time())<30) then
sayyq="<bgsound src=wav/duo.mid loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>�� ֪ �� ��</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font color=#CC0000 size=-1>�ᱦ����ÿ��������20:30��׼ʱ��[���ַ���]���俪ʼ��ս���ȼ���30�����ϡ��ǹٸ����ǳ�����Ա���ɲμӣ��ᱦ�ھ�<img src=sjfunc/guanjun.gif>����ý�������--�Ͻ��«���������Ͻ��«����Եõ�����3000�㣬���500�����������������书���޸���1000�㡣</font></div></td></tr></table>"
end if

'����������ÿСʱǰ20���ӿ����Զ���Ǯ�����Ըĵġ�
if rnd1<=7 and Minute(time())<8 then
	useronlinename=Application("sjjh_useronlinename"&nowinroom)
	online=Split(Trim(useronlinename)," ",-1)
	x=UBound(online)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("sjjh_usermdb")
	select case rnd1
		case 1,5
			randomize timer
			s=int(rnd*5000000)
			for i=0 to x
				conn.Execute "update �û� set ����=����+" & s & " where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>��ϵͳ��Ǯ��</font><B><font color=#ff0000>��Ǯඣ���Ǯඣ�<img src=img/251.gif><img src=img/251.gif><img src=img/251.gif> ���������ÿ���˷���"& s &"������������ַ��www.happyjh.com�����˽�������һ����Ŷ ^_^</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/022.gif><img src=img/022.gif><img src=img/022.gif><img src=img/022.gif></marquee>"
		case 2,7
			for i=0 to x
				conn.Execute "update �û� set ���=���+5 where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>��ϵͳ��ҡ�</font><B><font color=#ff0000>վ�����׵����л��Ҷԡ����ֽ�������֧�ֺͺ񰮣����и��ˣ���Ϊ�����ҵ�ÿ���˷��˽��5����</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif></marquee>"
		case 3
			'randomize timer
			'mm=int(rnd*50)
			mm=50
			for i=0 to x
				conn.Execute "update �û� set allvalue=allvalue+" & mm & " where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>��ϵͳ��㡿</font><B><font color=#ff0000>վ��·���˵أ����̿�����λ��СϺ��˿�������Ϊ�����ҵ�ÿ���˷��ž���" & mm & "�㣡</font></b>"
		case 4,6
			randomize timer
			s=int(rnd*100)
			for i=0 to x
				conn.Execute "update �û� set �ݶ�����=�ݶ�����+" & s & ",��=��+" & s & ",ľ=ľ+" & s & ",ˮ=ˮ+" & s & ",��=��+" & s & ",��=��+" & s & " where ����='" & online(i) & "'"
			next
			sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>��ϵͳ���š�</font><B><font color=#ff0000>վ�����׵���·���˵أ��������ĸ�λ���Ŷ���" & s & "��...�Ǻǣ�������:Ϊ�˹���������������书��ԯ������λ���Ͻ�ľ��ˮ���������Ը�" & s & "�㡣</font></b>"
	end select
	conn.close
	set conn=nothing
else
select case rnd1
case 8
banner="��л���֧�֡����ֽ�����.���׵��꽫�������׽����ϸ����ĵ�����������.������κ��������ʱ�����.ֻҪ���׵����ڶ�������ʱ���!�������ĵ�ַ��www.happyjh.com�����˽�������һ����Ŷ^_^;�����ֽ�����:����������ǿ�Ľ������������ĵ�ַ��www.happyjh.com.������������׳��򲻴�.�������������.����(��)�����˻���!�����������ǵĿ��ּ�԰!;�����ֽ����������û�ע��ʱ��ӵ��ս���ȼ�10��.����10000000.��һ�ε�½ϵͳ���Զ��������˷�.�������Ƕ��и��ϸߵ����.ʹ�ø���Ĺ���.��10���û�������Ǯ����ת�ʲ�������.��������ϧ��ĵ����ʺ�.��������ʺų���3��.����ʱɾ��.ˡ�����.;�����ֽ������ɻ��׵�������İ�����.ʹ��Ȩ����׵����������.����ṩ�������֡����ѽ���.ע���û���ʼ��һ����Ա.�����������������á��ȼ���60���ɼ�����Ϊ������Ա.֮���90������Ϊ������Ա.��120������Ϊ�ļ���Ա.��300������Ϊ�弶��Ա.�����������Զ�ȡ����Ա�ʸ�.����ǰ5��ϵͳ����ʾ��.л������֮�ͣ��������ĵ�ַ��www.happyjh.com�����˽�������һ����Ŷ^_^;վ����������:�������ѧ��.��ע�ⲻҪӰ����Ϣ.������Ӱ��ѧϰ.�����Թ����ݵ�.����Ҫ���⻨̫���ʱ��;���.�����ֽ���������ּ���������֡���ʶ����.ϣ�����ܸ�����������֪ʶ�ͻ���!��������˶������˸���Ҫ��ѧϰ����.�м�!;����ҳ��һ������������ˡ����ֽ������Ŀ������.���⻹����ǰ��4����½ҳ��.��½ʱ������ѡ���.Ҳ������ϵͳ���.Ŀǰ����11���������ķ��ҳ��.�����.����������35��.�����.ˢ�¾Ϳ���������ͬ��������.��20���������ʽ�Լ�����Ư���Ĺ�����.��������IE�汾��6.0����.�������.�Ͽ�������;Ϊ��ȷ�����ݵİ�ȫ.���Ҽ�ʱ�������뱣��.�����޸�����.��Ҫ�����Լ����ʺ�����������.���ⶪʧ.����������.�������ĵ�ַ��www.happyjh.com�����˽�������һ����Ŷ^_^;�����ֽ�������ȫ���.�������ҿ������֡����Խ��ѡ����Խ���.�����������ǹ�ͬά����������.���ؽ�������.Ӫ�����÷�Χ.������ֶ������������.��ɱ�������߷���IP.���Ծ�����!;�����ֽ�����ÿСʱǰ10����ϵͳ��Ǯ.���ֵ1000����.�ݵ�ÿ��������50000.�书����50.���10.վ��Ҳ��ʱ�������ߵ����ѷ�����������ҡ��������������������书��.ϣ�����ڳ�Ŷ.^_^.�벻Ҫ�Լ���վ��Ҫ��Ҫ�ǵ�.����������.��Ϊ����û���ܵ���.�мǣ�;����PK�����䣩ʱ��ΪÿСʱ�ĺ�30����.�����������볬��3�˲ſ��Խ���.һ���ɱ����Ϊ10.�ᱦһ����һ��ɱ��.����ٸ���Ա����.����ʱ�䵽ʱ�뼰ʱ֪ͨ.��ҿ��Ծ����.�����ֺͲ�Ը���ܵĿ�����������.�Է���ɱ.���˺���Ŷ^_^;��ս���ȼ���30�����ϡ��ǹٸ����ǳ�����Ա���ɲμӶᱦ����.�ᱦ����ÿ��������20:30����ʼ.Ҫ�μӶᱦ��������Աֻ���ڴ�����ʼǰ10�����ڽ���[���ַ���]����(��20:20-20:30֮�����).��ǰ�򳬹���ʱ�����޷�����ᱦ֮ս����.;�ᱦ����ʱ��ֹʹ��Ǭ��һ�������Ǵ󷨡�͵Ǯ�����������弼����Ƭ�����������롢���ҡ����ס�ͬ���ھ�������Ѩ����ҩ������.����֮��Ĺ��ܾ�����ʹ��.����ʱ�������촰�����½ǵĶᱦ(�ڷǶᱦ֮ս�����ֹʹ��).�ᱦǰ10���ӽ���[���ַ���]����.������ʼ�����������Ƿ�չر���.�Ƿ����붼���Դ��˻򱻴�.�ᱦ��ɺ�.�������Ͻ��«����Եõ��Ľ�����:����3000��.���500��.�������������书���޸���1000��.;����������ע���������ǳ����潭���������������.��ʲô�����벻Ҫ�����ʹٸ��ʱ���.�������ȵ���̳����鿴.�������б���������ϸ�ƶȺ͹���.Ҳ�в�����ҵ��ĵ�����.���Ż���������ܶ����.�������Ľ�������ֵ!ף������!"
banners=Split(Trim(banner),";",-1)
total=UBound(banners)
randomize timer
x=int(rnd*(total+1))
sayyq="<bgsound src=wav/gonggao.wav loop=1><table width=85% cellspacing=0 cellpadding=2 bgcolor=#009933 align=center bordercolorlight=000000 bordercolordark=FFFFFF border=1><tr><td align=center><font color=FFFFFF>�� ϵ ͳ �� Ϣ ��</font></td><tr><td bgcolor=CCCCFF> <div align='center'><font size=-1>"&banners(x)&"</font></div></td></tr></table>"
case 9,7,5
'ȡ����Ц��
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("ejoke.asp")
Conn.Open connstr
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
                Application("sjjh_joke")=joke
                Application.UnLock
                sayyq="<bgsound src=wav/haha.WAV loop=1><img src=joke.GIF><font color=009933>������һЦ��</font><img src=jokepic.GIF><font color=666666>" & Application("sjjh_joke") & "</font>"			'��������				
case 11
		Application.Lock
		Application("sjjh_dalie")="�ϻ�"
		Application.UnLock
		sayyq="<bgsound src=wav/hu.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>ͻȻһֻҰ��<img src=img/laohu.gif>���뽭�������ˣ�������ǿ�ȥ���԰���</font></b>"
case 12
		jstl=int(rnd*50000)+1000
		Application.Lock
		Application("sjjh_kl")=s
		Application.UnLock
		abc="<a href='dog.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/dog.gif' border='0'></a>"
		sayyq="<bgsound src=wav/dog.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>ͻȻ��һ�����С���������ѽ�������������¹����˼�У�<bgsound src=wav/gougou.wav loop=1>��е���Ҫҧ�ң���Ҫҧ�ң������£�һͷ�������������ң�����������+"&jstl&"</font><b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 13
		jstl=int(rnd*200000)+10000
		Application.Lock
		Application("sjjh_yb")=jstl
		Application.UnLock
		abc="<a href='yb.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>������ʲô���ӣ�ͻȻ������ϵ�����һ����Ԫ������ֵ��+"&jstl&"��!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 14
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("sjjh_kl")=jstl
		Application.UnLock
		abc="<a href='ws.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/bt.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/bsj.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>�ۣ������ֽ���������һ����ʿ��Ҫ�Ҵ�ұ��䡣��ʿ������"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 15
		jstl=int(rnd*500000)+10000
		Application.Lock
		Application("sjjh_yb")=jstl
		Application.UnLock
		abc="<a href='yb.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>�������Ǻ����ӣ�"&Application("sjjh_user")&"����E�߷��ͷ�Ƚ���500��ѽ�����������ֵ���һ����������ֵ��+"&jstl&"��!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 16
		jstl=int(rnd*50000)+1000
		Application.Lock
		Application("sjjh_kl")=jstl
		Application.UnLock
		abc="<a href='lj.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/lj.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/baguo.WAV loop=1><font color=red>����Ϣ��</font><b><font color=red>�˹��������֡����ֽ�������������Ů��ɱ����ѽ����ѽ~~~~~~�˹�����������"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 17
		jstl=int(rnd*50000)+1000
		Application.Lock
		Application("sjjh_kl")=jstl
		Application.UnLock
		abc="<a href='boy.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
		sayyq="<bgsound src=wav/FAQIAN.WAV loop=1><font color=red>����Ϣ��</font><b><font color=red>һ��˥����˵�����������Ů�ر�࣬�ܵ�����������Ů�����׼���ð��Ӵ�ѽ�������˹ٸ��н�����˥��������+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 18
		jstl=int(rnd*10000)+10000
		Application.Lock
		Application("sjjh_kl")=jstl
		Application.UnLock
		abc="<a href='mm.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/mm.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/wn2.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>�׹�~~~�����ֽ���������λ����Ů����˭�ܱ�����Ů�飬��Ů������"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 19
		jstl=int(rnd*20000)+10000
		Application.Lock
		Application("sjjh_qi")=jstl
		Application.UnLock
		abc="<a href='qi.asp?tl="&Application("sjjh_qi")&"' target='d'><img src='img/tank.gif' border='0'></a>"
		sayyq="<bgsound src=wav/Bombs020.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>һ�Ѿ�����ǹ���������ֽ���������ԭ������ʿ����Ŀ�ɿڴ���Ҳ��֪������ʲô���������ȴ�����˵����������ǹ������+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 20
		jstl=int(rnd*10000)+10000
		Application.Lock
		Application("sjjh_kl")=jstl
		Application.UnLock
		abc="<a href='wn.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/wn.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/wn1.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>վ��ΪŮ�����Ǵ����ˡ����ֽ�������˧�������������������ǵ�˫�֡�ҡ����~~~��������������"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 21
		jstl=int(rnd*10000)+10000
		Application.Lock
		Application("sjjh_ybb")=jstl
		Application.UnLock
		abc="<a href='ybb.asp?tl="&Application("sjjh_ybb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/Phant008.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>E����ǲ���Ӵ��������ֽ�����������ɨˮͰ�����������ҵ��˻�������������ÿ�����������ߣ�-"&jstl&"��!</font></b><br><marquee width=100% behavior=alternate scrollamount=15><img src='img/10.GIF' border='0'>"&abc&"</marquee>"
case 22
	jstl=int(rnd*5)+1
	Application.Lock
	Application("sjjh_by")=jstl
	Application.UnLock
	abc="<a href='by.asp?tl="&Application("sjjh_by")&"' target='d'><img src='img/by.GIF' border='0'></a>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>������Ϣ����֪��˭��ҩ����·���ˣ�+"&jstl&"!˭������˭�ģ�</font></b><br><marquee width=100% behavior=alternate scrollamount=50>"&abc&"</marquee>"
case 23
	jstl=int(rnd*100)+10
	Application.Lock
	Application("sjjh_yb")=jstl
	Application.UnLock
	abc="<a href='cd.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/cd.GIF' border='0'></a>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>������㡿�����ֽ�������������+"&jstl&"��!˭������˭�ģ�</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 24
	jstl=int(rnd*50)+10
	Application.Lock
	Application("sjjh_yb")=jstl
	Application.UnLock
	abc="<a href='dd.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/dd.GIF' border='0'></a>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>�������㡿�����ֽ��������������+"&jstl&"��!˭������˭�ģ�</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 25
	jstl=int(rnd*2)+1
	Application.Lock
	Application("sjjh_yb")=jstl
	Application.UnLock
	abc="<a href='jk.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/jk.GIF' border='0'></a>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>�����𿨡������ֽ������������+"&jstl&"��!˭������˭�ģ�</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 26
		jstl1=int(rnd*50000)+100
		Application.Lock
		Application("sjjh_yb1")=jstl1
		Application.UnLock
		abc="<a href='apple.asp?tl="&Application("sjjh_yb1")&"' target='d'><img src='img/apple.gif' border='0'></a>"
		sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>ͻȻ������ϵ�����һ��ƻ�������˼�����"&jstl1&"��!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 27
		rftl=int(rnd*50000)+10000
		Application.Lock
		Application("sjjh_rf")=rftl
		Application.UnLock
		abc="<a href='rf.asp?tl="&Application("sjjh_rf")&"' target='d'><img src='img/rf.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/by.wav loop=1><b><font color=#FF0000>��׽�˷���</font></b><font color=#663300>�˷���:Ҷ����ɶ���˷��ӣ������ǿ����ĺ����յ�����ȥ����Ҫ�����Ҳ����Ļ����Ҿͼ�һ������ɱһ�������˷���������+"&rftl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 28
		piaoke=int(rnd*50000)+10000
		Application.Lock
		Application("sjjh_piaoke")=piaoke
		Application.UnLock
		abc="<a href='piaoke.asp?tl="&Application("sjjh_piaoke")&"' target='d'><img src='img/piaoke.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/xiaotou.wav loop=1><b><font color=#FF0000>��ץС͵��</font></b><font color=#663300>����ʼ:����ͷ����Ҳ��ƽ��������Ȼ��λ��ͷ���Ե�С͵���뽭����˭������ס������л����</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 29
		jstl=int(rnd*2)+1
		Application.Lock
		Application("sjjh_jinbi")=jstl
		Application.UnLock
		abc="<a href='jinbi.asp?tl="&Application("sjjh_jinbi")&"' target='d'><img src='img/jinbi.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>������ʲô���ӣ�ͻȻ�����ϵ���+"&jstl&"�����!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 30
    jstl=int(rnd*50000)+1000
	Application.Lock
	Application("sjjh_kl")=jstl
	Application.UnLock
	abc="<a href='DABIAN.ASP?tl="&Application("sjjh_kl")&"' target='d'><img src='img/ying1.gif' border='0'></a>"
    sayyq="<bgsound src=wav/niao.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>һֻ�����ڤ���Ƿ��ٽ��������ң�̫���񰡣�һ�����кö����ģ���ҿ��!��������"&jstl&"</font></b><br><marquee width=100%  scrollamount=38>"&abc&"</marquee>"
case 31
	jstl=int(rnd*0)+1
	Application.Lock
	Application("sjjh_kp")=jstl
	Application.UnLock
	abc="<a href='kp2.asp?tl="&Application("sjjh_kp")&"' target='d'><img src='img/kapian.GIF' border='0'></a>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>һ�Ž�������Ŀ�Ƭ���ڽ���·�ϣ���λ��СϺ������!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 32
	jstl=int(rnd*3)+1
	Application.Lock
	Application("sjjh_py")=jstl
	Application.UnLock
	abc="<a href='py2.asp?tl="&Application("sjjh_py")&"' target='d'><img src='img/py.GIF' border='0'></a>"
	sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>����û���ɣ���ҩɢ��һ�أ�����ҩ����·�ϣ�˭����˭��!</font></b><br><marquee width=100% behavior=alternate scrollamount=45>"&abc&"</marquee>"
case 33
	jstl=int(rnd*0)+1
	Application.Lock
	Application("sjjh_lw")=jstl
	Application.UnLock
	abc="<a href='lw.asp?tl="&Application("sjjh_lw")&"' target='d'><img src='img/lw.jpg' border='0'></a>"
	sayyq="<bgsound src=wav/zhanfa.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>վ��Ϊ����ĺ������������ˣ�ףԸ������쿪�ġ��������⣡</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 34
	jstl=int(rnd*50000)+10000
	Application.Lock
	Application("sjjh_kl")=jstl
	Application.UnLock
	abc="<a href='liumang.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/liumang.gif' border='0'></a>"
	sayyq="<bgsound src=wav/zhu.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>һ�������˵�����������Ů�ر�࣬�ܵ�����������Ů�����׼���ð��Ӵ�ѽ�������˹ٸ��н�������å���������"&s+100&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 35
		jstl1=int(rnd*50000)+100
		Application.Lock
		Application("sjjh_yb1")=jstl1
		Application.UnLock
		abc="<a href='bieshu.asp?tl="&Application("sjjh_yb1")&"' target='d'><img src='img/bieshu.gif' border='0'></a>"
		sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>�����ֽ�����Ϊ�˸�л�����ǵ�֧�ֺͺ񰮣��ͺ�������һ��������������"&jstl1&"��!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 36
		jstl=int(rnd*50000)+10000
		Application.Lock
		Application("sjjh_kl")=jstl
		Application.UnLock
		abc="<a href='shu.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/shu.gif' border='0'></a>"
		sayyq="<bgsound src=wav/shu.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>ͻȻ�佭������һֻ��������󡭡�������ѽ������������������˼�У���е���Ҫҧ�ң���Ҫҧ�ң������£�����һͷ���󴳽������ң�����������+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 37
		jstl=int(rnd*2)+1
		Application.Lock
		Application("sjjh_jinbi")=jstl
		Application.UnLock
		abc="<a href='bxinshou.asp?tl="&Application("sjjh_jinbi")&"' target='d'><img src='img/xinshou.GIF' border='0'></a>"
		sayyq="<bgsound src=wav/xinshou.wav loop=1><font color=#000000>���չ����֡�</font><b><font color=red>���������������ֽ������ˣ�վ��˵���������������չ������߽�:+"&jstl&"�����!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 38,2
'����ֵ*1000+Сʱֵ*100+���ӣ���������19��50��Ϊ6*10000+19*100+50=61950����������20:00��Ϊ6*10000+20*100+0=62000
	if weekdate=7 and sjz<72030 then
		Set conn=Server.CreateObject("ADODB.CONNECTION")
		Set rs=Server.CreateObject("ADODB.RecordSet")
		conn.open Application("sjjh_usermdb")
		rs.open "select * from �ᱦ",conn
		if (rs("��ȡ")=true or rs("����")=true) then
			conn.execute "update �ᱦ set �ھ�=0,��ȡ=false,��������=0,����=false,ʱ��=now()"
		end if
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		if sjz>=71820 and sjz<72023 then
			yyyy="�����ᱦ����������ʼ��������Ȥ�ᱦ������<font color=red><b>20:20��20:30</b></font>֮�����ᱦ����[���ַ���]������20:30�󽫲����ٽ���ᱦ���䣬�ᱦ����<b>20:30</b>��׼ʱ��ʼ�����λ����׼������������[���ַ���]������ֻʣ���һ����ʱΪ������׼���ھ�Ϊ���һ������[���ַ���]�ڵ��ˡ�<font color=blue><b>��ƷΪ������3000�㣬���500������Ϊ�ھ�����������������书���޸�1000�㡣</b></font>"
		else
			if sjz>=72020 and sjz<72030 then
				yyyy="�����ᱦ�������ڿ���<b><font color=blue>��ʼ����ᱦ��ս����[���ַ���]��</font></b>�������������ץ��ʱ����<b><font color=blue>20:30֮ǰ</font></b>����ᱦ��ս���䡣�ᱦ����20:30����ʽ��ʼ���ڶᱦ����ʱ��ֹʹ�ÿ�Ƭ�Լ����弼�����λ�ڽ���ǰ���ó��׼����"
			else
				yyyy="�����ǽ����ᱦ����ʱ�䣬����Ȥ�μ��߿���ÿ����������<b><font color=red>20:20��</font>��<font color=red>20��30��</font></b>֮�����[���ַ���]����ǰ�򳬹����ʱ�佫�����ٽ���[���ַ���]���ᱦ������[���ַ���]ֻʣ���һ����ʱΪ������־���������˼�Ϊʤ���ߡ��ڶᱦ����ʱ��ֹʹ�ÿ�Ƭ�����弼����ƷΪ����3000�㣬���500��������������������书���޸�1000�㡣"
			end if
		end if
		sayyq="<bgsound src=wav/tixing.wav loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>�� �� �� ��</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&yyyy&"</font></div></td></tr></table>"
	end if
case 39
		jstl=int(rnd*5000)+1000
		Application.Lock
		Application("sjjh_kl")=jstl
		Application.UnLock
		abc="<a href='kl.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/cins.GIF' border='0' width='60' height='80'></a>"
		sayyq="<bgsound src=wav/gui.wav loop=1><font color=red>����Ϣ��</font><b><font color=red>ͻȻ������³¡�������ʬѽ����һŮ�Ӽ�У�һͷ��ʬ���������ң���ʬ������+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case 42,43,44
'ȡ�������
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("questions.asp")
Conn.Open connstr
'ȡ���ļ�¼����
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
                Application("sjjh_ask")=ask
                Application("sjjh_reply")=reply
                Application("sjjh_askuser")="������"
                Application("sjjh_asksilver")=int(rnd*9999)+100000
                Application.UnLock
                sayyq="<bgsound src=wav/chuti.WAV loop=1><b>��ϵͳ���ʡ�<font color=balck>" & Application("sjjh_ask") & "</font>��ȷ����ʲô��[������]:<font color=red>��" & Application("sjjh_askuser")&"��</font></b><font color=green><b>������"&Application("sjjh_asksilver")&"��!</b></font>"			'��������		
case 10
 if Application("sjjh_baowu")>0 and sjjh_grade<6 then 
   Set conn=Server.CreateObject("ADODB.CONNECTION") 
   Set rs=Server.CreateObject("ADODB.RecordSet") 
   conn.open Application("sjjh_usermdb") 
   rs.open "select id,����,����,�ȼ�,����,��� from �û� where ����='" & Application("sjjh_baowuname") & "'",conn 
   if rs.eof or rs.bof  then 
    rs.close 
    rs.open "select ���� from �û� where ����='"&sjjh_name&"'" 
    if rs("����")="����" then 
     sayyq1="������Ϣ��##�ڽ�������żȻ�����˽�������<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname")&"</font>���������Լ��Ǹ������ˣ������˽�̰��ֻ�ÿ���һ�۱����Ȼ��ȥ~~~��" 
     sayyq=Replace(sayyq1,"##",zj,1,3,1)  
    else 
     conn.execute  "update �û� set ����=false,��������=0,����='"& Application("sjjh_baowuname") &"' where ����='" & sjjh_name &"'" 
     sayyq1="������Ϣ��##����ӵ�н�������<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname")&"</font>" & Application("sjjh_baowusm") &"��λ��ʿ�������ᣡ" 
     sayyq=Replace(sayyq1,"##",zj,1,3,1)  
    end if 
   else 
    baouser=rs("����") 
    baowuid=rs("id") 
    if Instr(LCase(Application("sjjh_useronlinename"&nowinroom))," "&LCase(baouser)&" ")=0 then 
     conn.execute  "update �û� set ����='��',��������=0 where ����='"& baouser &"'" 
    else 
     sayyq1="������Ϣ����������<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname")&"</font>�ѱ�["&rs("����")&"]"&rs("���")&baouser&"���ߣ���λ��������ȥ��..." 
     sayyq=Replace(sayyq1,"##",zj)  
    end if 
   end if 
   rs.close 
   set rs=nothing 
   conn.close 
   set conn=nothing 
  end if
end select			
end if
end if
if sayyq<>"" then
	sayyq=replace(sayyq,"'","\'")
	sayyq=replace(sayyq,chr(34),"\"&chr(34))
	act="����"
	saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & sayyq & chr(39) &",0," & nowinroom & ");<"&"/script>"
	AddMsg SayStr
end if
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