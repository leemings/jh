<%
Response.Expires=-1
mycorp=session("Ba_jxqy_usercorp")
mygrade=Session("Ba_jxqy_usergrade")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
maxnosaytime=Application("Ba_jxqy_nosaytime")
username=Session("Ba_jxqy_username")
onlinename=Application("Ba_jxqy_onlinename"&chatroomsn)
lip=Application("Ba_jxqy_lockip")
aberrantname=Application("Ba_jxqy_aberrantname")
msg=server.HTMLEncode(trim(Request.Form("talkmsgtmp")))
namecolor=Request.Form("namecolor")
wordcolor=Request.Form("wordcolor")
sendto=Request.Form("sendto")
expression=Request.Form("expression")
isprivacy=Request.Form("isprivacy")
act=0
if isprivacy="on" and sendto<>"���" then 
	isprivacy=1
else
	isprivacy=0
end if
if username="" or instr(onlinename,";"&username&";")=0 then Response.Redirect "chaterror.asp?id=000"
if len(msg)>100 then msg=Left(msg,100)
if msg="" or msg="//" or msg="/#" then Response.End
nowtime=now()
if instr(lip,";"&session("Ba_jxqy_userip")&";")<>0 then Response.Write "<script language=javascript>top.location.href="&chr(34)&"chaterror.asp?id=008"&chr(34)&";</script>"
nosaytime=datediff("s",session("Ba_jxqy_userlasttalktime"),nowtime)
if nosaytime>maxnosaytime then 
	Response.Write "<script language=javascript>top.location.href="&chr(34)&"chaterror.asp?id=001"&chr(34)&";</script>"
	Response.End
elseif nosaytime<3 then
	Response.Write "<script language=JavaScript>{alert('�л�����˵����ҭ���ˣ�');}</script>"
	Response.End
end if
if not (sendto="���" or instr(onlinename,";"&sendto&";")<>0) then
	Response.Write "<script Language=Javascript>alert('"&sendto&"�����������У����ܶ��䷢�ԡ�');parent.resfrm.location.href='onlinelist.asp';</script>"
	Response.end
end if
session("Ba_jxqy_userlasttalktime")=nowtime
if instr(aberrantname,username)<>0 then
	abl=Application("Ba_jxqy_aberrantlist")
	ablubd=ubound(abl)
	for i=1 to ablubd step 4
		tt=datediff("s",abl(i+3),nowtime)
		if abl(i)=username and tt<0 and abl(i+2)="���" then
			Response.Write "<script language=javascript>alert('����У��������ã�');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="����"and left(msg,2)="//" then
			Response.Write "<script language=javascript>alert('����״̬�У��޷�ʹ���κ����');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="���" then
			olt=Application("Ba_jxqy_onlinelist")
			oltubd=ubound(olt)
			randomize()
			rnds=clng((oltubd\8-1)*rnd())
			sendto=olt(rnds*8+1)
		elseif abl(i)=username and tt<0 and instr(";����;����;����;����;",";"&abl(i+2)&";") then
			Response.Redirect "chaterror.asp?id=000"
		end if	
	next 
end if
if left(msg,2)="//" then
	if chatroomsn=1 or Application("Ba_jxqy_systemname"&chatroomsn)=mycorp or mycorp="�ٸ�"  then
	act=Trim(mid(msg,3,3))
	msg=Trim(mid(msg,6))
	if instr(msg,"'")<>0 or (msg="" and  instr(";�鿴;����;��Ǯ;���;����;����;͵��;",";"&act&";")=0) then
		Response.Write "<script language=javascript>alert('"&act&"����Ӧ�ô��к��ʵĲ�����');</script>"
		Response.End
	elseif instr(onlinename,";"&sendto&";")=0 and instr(";����;�鿴;Ͷ��;��Ǯ;����;���;���;����;����;����;͵��;����;",";"&act&";")<>0 then
		Response.Write "<script language=javascript>alert('"&act&"����˭����');</script>"
		Response.End
	'elseif sendto=username and instr(";����;Ͷ��;��Ǯ;����;���;���;����;����;����;͵��;����;",act)<>0 then
	'	Response.Write "<script language=javascript>alert('"&act&"���Լ�����');</script>"
	'	Response.End
	end if
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	isprivacy=0
	select case act
		case "�鿴"
			msg=seeequip(username,sendto)
			isprivacy=1
		case "����"
			msg=attack(username,sendto,msg,mygrade)
		case "Ͷ��"
			msg=cast(username,sendto,msg)
		case "��Ǯ"
			msg=sendmoney(username,sendto,msg)
		case "����"
			msg=alert(username,sendto,mycorp,mygrade,chatroomsn,msg)
		case "��Ǯ"
			msg=distributemoney(username,mycorp,mygrade,chatroomsn)
		case "����"
			msg="[msg]"&msg&"[/msg]"
		case "����"
			msg=bulletin(username,sendto,mycorp,mygrade,chatroomsn,msg)	
		case "���"
			msg=degree(username,sendto,mycorp,msg)
		case "���"
			msg=dislodge(username,sendto,mycorp)
		case "����"
			msg=impart(username,sendto,msg)
		case "����"
			msg=burglemp(username,sendto)
		case "͵��"
			msg=burglemoney(username,sendto)
		case "����"
			msg=present(username,sendto,msg)
		case "����"
			msg=card(username,sendto,msg)	
		case else
			msg="<font color=FF0000>��ϵͳ��</font>##���Բ��𣬲�������޷�ִ��"	
	end select
	act=2
	set rst=nothing
	conn.Close
	set conn=nothing
	else
		Response.Write "<script language=javascript>alert('�㲻�Ǳ�����ӣ������ڴ���Ұ��');</script>"
		Response.End
	end if
else
	if left(msg,2)="/#" then
		msg="##<font color="&wordcolor&">"&mid(msg,3)&"</font>"
		act=1
	end if
	randomize()
	rnd1=cint(rnd()*300)
	if rnd1<8 then
		chance=array("����","����","����","����","����","����","����","����","����")
		chancetxtgood=array("�����ִ�Ħ��ʦ��̳������##��������","�䵱���������������Ღ�����##��������","��Ѫ���������ػ��˼䣬ͻ������##��������","����������϶����䣬�����������##��������","##�����˾ƣ�ȴ����������͵����Ʒ����ʹ���������࣬��������",";������ͯ�ܲ�ͨ��##֪�����ݷֵ���Ҫ�ԣ����ָ�������","##���ڽ������ߣ�����MM��ȴ��������Ӳ��������","##���ڻ�ɽ�������򾣬����Ҳ�벻���������˵�еĽ��������������˶�����","�ڷ粨ͤ����������Ĺ��##��������")
		chancetxtbad=array("##��ʱ�䰾ս�����,���¾������ã���·Ҳˤ��,������˶��½���","##��·�ϼ�����һ��������С���������֮�ģ�����ȴ������͵͵�ػ�ȥ������","##�չؿ�����ʬ������ȴ����Ҫ�죬����ֻ�˶�����","��Ц##û����ѧѧ���������Բн��������ֻ��������˰̣�������˶�����","##�������¹��˹�仨¥����������֪�������½���","�����������ʵȻ�������ۣ�ָ��##˵����������ɣ����������ʾ������,��Ȼ##����Ǻ�ش�ԩ�������ֻ��Ǳ���ȥ��","##���Ѳ���,������ƭȥ��","##�����䳡��ʶ��Ѧ󴣬��̸�����������˶�������","##;���տն������İ�ʦѧ�գ����Ϲ���û��ѧ��������ȴ�½���")
		rnd2=cint(rnd()*200)-80
		if rnd2>0 then 
			rndtxt=chancetxtgood(rnd1)&rnd2
		else
			rnd2=rnd2-1
			rndtxt=chancetxtbad(rnd1)&abs(rnd2)
		end if	
		set conn=Server.CreateObject("adodb.connection")
		conn.open Application("Ba_jxqy_connstr")
		conn.execute "update �û� set "&chance(rnd1)&"="&chance(rnd1)&"+"&rnd2&" where ����='"&username&"'"
		conn.close
		set conn=nothing
		talkarr=Application("Ba_jxqy_talkarr")
		talkpoint=clng(Application("Ba_jxqy_talkpoint"))
		dim nta(600)
		j=1
		for i=11 to 600 step 10
			nta(j)=talkarr(i)
			nta(j+1)=talkarr(i+1)
			nta(j+2)=talkarr(i+2)
			nta(j+3)=talkarr(i+3)
			nta(j+4)=talkarr(i+4)
			nta(j+5)=talkarr(i+5)
			nta(j+6)=talkarr(i+6)
			nta(j+7)=talkarr(i+7)
			nta(j+8)=talkarr(i+8)
			nta(j+9)=talkarr(i+9)
			j=j+10
		next
		nta(591)=talkpoint+1
		nta(592)=1
		nta(593)=0
		nta(594)=username
		nta(595)="���"
		nta(596)=""
		nta(597)=namecolor
		nta(598)=wordcolor
		nta(599)="<font color=FF0000>��������</font>"&rndtxt&"<font class=timsty>("&nowtime&")</font>"
		nta(600)=chatroomsn
		Application.Lock
		Application("Ba_jxqy_talkpoint")=talkpoint+1
		Application("Ba_jxqy_talkarr")=nta
		Application.UnLock
		erase nta
		erase chancetxtbad
		erase chancetxtgood
		erase chance
	end if	
end if
msg=replace(msg,"\","\\")
msg=replace(msg,"/","\/")
msg=replace(msg,chr(34),"\"&chr(34))
msg=replace(msg,chr(39),"\"&chr(39))
talkarr=Application("Ba_jxqy_talkarr")
talkpoint=clng(Application("Ba_jxqy_talkpoint"))
dim newtalkarr(600)
j=1
for i=11 to 600 step 10
	newtalkarr(j)=talkarr(i)
	newtalkarr(j+1)=talkarr(i+1)
	newtalkarr(j+2)=talkarr(i+2)
	newtalkarr(j+3)=talkarr(i+3)
	newtalkarr(j+4)=talkarr(i+4)
	newtalkarr(j+5)=talkarr(i+5)
	newtalkarr(j+6)=talkarr(i+6)
	newtalkarr(j+7)=talkarr(i+7)
	newtalkarr(j+8)=talkarr(i+8)
	newtalkarr(j+9)=talkarr(i+9)
	j=j+10
next
newtalkarr(591)=talkpoint+1
newtalkarr(592)=act
newtalkarr(593)=isprivacy
newtalkarr(594)=username
newtalkarr(595)=sendto
newtalkarr(596)=expression
newtalkarr(597)=namecolor
newtalkarr(598)=wordcolor
newtalkarr(599)=msg&"<font class=timsty>��"&time()&"��<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
dim showarr()
mytalkpoint=session("Ba_jxqy_usertalkpoint")
newtalkpoint=0
talkarr=Application("Ba_jxqy_talkarr")
j=1
for i=1 to 600 step 10
	newtalkpoint=talkarr(i)
	if newtalkpoint>mytalkpoint and cstr(talkarr(i+9))=cstr(chatroomsn) and (talkarr(i+2)="0" or talkarr(i+4)="���" or (talkarr(i+2)="1" and (talkarr(i+3)=username or talkarr(i+4)=username))) then
	redim preserve showarr(j),showarr(j+1),showarr(j+2),showarr(j+3),showarr(j+4),showarr(j+5),showarr(j+6),showarr(j+7),showarr(j+8),showarr(j+9)
	showarr(j)=talkarr(i)
	showarr(j+1)=talkarr(i+1)
	showarr(j+2)=talkarr(i+2)
	showarr(j+3)=talkarr(i+3)
	showarr(j+4)=talkarr(i+4)
	showarr(j+5)=talkarr(i+5)
	showarr(j+6)=talkarr(i+6)
	showarr(j+7)=talkarr(i+7)
	showarr(j+8)=talkarr(i+8)
	showarr(j+9)=talkarr(i+9)
	j=j+10
	end if
next
if j>1 then
Response.Write "<script language=javascript>"
	for i=1 to ubound(showarr) step 10
		Response.Write "parent.showmsg('"&showarr(i+1)&"','"&showarr(i+2)&"','"&showarr(i+3)&"','"&showarr(i+4)&"','"&showarr(i+5)&"','"&showarr(i+6)&"','"&showarr(i+7)&"','"&showarr(i+8)&"');"
	next
Response.Write "</script>"
end if
erase talkarr
erase showarr
if newtalkpoint>mytalkpoint then session("Ba_jxqy_usertalkpoint")=newtalkpoint
function attack(un,st,mg,gr)
	timetmp=now()
	rst.Open "select ta.��������,ta.��������+tu.���� as ����,ta.��Ч,ta.����˵��,tu.����,tu.����,tu.����,tu.protect from ���� ta inner join �û� tu ON tu.����=ta.���� where ta.����='"&un&"' and ta.��ʽ='"&mg&"'",conn
	if rst.EOF or rst.BOF then
		attack="<font color=FF0000>��������</font>##ҧ���гݵض�%%˵��'����ѧ����"&mg&",һ������������һ�ȸ��£�"
	else
		basemp=rst("��������")
		att=rst("����")
		especial=rst("��Ч")
		fmorcal=rst("����")
		atdeclaration=rst("����˵��")
		if isnull(especial) then especila="��"
		if rst("����")<basemp then
			attack="<font color=FF0000>��������</font>##����������"&mg&"����%%����ϧ������������㡣"
		elseif rst("����")<Application("Ba_jxqy_newuser") then
			attack="<font color=FF0000>��������</font>##�����޸���֮���������ȱ�����ˣ�"
		elseif rst("protect")>timetmp then
			attack="<font color=FF0000>��������</font>##�����ڴ��ڱ����׶Σ��޷�����%%��"
		else
			rst.Close
			rst.Open "select ����,����,����,����,�ؼ�,����,״̬,protect from �û� where ����='"&st&"'",conn
			if rst("����")<Application("Ba_jxqy_newuser") then
				attack="<font color=FF0000>��������</font>%%���Ǹ����ӣ�##����ô�����������أ�"
			elseif rst("protect")>timetmp then
				attack="<font color=FF0000>��������</font>%%���ڴ����ܱ����׶Σ�##���޷�ʹ��"&mg&"������"	
			else	
				resatt=att-rst("����")
				stespecial=rst("�ؼ�")
				condition=rst("״̬")
				tmorcal=rst("����")
				Maxattack=Application("Ba_jxqy_Maxattack")*gr\20
				randomize()
				if resatt<0 then 
					resatt=clng(rnd()*100)+1
				elseif resatt>Maxattack then 
					resatt=Maxattack-clng(rnd()*100)
				end if
				morcal=clng(fmorcal-tmorcal*0.01)
				if morcal>=2100000000 then 
					morcal=2100000000
				elseif morcal<=-2100000000 then
					morcal=-2100000000
				end if
				if resatt>rst("����") and condition<>"����" then
					especial="����"
					especialtime=1800
					conn.execute "update �û� set ����="&morcal&",����=����-"&basemp&",����=����+1 where ����='"&un&"'"
					conn.execute "insert into Ӣ����(ʱ��,����,����,����) values(now(),'"&st&"','"&un&"','"&mg&"')"
					attack="<font color=FF0000>��##������</font>"&atdeclaration&"������%%�����ˣ���"
				elseif 	resatt>rst("����") and condition="����" then
					especial="����"
					especialtime=1800
					conn.execute "update �û� set ����=����-1,����=����-"&basemp&",����=����+1 where ����='"&un&"'"
					attack="<font color=FF0000>��##������</font>##������%%����Ȼ��ʼ��ʬ�����¼�1����"
				else
					especialtime=resatt\100
					if especialtime<10 then 
						especial="��"
					elseif especialtime>180 then 
						especialtime=180
					end if	
					conn.Execute "update �û� set ����=����-"&basemp&",����=����+1 where ����='"&un&"'"
					conn.Execute "update �û� set ����=����-"&resatt&",����=����+1 where ����='"&st&"'"
					attack="<font color=FF0000>��##������</font>"&atdeclaration&"��ʹ%%��������"&resatt&"��<bgsound src='../mid/dw.wav' loop=1>"
				end if
				if instr(stespecial,";��"&especial&";")<>0 then especial="��"
				if especial<>"��" then	attack=attack&newaberrant(st,un,especial,especialtime)
			end if
		end if
	end if
	rst.Close
end function
function seeequip(un,st)
	rst.Open "select ����,���,�ȼ�,��ż from �û� where ����='"&st&"'",conn
	seeequip="<font color=FF0000>���鿴��</font>##�õ�����%%���鱨�����ɣ�"&rst("����")&"����ݣ�"&rst("���")&"����ż��"&rst("��ż")&"��"
	if rst("�ȼ�")<3 then
		seeequip=seeequip&"������������֡�"
	elseif rst("�ȼ�")<6 and  rst("�ȼ�")>=3 then
		seeequip=seeequip&"�������е��ʵ����"
	elseif rst("�ȼ�")<9 and  rst("�ȼ�")>=6 then
		seeequip=seeequip&"��������ͷ��С��"
	else
		seeequip=seeequip&"��������ò�Ҫ�ǡ�"
	end if
	rst.Close
end function
function cast(un,st,mg)
	tt=now()
	rst.Open "select tc.����,tc.��Ч,tu.����,tu.����,tu.protect from ��Ʒ tc inner join �û� tu on tc.������=tu.���� where tc.����='"&mg&"' and tc.������='"&un&"' and ����>0",conn
	if rst.EOF or rst.BOF then 
		cast="<font color=FF0000>��Ͷ����</font>##���㲢û��"&mg&"�������޷�Ͷ����"
	elseif rst("����")<Application("Ba_jxqy_newuser") then 
		cast="<font color=FF0000>��Ͷ����</font>##�㻹�Ǹ����ˣ����ǻ�����"
	elseif rst("protect")>tt then
		cast="<font color=FF0000>��Ͷ����</font>##�����ڴ��ڱ����׶�,�޷�����%%��"
	else
		fmorcal=rst("����")
		athp=rst("����")
		atespecial=rst("��Ч")
		rst.Close
		rst.Open "select ����,����,����,�ؼ�,״̬,protect from �û� where ����='"&st&"'",conn
		if rst("����")<Application("Ba_jxqy_newuser") then 
			cast="<font color=FF0000>��Ͷ����</font>%%�㻹�Ǹ����ˣ�����ô���������أ�"
		elseif rst("protect")>tt then
			cast="<font color=FF0000>��Ͷ����</font>%%���ڴ������������׶Σ�����##�޷���֮ʹ��"&mg
		else
			stespecial=rst("�ؼ�")
			stmoral=rst("����")
			condition=rst("״̬")
			sthp=rst("����")
			if isnull(atespecial) then atespecial="��"
			conn.execute "update �û� set ����=����-1 where ����='"&un&"'"
			conn.execute "update ��Ʒ set ����=����-1 where ����='"&mg&"' and ������='"&un&"'"
			if instr(stespecial,";��"&atespecial&";")<>0 then atespecial="��" 
			if condition="����" then
				atespecial="��"
				conn.execute "update �û� set ����=����-1 where ����='"&un&"'"
				cast="<font color=FF0000>��Ͷ����</font>%%ҵ��������##��Ȼ����֮Ͷ��"&mg&"��������˼�1<bgsound src='../mid/anqi.wav' loop=1>"
			elseif sthp>athp then
				conn.execute "update �û� set ����=����-"&athp&",����=����+1 where ����='"&st&"'"
				especialtime=athp
				cast="<font color=FF0000>��Ͷ����</font>##��%%Ͷ����"&mg&"��ʹ֮��������<bgsound src='../mid/anqi.wav' loop=1>"&athp
			else
				morcal=clng(fmorcal-tmorcal*0.01)
				if morcal>2100000000 then
					morcal=2100000000
				elseif morcal<-2100000000 then
					morcal=-2100000000
				end if	
				atespecial="����"
				especialtime=1800
				conn.execute "update �û� set ����="&morcal&" where ����='"&un&"'"
				conn.execute "insert into Ӣ����(ʱ��,����,����,����) values(now(),'"&st&"','"&un&"','"&mg&"')"
				cast="<font color=FF0000>��Ͷ����</font>##��%%Ͷ����"&mg&"����ֱ�����%%����"
			end if
			if atespecial<>"��" then cast=cast&newaberrant(st,un,atespecial,especialtime)
			Response.Write "<script language=javascript>parent.resfrm.location.href="&chr(34)&"miniweapon.asp"&chr(34)&";</script>" 
		end if	
	end if	
	rst.Close
end function
function sendmoney(un,st,mg)
	if not isnumeric(mg) then
		sendmoney="<font color=FF0000>����Ǯ��</font>##�����ȷ����Ǯ"&mg&"��%%��"
	else
		mg=clng(mg)	
		if mg<1 then
			sendmoney="<font color=FF0000>����Ǯ��</font>##�����ȷ����Ǯ"&mg&"��%%��"
		else	
			rst.Open "select * from �û� where ����='"&un&"' and ����>="&mg,conn
			if rst.EOF or rst.BOF then
				sendmoney="<font color=FF0000>����Ǯ��</font>%%Ц�Ŷ�##˵��'�����������⣬��������"&mg&"�������������Һ���'"
			else
			conn.execute "update �û� set ����=����+"&mg&" where ����='"&st&"'"
			conn.execute "update �û� set ����=����-"&mg&" where ����='"&un&"'"
			sendmoney="<font color=FF0000>����Ǯ��</font>##��"&mg&"�������͸���%%<bgsound src='../mid/thanks.wav' loop=1>"
			end if
		end if	
	end if		
end function
function alert(un,st,co,gr,sn,mg)
	if (co="�ٸ�" and gr>=Application("Ba_jxqy_kickguestright")) or (Application("Ba_jxqy_systemname"&sn)=co and gr>7) then
		alert="<font color=FF0000 size=3>�����桿%%,�������ؾ�����:"&mg&"</font>"
	else
		alert="<font color=FF0000>�����桿</font>%%,##���㷢��Ч����:"&mg
	end if
end function
function bulletin(un,st,co,gr,sn,mg)
	if (co="�ٸ�" and gr>=Application("Ba_jxqy_kickguestright")) or (Application("Ba_jxqy_systemname"&sn)=co and gr>7) then
		bulletin="<font color=FF0000 size=3>�����桿</font><table bgcolor=FFFF00 width=80% align=center border=3><tr><td bgcolor=00FF00 align=center>�� �� �� ��</td></td><tr><td>"&msg&"</td></tr></table>"
	else
		bulletin="<font color=FF0000>�����桿</font>##������Ȩ���湫��:"&mg
	end if
end function
function distributemoney(un,co,gr,sn)
	randomize()
	if co="�ٸ�" and gr>=Application("Ba_jxqy_unlockipright") then
		money=clng(rnd()*2000)+1000
		onlinelist=Application("Ba_jxqy_onlinelist")
		for i=1 to ubound(onlinelist) step 8
			conn.execute "update �û� set ����=����+"&money&" where ����='"&onlinelist(i)&"'"
		next
		erase onlinelist
		distributemoney="<font color=FF0000>����Ǯ��</font>##����ٸ��������֣�ÿ�˵õ�����<bgsound src='../mid/faqian.wav' loop=1>"&money
	else
		onlinenum=Application("Ba_jxqy_allonlinenum")
		everymoney=clng(rnd()*300)+200
		money=onlinenum*everymoney
		rst.Open "select * from �û� where ����='"&un&"' and ����>="&money,conn
		if rst.EOF or rst.BOF then
			distributemoney="<font color=FF0000>����Ǯ��</font>##�����Ǯ�������������������Ӵ���´ΰɣ��´�������ͺ��ˡ�"
		elseif onlinenum<10 then
			distributemoney="<font color=FF0000>����Ǯ��</font>##�����ѹ��û�м����ˣ����û���κ����壬�´ΰɣ��´�������ͺ��ˡ�"	
		else
			conn.execute "update �û� set ����=����-"&money&",����=����+"&onlinenum&" where ����='"&un&"'"
			onlinelist=Application("Ba_jxqy_onlinelist")
			for i=1 to ubound(onlinelist) step 8
				conn.execute "update �û� set ����=����+"&everymoney&" where ����='"&onlinelist(i)&"'"
			next
			erase onlinelist
			distributemoney="<font color=FF0000>����Ǯ��</font>##�����������ó�"&money&"��Ǯ�Ӵ�������ڳ�ÿ�˵õ�"&everymoney&"�����ӵ�����<bgsound src='../mid/faqian.wav' loop=1>"
		end if
		rst.Close 	
	end if
end function
function degree(un,st,co,mg)
	if mg="����" then
		degree="<font color=FF0000>����⡿</font>##,���޷����%%Ϊ����"
	else
		rst.Open "select * from �û� where ����='"&un&"' and ����='"&co&"' and ���='����'",conn
		if rst.EOF or rst.BOF then
			degree="<font color=FF0000>����⡿</font>##,�㲻�Ǳ������ţ��޷����%%"
		else
			rst.Close
			rst.Open "select * from �û� where ����='"&st&"' and ����='"&co&"' and ���<>'����'",conn
			if rst.EOF or rst.BOF then
				degree="<font color=FF0000>����⡿</font>%%���Ǳ��ɵ��ӣ�����##�޷���⣡"
			else
				conn.execute "update �û� set ���='"&mg&"' where  ����='"&st&"'"
				degree="<font color=FF0000>����⡿</font>##���%%Ϊ"&co&mg&"��"
			end if
		end if
		rst.Close
	end if		
end function
function dislodge(un,st,co)
	rst.Open "select * from �û� where ����='"&un&"' and ����='"&co&"' and ���='����'",conn
	if rst.EOF or rst.BOF then
		dislodge="<font color=FF0000>�������</font>##,�㲻�Ǳ������ţ��޷���%%���ɽ��"
	else
		rst.Close
		rst.Open "select * from �û� where ����='"&st&"' and ����='"&co&"' and ���<>'����'",conn
		if rst.EOF or rst.BOF then
			dislodge="<font color=FF0000>�������</font>%%���Ǳ��ɵ��ӣ�����##����Ȩ���ã�"
		else
			conn.execute "update �û� set ���='��',����='��' where  ����='"&st&"'"
			dislodge="<font color=FF0000>�������</font>##��%%����������"&co&"��"
		end if
	end if
	rst.Close
end function
function impart(un,st,mg)
	if not isnumeric(mg) then
		impart="<font color=FF0000>��������</font>##�����ȷ�Ǵ���"&mg&"��%%��"
	else
		mg=clng(mg)	
		if mg<100 then
			impart="<font color=FF0000>��������</font>##�����ȷ�Ǵ���"&mg&"��%%��Ҳ̫���˵����"
		else	
			rst.Open "select * from �û� where ����='"&un&"' and ����>="&mg,conn
			if rst.EOF or rst.BOF then
				impart="<font color=FF0000>��������</font>%%Ц�Ŷ�##˵��'��ʵ��ֻ�����Ҽ�����Ҳ��ܿ��ĵĲ�����ô"&mg&"�࣡'"
			else
				conn.execute "update �û� set ����=����+"&mg*0.8&",����=����+"&mg*0.01&" where ����='"&st&"'"
				conn.execute "update �û� set ����=����-"&mg&" where ����='"&un&"'"
				impart="<font color=FF0000>��������</font>##��"&clng(mg*0.8)&"���������ڸ���%%<bgsound src='../mid/xixing.wav' loop=1>"
			end if
		end if	
	end if
end function
function burglemp(un,st)
	randomize()
	odds=rnd()*10000
	if odds>1999 then
		rst.Open "select * from �û� where ����='"&un&"'",conn
		uhp=rst("����")	
		rst.Close
		if odds*0.2>uhp then odds=uhp
		conn.execute "update �û� set ����=����-"&odds*0.2&",����=����-"&odds*0.2&" where ����='"&un&"'"
		burglemp="<font color=FF0000>�����Ǵ󷨡�</font>##���ڹ��ڳ��϶�%%��ʩ���Ǵ󷨣���Υ��������������Ҵ�ö������ã������͵����½�"&clng(odds*0.2)&"��"&newaberrant(un,"���","���",120)
	else
		rst.Open "select * from �û� where ����='"&st&"'",conn
		smp=rst("����")
		rst.Close
		if odds>smp then odds=smp
		if odds>1900 then odds=smp*0.5
		conn.execute "update �û� set ����=����+"&odds&",����=����-"&odds*0.1&" where ����='"&un&"'"
		conn.execute "update �û� set ����=����-"&odds&" where ����='"&st&"'"
		burglemp="<font color=FF0000>�����Ǵ󷨡�</font>##͵͵ʹ���˽�������ɻ�����Ǵ󷨣�������%%"&clng(odds)&"������<bgsound src='../mid/xixing.wav' loop=1>"
	end if	
end function
function burglemoney(un,st)
	randomize()
	odds=rnd()*10000
	rst.Open "select * from �û� where ����='"&un&"'",conn
	uhp=rst("����")	
	rst.Close
	if odds>3999 then
		if odds*0.2>uhp then odds=uhp
		conn.execute "update �û� set ����=����-"&odds*0.2&",����=����-"&odds*0.2&" where ����='"&un&"'"
		burglemoney="<font color=FF0000>��͵�ԡ�</font>##���ڹ��ڳ��϶�%%ʵʩ�������٣���Υ��������������Ҵ���ˣ������͵����½�"&clng(odds*0.2)&"��"&newaberrant(un,"���","���",180)
	elseif odds>0 and odds<4000 then
		rst.Open "select * from ��Ʒ  where ������='"&st&"' and ����>0 order by �۸� desc",conn
		if rst.EOF or rst.BOF then
			odds=clng(odds*0.2)
			if odds>uhp then odds=uhp
			conn.execute "update �û� set ����=����-"&odds&",����=����-"&odds&" where ����='"&un&"'"
			burglemoney="<font color=FF0000>��͵�ԡ�</font>%%�������ǣ����޳��##��Ȼ���֮����͵�ԣ�����Ҵ�ı������ˣ������͵����½�"&odds&"��"&newaberrant(un,"���","����",180)
		else
			cname=rst("����")
			set rst2=server.CreateObject("adodb.recordset")
			rst2.Open "select * from ��Ʒ where ������='"&un&"' and ����='"&cname&"'",conn
			if rst2.EOF or rst2.BOF then
				conn.execute "insert into ��Ʒ(����,����,����,����,����,����,��Ч,����,�۸�,������,����,װ��) values('"&cname&"','"&rst("����")&"',"&rst("����")&","&rst("����")&","&rst("����")&","&rst("����")&",'"&rst("��Ч")&"',1,"&rst("�۸�")&",'"&un&"',false,false)"
			else
				conn.execute "update ��Ʒ set ����=����+1 where ������='"&un&"' and ����='"&cname&"'"
			end if
			conn.execute "update ��Ʒ set ����=����-1 where ������='"&st&"' and ����='"&cname&"'"
			rst2.Close
			set rst2=nothing
			burglemoney="<font color=FF0000>��͵�ԡ�</font>##͵͵ʹ�����ֿտյ��ֶΣ�͵����%%<bgsound src='../mid/xiaotou.wav' loop=1>����Я����"&cname
		end if
		rst.Close
	else
		rst.Open "select * from �û� where ����='"&st&"'",conn
		umoney=rst("����")	
		rst.Close
		if odds>umoney then odds=umoney
		if odds>1900 then odds=umoney*0.5
		conn.execute "update �û� set ����=����+"&odds&",����=����-"&odds*0.01&" where ����='"&un&"'"
		conn.execute "update �û� set ����=����-"&odds&" where ����='"&st&"'"
		burglemoney="<font color=FF0000>��͵�ԡ�</font>##͵͵ʹ�����ֿտյ��ֶΣ�͵����%%"&clng(odds)&"������<bgsound src='../mid/xiaotou.wav' loop=1>"
	end if	
end function
function present(un,st,mg)
	rst.Open "select * from ��Ʒ  where ������='"&un&"' and ����>0 and ����='"&mg&"'",conn
	if rst.EOF or rst.BOF then
		present="<font color=FF0000>�����͡�</font>##�������ǣ����޳����ȻҲ����"&mg&"�͸�%%��"
	else
		set rst2=server.CreateObject("adodb.recordset")
		rst2.Open "select * from ��Ʒ where ������='"&st&"' and ����='"&mg&"'",conn
		if rst2.EOF or rst2.BOF then
			conn.execute "insert into ��Ʒ(����,����,����,����,����,����,��Ч,����,�۸�,������,����,װ��) values('"&mg&"','"&rst("����")&"',"&rst("����")&","&rst("����")&","&rst("����")&","&rst("����")&",'"&rst("��Ч")&"',1,"&rst("�۸�")&",'"&st&"',false,false)"
		else
			conn.execute "update ��Ʒ set ����=����+1 where ������='"&st&"' and ����='"&mg&"'"
		end if
		conn.execute "update ��Ʒ set ����=����-1 where ������='"&un&"' and ����='"&mg&"'"
		rst2.Close
		set rst2=nothing
		present="<font color=FF0000>�����͡�</font>##��"&mg&"��󷽷��ĵݵ�%%��ǰ˵����һ�����⣬�Ʊ��ѿ�����"
	end if
	rst.Close
end function
function card(un,st,mg)
	if Application("Ba_jxqy_fellow")=True then
		rst.Open "select tc.* from card tc inner join �û� tu ON tu.����=tc.username where tc.username='"&un&"' and tc.cname='"&mg&"' and tc.cnumber>0 and tu.��Ա=true",conn
	else
		rst.Open "select * from card where username='"&un&"' and cname='"&mg&"' and cnumber>0",conn
	end if	
	if rst.EOF or rst.BOF then
		card="<font color=FF0000>�����ߡ�</font>##���㲢û��"&mg&"�ɶ�%%ʹ�ã�"
	else
		especial=trim(rst("cespecial"))
		especialtime=rst("ctime")
		esptime=dateadd("s",especialtime,now())
		rst.Close
		rst.Open "select �ؼ� from �û� where ����='"&st&"'"
		if rst.EOF or rst.bof then Response.End 
		if instr(rst("�ؼ�"),";��"&especial&";")<>0 then
			card="<font color=FF0000>�����ߡ�</font>##��%%ʹ����"&mg&",����%%���п�"&especial&"���ԣ����û���κ�Ч��"
		else
			card="<font color=FF0000>�����ߡ�</font>##��%%ʹ����"&mg&newaberrant(st,un,especial,especialtime)
		end if
		conn.execute "update card set cnumber=cnumber-1 where username='"&username&"' and cname='"&mg&"'"
		if instr(";����;����;����;",especial)<>0 then conn.execute "update �û� set ״̬='"&especial&"',����¼ʱ��='"&esptime&"' where ����='"&st&"'"
	end if
	rst.Close 
end function
function newaberrant(st,un,especial,especialtime)
	nowtime=now()
	aberrantlist=Application("Ba_jxqy_aberrantlist")
	aberrantlistubd=ubound(aberrantlist)
	dim newaberrantlist()
	newaberrantname=";"
	j=1
	addsucf=false
	for i=1 to aberrantlistubd step 4
		if addsucf=false and aberrantlist(i+2)=especial and aberrantlist(i)=st then
			redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
			newaberrantlist(j)=st
			newaberrantlist(j+1)=un
			newaberrantlist(j+2)=especial
			newaberrantlist(j+3)=dateadd("s",especialtime,aberrantlist(i+3))
			newaberrantname=newaberrantname&aberrantlist(i)&";"
			j=j+4
			addsucf=true
			if especial<>"����" then newaberrant=newaberrant&","&especial&"�̶ȼ���"&especialtime
		elseif datediff("s",aberrantlist(i+3),timetmp)<0 then
			redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
			newaberrantlist(j)=aberrantlist(i)
			newaberrantlist(j+1)=aberrantlist(i+1)
			newaberrantlist(j+2)=aberrantlist(i+2)
			newaberrantlist(j+3)=aberrantlist(i+3)
			newaberrantname=newaberrantname&aberrantlist(i)&";"
			j=j+4
		end if
	next
	if addsucf=false then
		redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
		newaberrantlist(j)=st
		newaberrantlist(j+1)=un
		newaberrantlist(j+2)=especial
		newaberrantlist(j+3)=dateadd("s",especialtime,nowtime)
		newaberrantname=newaberrantname&st&";"
		if especial<>"����" then newaberrant=newaberrant&","&especial&especialtime&"��"
	end if
	Application.Lock
	Application("Ba_jxqy_aberrantlist")=newaberrantlist
	Application("Ba_jxqy_aberrantname")=newaberrantname
	Application.UnLock
	erase newaberrantlist
	erase aberrantlist
	if especial="���" then
		newaberrant=newaberrant&"<bgsound src='../mid/dianxu.wav' loop=1>"
	elseif especial="����" then
		lastlogintime=dateadd("s",especialtime,nowtime)
		lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
		conn.execute "update �û� set ״̬='����',����¼ʱ��="&lastlogintimetype&" where ����='"&st&"'"
		newaberrant=newaberrant&"<bgsound src='../mid/oh_no.wav' loop=1>"
	end if
end function
%>
