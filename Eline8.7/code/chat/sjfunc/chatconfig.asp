<%
sub dbdz(bb,cz)	'���Ƿ��ڶᱦ��ս�����ж�
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
fjname=chatinfo(0)
if chatinfo(5)<>0 then
	erase sjjh_roominfo
	erase chatinfo
	Response.Write "<script language=JavaScript>{alert('��ʾ���ڡ�"&fjname&"�����䲻����ʹ��"&cz&"��');}</script>"
	Response.End
end if
if bb=1 then
	if chatinfo(0)="����E��" then
		erase sjjh_roominfo
		erase chatinfo
		Response.Write "<script language=JavaScript>{alert('��ʾ���ᱦ����ʱ�����µ���������½ǹ������еĹ��ܻ����ᱦ��������ʹ��"&cz&"��');}</script>"
		Response.End
	end if
end if
f=Minute(time())
if f<30 and fjname="E�߽���" then
	erase sjjh_roominfo
	erase chatinfo
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&fjname&"��ֻ����ÿСʱ�ĺ�30���Ӳſ���ʹ��"&cz&"�����ܵ�������E�ߡ�����ȥ�ɣ�"&fjname&"');}</script>"
	Response.End
end if
erase sjjh_roominfo
erase chatinfo
end sub

sub roompd(zsm)
	sjjh_roominfo=split(Application("sjjh_room"),";")
	chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
	fjname=chatinfo(0)
	erase chatinfo
	erase sjjh_roominfo
	if fjname<>"����E��" then
		Response.Write "<script language=JavaScript>{alert('��ʾ��"& zsm &"ʹ��ֻ���ڡ�����E�ߡ�����ſ��ԣ�');}</script>"
		Response.End
	end if
	s=Hour(time())
	f=Minute(time())
	m=Second(time())
	weekdate=weekday(date())
	sjz=weekdate*10000+s*100+f
	'����ֵ*10000+Сʱֵ*100+���ӣ���������19��50��Ϊ6*10000+19*100+50=61950����������20:00��Ϊ6*10000+20*100+0=62000
	if sjz<72030 then
		Response.Write "<script language=JavaScript>{alert('��ʾ���ᱦ������δ��ʼ�����Ժ�');}</script>"
		Response.End
	end if
end sub
function zxrsbd(slname,slid)	'���ᱦ�������������ж�
	zxrs=application("sjjh_onlinelist"&nowinroom)
	 '��ʼͳ��������Ա��Ŀ
	rjsq=ubound(zxrs)
	if rjsq<=1 then
		erase zxrs
		Set conn1=Server.CreateObject("ADODB.CONNECTION")
		Set rs1=Server.CreateObject("ADODB.RecordSet")
		conn1.open Application("sjjh_usermdb")
		conn1.execute "update �ᱦ set �ھ�="& slid &",ʱ��=now(),��ȡ=false,��������=0,����=true"
		set rs1=nothing
		conn1.close
		set conn1=nothing
		gg="<font color=blue>"& slname &"</font>��������ս�����ڶ����<font color=red><b>�ᱦ�����ھ�</b></font>��վ��"& application("sjjh_user") &"����������<font color=red><b>�Ͻ��«</b></font>������<font color=blue>"& slname &"</font><img src=sjfunc/guanjun.gif>���˱���Ҫ�������������ѵĸ������޼��õ���ҡ�"
		zxrsbd="<bgsound src=wav/gonggao.wav loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>�� �� �� ��</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
		exit function
	end if
	for i=1 to rjsq
		sfgf=split(zxrs(i),"|")
		if sfgf(0)<>slname and sfgf(2)<>"�ٸ�" then
			erase sfgf
			erase zxrs
			exit function
		end if
		erase sfgf
	next
	erase zxrs
	Set conn1=Server.CreateObject("ADODB.CONNECTION")
	Set rs1=Server.CreateObject("ADODB.RecordSet")
	conn1.open Application("sjjh_usermdb")
	conn1.execute "update �ᱦ set �ھ�="& slid &",ʱ��=now(),��ȡ=false,��������=0,����=true"
	set rs1=nothing
	conn1.close
	set conn1=nothing
	gg="<font color=blue>"& slname &"</font>��������ս�����ڶ����<font color=red><b>�ᱦ�����ھ�</b></font>��վ��"& application("sjjh_user") &"����������<font color=red><b>�Ͻ��«</b></font>������<font color=blue>"& slname &"</font><img src=sjfunc/guanjun.gif>���˱���Ҫ���������������ѵĸ������޼��õ�500����ҡ�"
	zxrsbd="<bgsound src=wav/gonggao.wav loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>�� �� �� ��</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
end function
function exitltzxpd(inroom) '�˳��������ж�
	sjjh_roominfo=split(Application("sjjh_room"),";")
	chatinfo=split(sjjh_roominfo(inroom),"|")
	fjname=chatinfo(0)
	erase sjjh_roominfo
	erase chatinfo
	if fjname<>"����E��" then
		exitltzxpd=""
		exit function
	end if
	session("nowinroom")=0
	Set conn1=Server.CreateObject("ADODB.CONNECTION")
	Set rs1=Server.CreateObject("ADODB.RecordSet")
	conn1.open Application("sjjh_usermdb")
	rs1.open "select * from �ᱦ",conn1,2,2
	xb=rs1("����")
	rs1.close
	if xb then
		set rs1=nothing
		conn1.close
		set conn1=nothing
		exitltzxpd=""
		exit function
	end if
	zxrs=application("sjjh_onlinelist"&inroom)
	 '��ʼͳ��������Ա��Ŀ
	rjsq=ubound(zxrs)
	jsq=0
	sname=""
	if rjsq<=1 then
		set rs1=nothing
		conn1.close
		set conn1=nothing
		exitltzxpd=""
		exit function
	end if
	for i=1 to rjsq
		sfgf=split(zxrs(i),"|")
		if sfgf(0)<>session("sjjh_name") and sfgf(2)<>"�ٸ�" then
			sname=sfgf(0)
			jsq=jsq+1
		end if
		if jsq>1 then
			erase zxrs
			set rs1=nothing
			conn1.close
			set conn1=nothing
			exitltzxpd=""
			exit function
		end if
		erase sfgf
	next
	erase zxrs
	if sname<>"" then
		rs1.open "select id from �û� where ����='"&sname&"'",conn1
		if not (rs1.eof or rs1.bof) then
			slid=rs1("id")
			rs1.close
			conn1.execute "update �ᱦ set �ھ�="& slid &",ʱ��=now(),��ȡ=false,��������=0,����=true"
		else
			rs1.close
			set rs1=nothing
			conn1.close
			set conn1=nothing
			exitltzxpd=""
			exit function
		end if
	else
		set rs1=nothing
		conn1.close
		set conn1=nothing
		exitltzxpd=""
		exit function
	end if
	set rs1=nothing
	conn1.close
	set conn1=nothing
	gg="<font color=blue>"& sname &"</font>��������ս�����ڶ����<font color=red><b>�ᱦ�����ھ�</b></font>��վ��"& application("sjjh_user") &"����������<font color=red><b>�Ͻ��«</b></font>������<font color=blue>"& sname &"</font><img src=sjfunc/guanjun.gif>���˱���Ҫ���������������ѵĸ������޼��õ�500����ҡ�"
	exitltzxpd="<bgsound src=wav/gonggao.wav loop=1><BR><BR><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>�� �� �� ��</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
end function
sub dbxx(xx)
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
	t="<font class=t>(" & sj & ")</font>"
	says="<font color=CC0000><b>���ᱦ��Ϣ��</b></font>" & xx			'��������
	says=replace(says,chr(39),"\'")
	says=replace(says,chr(34),"\"&chr(34))
	act="��Ϣ"
	towhoway=0
	towho="���"
	addwordcolor="660099"
	saycolor="008888"
	addsays="��"
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
	addmsg saystr
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
end sub
%>