<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'���齫
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
action=Request("action")
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
  says=bdsays(says)
end if
says=Ucase(says)
act=1
towhoway=0
i=instr(says,"$")
fnn1=trim(mid(says,i+1))
select case action
	case 1
		says=dmjfp(fnn1,aqjh_name,towho)
		says="<font color=green><b>�����ơ�</b><font color=" & saycolor & ">"&says&"</font>"
	case 2
		says=dmjGetmj(aqjh_name,towho)
		says="<font color=green><b>�����ơ�</b><font color=" & saycolor & ">"&says&"</font>"
	case 3
		says=dmjPeng("��",aqjh_name,towho)
		says="<font color=green><b>�����ơ�</b><font color=" & saycolor & ">"&says&"</font>"
	case 4
		says=dmjPeng(fnn1,aqjh_name,towho)
		says="<font color=green><b>�����ơ�</b><font color=" & saycolor & ">"&says&"</font>"
	case 5
		says=dmjPeng(fnn1,aqjh_name,towho)
		says="<font color=green><b>�����ơ�</b><font color=" & saycolor & ">"&says&"</font>"
	case 6
		says=dmjPeng(says,aqjh_name,towho)
		says="<font color=green><b>�����ơ�</b><font color=" & saycolor & ">"&says&"</font>"
	case else
		call ErrALT("��������")
end select

call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)


'���齫
function dmjfp(fn1,from1,to1)
if to1="���" and fn1<>"������" then call ErrALT("��ѡ��˵�����󣬲������ҷ����������!")
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
chatroomsn=session("nowinroom")

if left(fn1,2)="��֤" then
	Errstr="����������ʽ��\n��ȷʾ����\n/���� ��֤|��\n/���� ��֤|ʤ"
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then call ErrALT(Errstr)
	fn_62=left(arr_fn1(1),1)

	if fn_62<>"��" and fn_62<>"ʤ" then call ErrALT(Errstr)
	if grade<6 then call ErrALT("�㲻�ǹٸ�����Ա�����ܽ����ƾֹ�֤��" )
	from1=to1
end if

f=Minute(time()) 

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dmj.asp"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�����ķ�����֧��jet4.0����ʹ�����ص����ӷ�������߳�������
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& from1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrALT(from1 & "û�в���[�齫]�ƾ֣�")
end if
isMy=rs("isMy")
isFp=rs("isFp")
isGet=rs("isGet")
Mymj=rs("Mymj")
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
mjID=rs("mjID")
rs.close

'------------------------�¸�ʽ------------------------
dim xia_class,xia_fir,xia_sql
xia_fir=left(xiazhu,1)
xiazhu=mid(xiazhu,2)

select case xia_fir
	case "1"
		xia_class="���"
		xia_sql="���"
	case "2"
		xia_class="����"
		xia_sql="����"
	case "3"
		xia_class="����"
		xia_sql="�ݶ�����"
	case else
		call ErrALT("�Ƿ�������")
end select
'------------------------�¸�ʽ------------------------
d_sql="delete from dmj where ufrom='" & from1 & "' or uto='" & from1 &"'"
d_sql2="delete from mjInfo where id=" & mjID

if fn_62<>"" then
	if fn_62="��" then
		s_pk=dmjto
		f_pk=to1
	elseif fn_62="ʤ" then
		f_pk=dmjto
		s_pk=to1
	End if

	Set Rs=conn.Execute(d_sql)
	Set Rs=conn.Execute(d_sql2)
	conn.close
	connstr=Application("aqjh_usermdb")
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open connstr

	'------------------------�¸�ʽ------------------------
	conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*3 where ����='"& s_pk &"'"
	conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where ����='"& f_pk &"'"
	conn.close
	dmjfp="���ٸ���֤��" & ErrSays(f_pk,"������[�齫�ƾ�]�涨�ܾ�����ƾֲ�������<br> �ڡڡڡ� �ٸ���Ա[ " & nickname & " ]�ֲö������[ " & s_pk & " ]" & xiazhu & xia_class & sayxia & " �ۡۡۡ� ")
elseif fn1<>"������" then
		if to1<>dmjto then
			dmjfp= ErrSays(nickname,"��Ķ��ֲ���[" & to1 & "]����")
			conn.close
			exit function
		end if
		if not isMy then
			dmjfp= ErrSays(nickname,"���ڲ�������ƣ���ȴ��Է����ƣ�")
			conn.close
			set rs=nothing
			set conn=nothing
			exit function
		end if
		if isGet and (not isFp) then
			dmjfp= ErrSays(nickname,"��Ӧ�������ƣ�Ȼ���ٴ���,��[ϴ��]�����[�� ��]ʹ��(/����)����")
			%>
			<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/����";</script>
			<%
			conn.close
			set rs=nothing
			set conn=nothing
			exit function
		end if
		if instr(Mymj,fn1 & "|")<>0 then
			Mymj=replace(Mymj,fn1 & "|","",1,1,1)
		else
			dmjfp= ErrSays(nickname,"����û�и������û�������ư���")
			conn.close
			set rs=nothing
			set conn=nothing
			exit function
		end if
		sql="update dmj set Mymj='" & Mymj & "',oldmj='" & fn1 & "',isGet=true,isFp=false,isMy=false,logtime='" & now() & "' where ufrom='"& nickname & "'"
		Set Rs=conn.Execute(sql)
		sql="update dmj set isGet=true,isFp=false,isMy=true where ufrom='"& dmjto & "'"
		Set Rs=conn.Execute(sql)
		conn.close
		dmjfp= ErrSays(nickname,"�ó�һ��������ָͷ���˴�,��......������������ԭ���� " & showMj(fn1))%>
		<script language="JavaScript">parent.f3.location.reload();parent.f2.document.forms[0].sytemp.value="/����";</script>
		<%
elseif fn1="������" then
		Set Rs=conn.Execute(d_sql)
		Set Rs=conn.Execute(d_sql2)
		conn.close
		connstr=Application("aqjh_usermdb")
		Set conn=Server.CreateObject("ADODB.CONNECTION")
		conn.open connstr 

'------------------------�¸�ʽ------------------------
		conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*3 where ����='"& dmjto &"'"
		conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"& xiazhu&" where ����='"& nickname &"'"
		conn.close
		dmjfp= ErrSays(nickname,"���䲻���ˣ����[" & dmjto & "]" & xiazhu  & xia_class & sayxia )
'------------------------�¸�ʽ------------------------
end if
set conn=nothing
set rs=nothing
end function
'/////////////////////////////////////////////////ץ��
function dmjGetmj(from1,to1)
if to1="���" and fn1<>"������" then call ErrALT("��ѡ��˵�����󣬲������ҷ����������!")
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
chatroomsn=session("nowinroom")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dmj.asp"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�����ķ�����֧��jet4.0����ʹ�����ص����ӷ�������߳�������
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)

if rs.eof and rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrAlt("��û�вμӴ��齫����ô���뵽������?")
else
	isMy=rs("isMy")
	isFp=rs("isFp")
	isGet=rs("isGet")
	Mymj=rs("Mymj")
	dmjto=rs("uto")
	xiazhu=rs("duz")
	mjID=rs("mjID")
	logtime=rs("logtime")
	rs.close

if not isMy then
	dmjGetmj= ErrSays(nickname,"���ں����������ư�,�ȶԷ������˲��������ƣ�")
	conn.close
	set rs=nothing
	set conn=nothing
	exit function
end if
if isFp and (not isGet) then
	dmjGetmj= ErrSays(nickname,"������Ӧ������(����),��[ϴ��]������齫ʹ��(/���齫 XXX) ����")
	conn.close
	set rs=nothing
	set conn=nothing
	exit function
end if
sql="select �齫 from mjInfo where id="& mjID 
Set Rs=conn.Execute(sql)
Getmjs=rs("�齫")
rs.close

if len(Getmjs)>39 then
Getmjs2=right(Getmjs,3)
Mymj=Mymj & Getmjs2
Getmjs=left(Getmjs,len(Getmjs)-3)
sql="update mjInfo set �齫='" & Getmjs & "' where id="& mjID 
Set Rs=conn.Execute(sql)
sql="update dmj set Mymj='" & Mymj & "',isGet=false,isFp=true,logtime='" & now() & "' where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
%>
<script language="JavaScript">parent.f3.location.href="dmj-xp.asp";alert("������<%=strmj2(left(Getmjs2,2))%>")</script>
<%
dmjGetmj= ErrSays(nickname,"����������һ���ƣ�С�ĵĿ��˿������˵��ۣ���֪����ʲô......" )
else
sql="delete from dmj where ufrom='" & nickname & "' or ufrom='" & dmjto &"'"
Set Rs=conn.Execute(sql)
sql="delete from mjInfo where id=" & mjID
Set Rs=conn.Execute(sql)
dmjGetmj= ErrSays(nickname,"ֻ��13�����ˣ����������ƣ����Ǵ�ɺ;֣�û��ʤ����" )
end if
end if
End Function
'////////////////////////////////////////////////////////////////////����
function dmjPeng(fn1,from1,to1)

if to1="���" and fn1<>"������" then call ErrALT("��ѡ��˵�����󣬲������ҷ����������!")
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
chatroomsn=session("nowinroom")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dmj.asp"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�����ķ�����֧��jet4.0����ʹ�����ص����ӷ�������߳�������
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select oldmj from dmj where uto='"& nickname & "' and ufrom='"& to1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrAlt("��û�и�[ " & to1 & " ]���齫!")
end if
oldmj=rs(0)
	rs.close
if fn1="��" then
	dmjPeng=nickname & ":" & Errsays(to1,"�ղŴ���һ�� " & showMj(oldmj)) 
	conn.close
	set rs=nothing
	set conn=nothing
	Exit Function
end if
sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
isMy=rs("isMy")
isFp=rs("isFp")
isGet=rs("isGet")
Mymj=rs("Mymj")
Gangmj4=rs("����")
isGang=true
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
if to1<>dmjto then
dmjPeng= ErrSays(nickname,"��Ķ��ֲ���[" & to1 & "]����")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
if not isMy then
dmjPeng= ErrSays(nickname,"���ں������㷢�ư�,�ȶԷ������˲����㷢�ƣ�")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
fn1=fn1 & "+"
fn1=replace(fn1,"++","+")
fn2=split(fn1,"+")

fn3=ubound(fn2)
if isFp and (not isGet) and fn3<>4 then
dmjPeng= ErrSays(nickname,"���Ѿ�����һ���ƣ��������ٳԱ��˵��ơ�")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
dim Mymj3
Mymj3=Mymj

for i=0 to fn3-1
if instr(Mymj3,fn2(i) & "|")<>0 then
Mymj3=replace(Mymj3,fn2(i) & "|","",1,1,1)
else
dmjPeng= ErrSays(nickname,"����û�и����������Щ����")
rs.close
conn.close
set conn=nothing
set rs=nothing
exit function
end if
next
Rtp=right(oldmj,1)
Ltp=left(oldmj,1)
select case fn3
case 2
if oldmj=fn2(0) and fn2(0) =fn2(1) then
	Mymj=Mymj & oldmj & "|"
	dmjPeng= ErrSays(nickname,"�������� " & showMj(oldmj) & " �Զ�[" & to1 & "]�����ӵõ��������ƣ�������Ц���������.....�ǺǺ�")
elseif Rtp=right(fn2(0),1) and right(fn2(0),1)=right(fn2(1),1) then
	if (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp-2 & Rtp)<>0) or (instr(fn1,Ltp+1 & Rtp)<>0 and instr(fn1,Ltp+2 & Rtp)<>0) or (instr(fn1,Ltp-1 & Rtp)<>0 and instr(fn1,Ltp+1 & Rtp)<>0) then
		dmjPeng= ErrSays(nickname,"�������� " & showMj(fn2(0)) & showMj(fn2(1)) & " �Զ� [" & to1 & "] " & showMj(oldmj) & "��.....�õ���������")
		Mymj=Mymj & oldmj & "|"
	else
		dmjPeng= ErrSays(nickname,"���������������ô�Ա��˵��ư�?")
		Mymj=""
	end if
else
	dmjPeng= ErrSays(nickname,"���������������ô�Ա��˵��ư�?")
	Mymj=""
end if

case 3
if oldmj=fn2(0) and fn2(0) =fn2(1) and fn2(1)=fn2(2) then
	Mymj=Mymj & oldmj & "|"
	Gangmj4=Gangmj4 & oldmj & "|"
	dmjPeng= ErrSays(nickname,"�������� " & showMj(oldmj) & " �Զ�[" & to1 & "],������Ц���������.....[��]�ܣ�����")
	sql="update dmj set Mymj='" & Mymj & "',isGet=true,isFp=false,isMy=true,����='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& nickname & "'"
	Set Rs=conn.Execute(sql)
	dmjPeng=dmjPeng 
else
	dmjPeng= ErrSays(nickname,"��������������� �Ա��˵��ư�?")
end if
Mymj=""
case 4
if instr(Gangmj4,fn2(0) & "|")<>0 then isGang=false
if fn2(0) =fn2(1) and fn2(1)=fn2(2) and fn2(2) =fn2(3) then
	if isGang then
		if isGet and (not isFp) then
			dmjPeng= ErrSays(nickname,"��Ӧ�������ƣ�Ȼ����[��]����,��[ϴ��]�����[����������]ʹ��(/����)����")
			%>
			<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/����";</script>
			<%
			conn.close
			exit function
		end if

		Gangmj4=Gangmj4 & fn2(0) & "|"
		dmjPeng= ErrSays(nickname,"��������......" & showMj(fn2(0)) & "��������Ц���������.....[��]�ܣ�����")
		sql="update dmj set isGet=true,isFp=false,isMy=true,����='" & Gangmj4 & "',logtime='" & now() & "' where ufrom='"& nickname & "'"
		Set Rs=conn.Execute(sql)
		conn.close
		dmjPeng=dmjPeng 
	else
		dmjPeng= ErrSays(nickname,showMj(fn2(0)) & "���������Ѿ����˰�?����ܣ��������㣿")
	end if
else
	dmjPeng= ErrSays(nickname,"���������������ô�ܸܰ�?")
end if
Mymj=""

case else
dmjPeng= ErrSays(nickname,"����Ĳ������Ż����ţ�")
Mymj=""
end select
if Mymj<>"" then
sql="update dmj set Mymj='" & Mymj & "',isGet=false,isFp=true,isMy=true,logtime='" & now() & "' where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dmj set isGet=true,isFp=false,isMy=false where ufrom='"& dmjto & "'"
Set Rs=conn.Execute(sql)
conn.close
end if
%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
set conn=nothing
set rs=nothing
end function
function showMj(Mj)
showMj="<input type=image border=0 src=""f2/mj/" & intMjp(Mj) & ".gif"" onclick=""javascript:parent.f3.location.href='f2/dmj-xp.asp';"" >"
end function
function strmj2(Mj)
dim Mj2
Mj2=replace(Mj,"1","һ")
Mj2=replace(Mj2,"2","��")
Mj2=replace(Mj2,"3","��")
Mj2=replace(Mj2,"4","��")
Mj2=replace(Mj2,"5","��")
Mj2=replace(Mj2,"6","��")
Mj2=replace(Mj2,"7","��")
Mj2=replace(Mj2,"8","��")
strmj2=replace(Mj2,"9","��")
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL=left(cmj,1)
if isNumeric(mjL) then 
mj2=right(cmj,1) & mjL
mj2=replace(mj2,"��","1")
mj2=replace(mj2,"Ͳ","2")
mj2=replace(mj2,"��","3")
else
mj2=replace(mj2,"����","10")
mj2=replace(mj2,"�Ϸ�","20")
mj2=replace(mj2,"����","30")
mj2=replace(mj2,"����","40")
mj2=replace(mj2,"����","41")
mj2=replace(mj2,"�װ�","42")
mj2=replace(mj2,"����","43")
end if
intMjp=mj2
end function
%>
