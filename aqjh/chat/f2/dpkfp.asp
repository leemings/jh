<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'���˿���
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
says=dpkfp(fnn1,aqjh_name,towho)
says="<font color=green><b>�����ơ�</b><font color=" & saycolor & ">"&says&"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function dpkfp(fn1,from1,to1)
on error Resume next

if to1="���" and fn1<>"������" then call ErrALT("��ѡ��˵�����󣬲������ҷ����������!"+fn1)

aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"

nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
chatroomsn=session("nowinroom")

if left(fn1,2)="��֤" then
	Errstr="����������ʽ��\n��ȷʾ����\n/����$ ��֤|��\n/����$ ��֤|ʤ"
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then call ErrALT(Errstr)
	fn_62=left(arr_fn1(1),1)

	if fn_62<>"��" and fn_62<>"ʤ" then call ErrALT(Errstr)
	if grade<6 then call ErrALT("�㲻�ǹٸ���Ա�����ܽ����ƾֹ�֤��")
	from1=to1
end if

if chatroomsn<>"0" then
	dpkfp=ErrSays(nickname,"����֤�������䡻�����ڴ������в����ڴ�ҵ��ٶ�" )
	exit function
End if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dpk.asp"
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr

sql="select * from dpk where ufrom='"& from1 & "'"
Set Rs=conn.Execute(sql)

if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	call ErrALT(from1 & "û�в���[�˿�]�ƾ֣�")
end if

dpk=rs("pk")
dpkto=rs("uto")
xiazhu=rs("duz")
fpok=rs("fp")
oldoldpn=Cint(rs("oldpn"))
oldmaxpk=rs("maxpk")
logtime=rs("logtime")
iszai=rs("iszai")
rs.close
'dpkfp=cstr(oldoldpn)
'exit function

dim pk_geshi,old_pk_geshi
pk_geshi=1

if oldoldpn<10 then oldoldpn=90
old_pk_geshi=CINT(left(oldoldpn,1))
oldoldpn=CINT(mid(oldoldpn,2))

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
		conn.close
		set rs=nothing
		set conn=nothing
		call ErrALT("�Ƿ�������")
end select

'------------------------�¸�ʽ------------------------
d_sql="delete from dpk where ufrom='" & from1 & "' or uto='" & from1 &"'"

if fn_62<>"" then
	if fn_62="��" then
		s_pk=dpkto
		f_pk=to1
	elseif fn_62="ʤ" then
		f_pk=dpkto
		s_pk=to1
	End if
	Set Rs=conn.Execute(d_sql)
	conn.close
	connstr=Application("aqjh_usermdb")
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open connstr 

	'------------------------�¸�ʽ------------------------
	conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*2 where ����='"& s_pk &"'"
	conn.close
	dpkfp="���ٸ���֤��" & ErrSays(f_pk,"������[�˿��ƾ�]�涨�ܾ�����ƾֲ�������<br> �ڡڡڡ� �ٸ�[ " & nickname & " ]�ö������[ " & s_pk & " ]" & xiazhu & xia_class & sayxia & " �ۡۡۡ� ")
elseif fn1<>"��Ҫ��" and fn1<>"������" then
  if to1<>dpkto then
    dpkfp= ErrSays(nickname,"��Ķ��ֲ���[" & to1 & "]����")
    conn.close
    exit function
  end if
  if fpok then
    dpkfp= ErrSays(nickname,"���ں������㷢�ư�,�ȶԷ������˲����㷢�ƣ�")
    conn.close
    exit function
  end if
'olddpk22=split(dpk,"|")
'olddpk3=ubound(olddpk22)
dim qysall
dim qysn
dim pkx
dim isArrpk(15)
dim isArrpkn(15)

fn1=fn1 & "+"
fn1=replace(fn1,"++","+")
fn2=split(fn1,"+")
fn3=ubound(fn2)
for i=0 to fn3-1
  if instr(dpk,fn2(i) & "|")<>0 then
     oldpn=oldpn+1
     pkx=mid(fn2(i),2)
     if fn2(i)="����" or fn2(i)="С��" then pkx=fn2(i)
     pkx=unconvpk(pkx)
     isArrpkn(pkx)=isArrpkn(pkx) + 1
     isArrpk(pkx)=isArrpk(pkx) & " " & showpk(fn2(i))
     qys=left(fn2(i),1)
     qysn=qysn+1
     if qysall<>qys then qysall=qysall & qys
  end if
  dpk=replace(dpk,fn2(i) & "|","")
next
maxpk=0
dim zaid,sang,yidu,pkyizha2,pkyizha3,liandui
zaid=0
sang=0
yidu=0
pkyizha2=0
pkyizha3=0
pkyizha3=1
liandui=2

for i=15 to 1 step -1
  if isArrpkn(i)=4 then
    zaid=zaid+4
    fpnow =fpnow & " �ĸ�(ը)" & isArrpk(i)
    if maxpk=0 then 
	maxpk=i
	pk_geshi=4
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=3 then
    fpnow =fpnow & " ����" & isArrpk(i)
    sang=sang+3
    if maxpk=0 then 
	maxpk=i
	pk_geshi=3
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=2 then
     fpnow =fpnow & " һ��" & isArrpk(i)
     yidu=yidu+2
     if isArrpkn(i-1)=2 then liandui=liandui+2
     if maxpk=0 then 
	maxpk=i
	pk_geshi=2
     end if
  end if
next
for i=15 to 1 step -1
  if isArrpkn(i)=1 then
    pkyizha2=pkyizha2+1
    fpnow =fpnow & " һ��" & isArrpk(i)
    if isArrpkn(i-1)=1 then pkyizha3=pkyizha3+1
    if maxpk=0 then 
	maxpk=i
	pk_geshi=1
    end if
  end if
next
dpk22=split(dpk,"|")
dpk3=ubound(dpk22)
if len(qysall)=1 and qysn>1 then 
qysall="<b><font color=red>[��һɫ]</font></b>"
if maxpk=0 then 
	pk_geshi=pk_geshi+5
end if
else
qysall=""
end if

dim Errfp
Errfp=false
dim ErrStr
ErrStr=""
dim show_old_pk
show_old_pk=oldmaxpk

if oldmaxpk=11 then show_old_pk="J"
if oldmaxpk=12 then show_old_pk="Q"
if oldmaxpk=13 then show_old_pk="K"
if oldmaxpk=14 then show_old_pk="A"
if oldmaxpk=15 then show_old_pk="2"
if oldmaxpk=16 then show_old_pk="С��"
if oldmaxpk=17 then show_old_pk="����"

        IF old_pk_geshi=2 then
		if oldoldpn<3 then
			e_Str="һ��"
		else
			e_Str="����"
		end if
	elseif old_pk_geshi=3 then
		e_Str="����"
	elseif old_pk_geshi=4 then
		e_Str="ը��"
	elseif old_pk_geshi=1 then
		if oldoldpn>4 then
			e_Str="һ��������һ���ƣ�"
		else
			e_Str="һ��" 
		end if
	end if
	if (old_pk_geshi-5>0 and oldoldpn<>0) then e_Str=e_Str & "��һɫ" 

if zaid<>4 and oldoldpn<>0 then
        
	if old_pk_geshi<>pk_geshi then
        	ErrStr= ErrSays(nickname,to1 & "�ղų�����[" & e_Str & "<B>" & show_old_pk & "</B>]��")
	elseif oldpn<>oldoldpn then
		'erralt(oldpn)
		ErrStr= ErrSays(nickname,to1 & "�ղų�����[" & oldoldpn & "]�Ű���")
	end if

	if ErrStr<>"" then
	dpkfp=ErrStr & "�����û�������ƣ�����[ϴ��]����[�� ��]" 
	conn.close
	exit function
	end if
end if


if pkyizha3=oldpn and oldpn>4 then 
qysall=qysall & "<b><font color=red>[һ����]</font></b>"
IsyitiaoL=true
else
IsyitiaoL=false
end if
if zaid=0 and yidu=0 and sang=0 and IsyitiaoL=false then
isgood=false
else
isgood=true
end if


if (yidu>0 and pkyizha2>0) then Errfp=true 
if (zaid=0 and (yidu>2 and yidu<>liandui) ) then Errfp=true '�������ϲ�������
if (sang<>0 and yidu=0 and pkyizha2>2) then Errfp=true '��������һ�Ի���һ����������

if (sang<>0 and oldpn<>5 and oldpn<>3) then Errfp=true '������������ 
if (zaid<>0 and oldpn>7) then Errfp=true 'ը������7��

if (oldpn=3 and yidu<>0) then Errfp=true '����������һ��

if (pkyizha2>2 and sang<>0) then Errfp=true '������Ų���һ��
if (isgood=false and oldpn>1) then Errfp=true 'û��ըû�ж�û����������һ���������Ҳ���һ����
if Errfp=true then
dpkfp= ErrSays(nickname,"���뷢����û�й��ɵ����ƣ���ô����˴��˿�Ҳ���᣿���������ȣ�" )
conn.close
exit function
end if
if maxpk=14 then maxpk=16
if maxpk=15 then maxpk=17
'�µ�14.15������ԭ����1415����
if maxpk=1 then maxpk=14
if maxpk=2 then maxpk=15

'response.write maxpk
'response.end
if maxpk<=Cint(oldmaxpk) then
  if (zaid>0 and iszai) or zaid=0 then
     'ErrALT(e_Str)
     dpkfp= ErrSays(nickname,to1 & "�ղų�����[" & e_Str & "<B>" & show_old_pk & "</B>]����������û���Ʊȸղŵ��ƴ�����[ϴ��]����[�� ��]")
     conn.close
     exit function
   end if
end if

if dpk3>0 then
sql="update dpk set pk='" & dpk & "',fp=true,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",logtime='" & now() & "' where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",iszai=" & iszai & " where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
dpkfp= ErrSays(nickname,"�����룬�ó�[" & oldpn & "]����������˦����������ԭ����......" & ismore & "<br> " & qysall & fpnow & " , ����" & dpk3 & "����")
%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
elseif oldpn=18 then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
dpkfp= ErrSays(nickname,"�������װ�(��������Ϸ�������׹����е���)������Ҫ���㣬��[" & dpkto & "]����")
else

Set Rs=conn.Execute(d_sql)
conn.close
connstr=Application("aqjh_usermdb")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr 

'------------------------�¸�ʽ------------------------
conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"& xiazhu&"*2 where ����='"& nickname&"'" 
conn.close
dpkfp= ErrSays(nickname,"������Ц�������е��Ʒ��˳���......" & ismore & "<br> " & fpnow & " , û�����ˣ���") & ErrSays(dpkto,"Ӯ��" & xiazhu & xia_class & sayxia )
'------------------------�¸�ʽ------------------------

end if

elseif fn1="��Ҫ��" then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
dpkfp= ErrSays(nickname,"�����������ƣ���[" & dpkto & "]����")
conn.close
elseif fn1="������" then

Set Rs=conn.Execute(d_sql)
conn.close
connstr=Application("aqjh_usermdb")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr 


'------------------------�¸�ʽ------------------------
conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"&xiazhu&"*2 where ����='"& dpkto &"'"
conn.close
dpkfp= ErrSays(nickname,"�����˷����������,���[" & dpkto & "]" & xiazhu & xia_class & sayxia)
'------------------------�¸�ʽ------------------------

end if
set conn=nothing
set rs=nothing
end function

function showpk(pk)
dim wpk
wpk=replace(pk,"��","<img src=f2/dpk/1.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"��","<img src=f2/dpk/2.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"÷","<img src=f2/dpk/3.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"��","<img src=f2/dpk/4.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"С��","<img src=f2/dpk/xw.gif border=0>")
wpk=replace(wpk,"����","<img src=f2/dpk/dw.gif border=0>")
showpk=wpk & "</font>"
end function
function unconvpk(cpk)
dim pk2
pk2=replace(cpk,"A","1")
pk2=replace(pk2,"J","11")
pk2=replace(pk2,"Q","12")
pk2=replace(pk2,"K","13")
pk2=replace(pk2,"С��","14")
pk2=replace(pk2,"����","15")
unconvpk=pk2
end function
%>