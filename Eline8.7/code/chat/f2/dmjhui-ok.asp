<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 

to1=LCase(trim(request.querystring("to1")))
nickname=Session("sjjh_name")
grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")

onlyDei2=intMjp(Request.querystring("onlyDei"))
if onlyDei2="" then onlyDei2="-1"
onlyDei=CINT(onlyDei2)

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="eDMJ.ASP"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�����ķ�����֧��jet4.0����ʹ�����ص����ӷ�������߳�������
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)

if rs.eof or rs.bof then 
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	call ErrALT("û��������ƾֻ��ѽ�����")
end if
Mymj=rs("Mymj")
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
mjID=rs("mjID")
rs.close

if to1<>dmjto then
	conn.close
	set conn=nothing
	set rs=nothing
	call ErrALT("��Ķ��ֲ���[" & to1 & "]����")
end if

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

dim oldqys
dim mjqys
mjqys=true
fnarr=split(Mymj,"|")
oldqys=right(fnarr(0),1)
for i=0 to ubound(fnarr)-1
	if oldqys<>right(fnarr(i),1) then mjqys=false
	fnint=fnint & intMjp(fnarr(i)) & "|"
next

dim ARRmj(43)
for j= 1 to 4
for i=10 to 43
if instr(fnint & "|",i & "|")<>0 then 
fnint=replace(fnint & "|",i & "|","",1,1,1)
ARRmj(i)=ARRmj(i)+1
end if
next
next


for i=10 to 43
if ARRmj(i)=4 then
	ARRmj(i)=0
	mjhui4=mjhui4+4
	fpnow =fpnow & " �����ӡ���" & strMjp(i) & strMjp(i) & strMjp(i) & strMjp(i) 
	zaid=true
end if
next

for i=10 to 43
if ARRmj(i)=3 then
	if i<>onlyDei then
		ARRmj(i)=0
		mjhui3=mjhui3+3
		fpnow =fpnow & " ��һ˳����" & strMjp(i) & strMjp(i) & strMjp(i) 
	else
		ARRmj(i)=1
		fpnow =fpnow & " ��һ�ԡ���" & strMjp(i) & strMjp(i) 
		mjhui2=mjhui2+2
	end if
end if
next

for i=10 to 43
if ARRmj(i)=2 and (i=onlyDei or onlyDei=-1) then
ARRmj(i)=0
	if i<41 then
		if ARRmj(i+1)=2 and ARRmj(i+2)=2 then
			ARRmj(i+1)=ARRmj(i+1)-2
			ARRmj(i+2)=ARRmj(i+2)-2
			fpnow =fpnow & " ��һ��˳����" & strMjp(i) & strMjp(i+1) & strMjp(i+2)
			fpnow =fpnow & " ��һ��˳����" & strMjp(i) & strMjp(i+1) & strMjp(i+2) 
			mjhui3=mjhui3+6
			mjhui2_3=mjhui2_3+6
			i=i+2
		else
			mjhui2=mjhui2+2
			fpnow =fpnow & "  ��һ�ԡ���" & strMjp(i) & strMjp(i) 
		end if
	else
		mjhui2=mjhui2+2
		fpnow =fpnow & " ��һ�ԡ���" & strMjp(i) & strMjp(i) 
	end if
end if
next

for i=10 to 43
if ARRmj(i)=2 then
ARRmj(i)=1
	if i<40 then
		if i<38 and ARRmj(i+1) >0 and ARRmj(i+2) >0 and i<>10 and i<> 20 and i<>30 and i+1<>10 and i+1<> 20 and i+1<>30 and i+2<>10 and i+2<> 20 and i+2<>30 then
			ARRmj(i+1)=ARRmj(i+1)-1
			ARRmj(i+2)=ARRmj(i+2)-1
			mjhui1=mjhui1+1
			fpnow =fpnow & " ��һ˳����" & strMjp(i) & strMjp(i+1) & strMjp(i+2) 
			i=i+2
		else
			mjhui0=mjhui0+1
			fpnow =fpnow & " ��һ�š���" & strMjp(i) 
		end if
	else
		fpnow =fpnow & " ��һ�š���" & strMjp(i) 
		mjhui0=mjhui0+1
	end if
end if
next


for i=10 to 43
if ARRmj(i)=1 then
	if i<40 then
		if i<38 and ARRmj(i+1) mod 3=1 and ARRmj(i+2) mod 3=1 and i<>10 and i<> 20 and i<>30 and i+1<>10 and i+1<> 20 and i+1<>30 and i+2<>10 and i+2<> 20 and i+2<>30 then
			mjhui1=mjhui1+1
			fpnow =fpnow & " ��һ˳����" & strMjp(i) & strMjp(i+1) & strMjp(i+2) 
			i=i+2
		else
			mjhui0=mjhui0+1
			fpnow =fpnow & " ��һ�š���" & strMjp(i) 
		end if
	else
		fpnow =fpnow & " ��һ�š���" & strMjp(i) 
		mjhui0=mjhui0+1
	end if
end if
next

mjhui2_3=mjhui2_3+mjhui2
'mod by hcs
'mjhui0=0
if mjhui0>0 then
	msg=fpnow & ErrALT2("����û�и������������Ҳ�к�������" & mjhui0 & "������"  ) 
elseif mjhui2>2 and mjhui2_3<>14 then
	msg=fpnow & ErrALT2("����û�и������������Ҳ�к��������û���߶Ծ����ֻ����һ������" ) 
elseif mjhui2<1 then
	msg=fpnow & ErrALT2("����û�и������������Ҳ�к�����һ������Ҳû�а�" ) 
else
'mod by hcs
'mjqys=true
newxiazhu=xiazhu

if mjqys then
	fpnow="����һɫ��" & fpnow
	newxiazhu=xiazhu * 2
elseif mjhui2_3=14 then
		fpnow="���߶��ӡ�" & fpnow
		'newxiazhu=xiazhu + (xiazhu/2)
		newxiazhu=newxiazhu * 2
elseif mjhui4>0 then
		fpnow="���ܡ�" & fpnow
		newxiazhu=xiazhu + (xiazhu/2)
end if

sql="delete from dmj where ufrom='" & nickname & "' or ufrom='" & dmjto &"'"
Set Rs=conn.Execute(sql)
sql="delete from mjInfo where id=" & mjID
Set Rs=conn.Execute(sql)
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")

'�ָ��������ʱ����ȥ��ע���䷽ֻ�ָ�ʣ�µģ�Ӯ���ָ�ȫ��������Ӯ��
conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+("& xiazhu &"*2-"& newxiazhu &") where ����='"& dmjto &"'"
conn.execute "update �û� set " & xia_sql & "=" & xia_sql & "+"& xiazhu &"*2+"& newxiazhu &" where ����='"& nickname &"'"
conn.close
set conn=nothing
duidu=ErrSays(nickname,"������Ц̯��,���ˣ���" & xia_class & "��" & xia_class & "...<br>" & fpnow & "<br>") & ErrSays(nickname,"��") & ErrSays(to1,"Ӯ��"& newxiazhu & xia_class & sayxia)
'------------------------�¸�ʽ------------------------
msg=nickname & "������Ц̯��,���ˣ���Ǯ��Ǯ...<br>" & fpnow & "<br>" & nickname & "��" & to1 & "Ӯ��"& newxiazhu & xia_class & sayxia
'msg=replace(msg,"f2/mj/","mj/")

duidu=duidu & "<script Language='JavaScript'>if(parent.myn=='" & nickname & "'||parent.myn=='" & to1 & "')parent.f4.showname();</script>"

says="<font color=red>�����齫��</font>��"&msg			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho=dmjto
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

End if

function convJS(Jss)
Jss=Replace(Jss,"\","\\")
Jss=Replace(Jss,"/","\/")
convJS=Replace(Jss,chr(34),"\"&chr(34))
end function
function strMjp(cmj)
strMjp = "<input type=image border=0 src=""f2/mj/" & cMj & ".gif"" >"
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

<body background="../bg1.gif" text="#FFFFFF" >
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="4" bordercolorlight="000000" bordercolordark="FFFFff">
  <tr> 
<td height="115"> 
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="4">
        <tr> 
          <td width="100%" height="29" align="center" style="FONT-SIZE: 9pt;cursor:hand;"><br>
            <%=replace(msg,"f2/mj/","mj/") %></td>
        </tr>
        <tr> 
          <td align="center" valign="top" height="58">
<input type=button value='���� �ա�' onClick="javascript:window.close()" style="height:20;font-size:9pt;background-color:#003300;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button223"> 
          </td>
        </tr>
      </table>
</td>
</tr>
</table>
