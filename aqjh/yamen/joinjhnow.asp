<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../const1.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("aqjh_disproxy")=1 then
if (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
end if
'��ֹվ���ύ����
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ֹվ���ύ���ݣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if Application("aqjh_closedoor")="1" then
Response.Write "<script Language=Javascript>alert('��ӭ���Ĺ��٣�\n���������������̫�ȶ�������վ��������ʱ�ر�����ע�Ṧ��\n���ڲ��ܽ���ע�ᣬ���Щʱ��������\n���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ����\n��л���Ա�������֧�ֺͺ񰮣�\n���������Ĳ��㣬���б�Ǹ��\nSEE YOU\n\n2004-4-7');window.parent.close();</script>"
Response.End
end if
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
name=LCase(trim(Request.form("name")))
sex=Request.form("sex")
country=Request.Form("country")
psw=trim(Request.form("psw"))
pswc=trim(Request.Form("pswc"))
psw2=trim(Request.Form("psw2"))
oicq=trim(Request.form("oicq"))
sex=Request.form("sex")
e_mail=LCase(Request.form("e_mail"))
jsr=trim(request.form("jsr"))
jsr1=jsr
face=request("face")
face="../ico/"&face&"-2.gif"

if not isnumeric(oicq) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ���������oicq��ʹ�����֣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
name=trim(name)
name=server.HTMLEncode(name)
if jsr=name then response.redirect "../error.asp?id=57"
if name="��" or name="δ��" then response.redirect "../error.asp?id=130"
if chuser(name) then Response.Redirect "../error.asp?id=120"
if chuser(jsr) then Response.Redirect "../error.asp?id=60"
if len(oicq)<4 or len(oicq)>=10 then Response.Redirect "../error.asp?id=50"
if instr(e_mail,"@")=0 then Response.Redirect "../error.asp?id=51"



if len(pswc)<5 or len(psw2)<5 then Response.Redirect "../error.asp?id=52"
for i=1 to len(name)
usernamechr=mid(name,i,1)
if asc(usernamechr)>0 then  Response.Redirect "../error.asp?id=120"
next
for i=1 to len(jsr)
usernamechr=mid(jsr,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=60"
next


 if instr(name,"or")<>0 or instr(sex,"or")<>0 or  instr(psw,"or")<>0  or instr(pswc,"or")<>0 or instr(email,"or")<>0 or instr(oicq,"or")<>0 or instr(ask,"or")<>0 or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if instr(name,"=")<>0 or instr(sex,"=")<>0  or  instr(psw,"=")<>0 or instr(pswc,"=")<>0 or instr(email,"=")<>0 or instr(oicq,"=")<>0 or instr(ask,"=")<>0  or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if Instr(name,"��������")>0 or Instr(name,"����")>0 or Instr(name,"��������")>0 or Instr(name,Application("aqjh_automanname"))>0 or Instr(name,"˫������")>0 or Instr(name,"վ��")>0 or Instr(name,"����")>0 or Instr(name,"ʱ��")>0 or Instr(name,"��")>0 or Instr(name,"��")>0 or Instr(name,"��")>0 or Instr(name,"���")>0 or Instr(name,"��")>0 then Response.Redirect "../error.asp?id=130"
if pswc<>psw then Response.Redirect "../error.asp?id=166"
if trim(request.form("Name"))="" or trim(request.form("psw"))="" or trim(Request.Form("e_mail"))="" or trim(request.form("oicq"))="" or trim(request.form("psw2"))="" then Response.Redirect "../error.asp?id=56"
if trim(request.form("Name"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=129"
if trim(request.form("psw2"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=485"
if left(name,3)="%20" OR InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name,"��")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 then Response.Redirect "../error.asp?id=120"
if left(jsr,3)="%20" OR InStr(jsr,"=")<>0 or InStr(jsr,"`")<>0 or InStr(jsr,"'")<>0 or InStr(jsr," ")<>0 or InStr(jsr,"��")<>0 or InStr(jsr,"'")<>0 or InStr(jsr,chr(34))<>0 or InStr(jsr,"\")<>0 or InStr(jsr,",")<>0 or InStr(jsr,"<")<>0 or InStr(jsr,">")<>0 then Response.Redirect "../error.asp?id=120"
if len(jsr)=0 then
	jsr="��"
end if
psw=md5(psw)
psw2=md5(psw2)
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
userip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� WHERE regip=" & SqlStr(userip) & " and DateDiff('s',regtime,#" & sj & "#)<=10",conn
'If not(Rs.Bof OR Rs.Eof) Then
'			rs.close
'			set rs=nothing
'			conn.close
'			set conn=nothing
'			Response.Redirect "../error.asp?id=67"
'			response.end
'end if
rs.close
rs.open "SELECT * FROM �û� WHERE ����='" & name & "'",conn
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=62"
			response.end
end if
rs.close
rs.open "SELECT n���� from npc WHERE n����='" & name & "'",conn,1,1
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=484"
			response.end
end if
rs.close



if jsr1<>"" then
rs.open "SELECT * FROM �û� WHERE ����='" & jsr & "'",conn
If Rs.Eof Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=58"
			response.end
end if
rs.close
end if
if sex="��" then
	touxiang="../jhimg/boy00.gif"
else
	touxiang="../jhimg/girl00.gif"
end if
rs.open "select top 1 * FROM �û� WHERE DateDiff('d',lasttime,now())>12000000 and grade<5",conn,1,3
if rs.eof or rs.bof then
rs.AddNew
rs("����")=name
rs("�Ա�")=sex
rs("����")=10
rs("����ͷ��")=face
rs("ͷ��")=touxiang
rs("����")=psw
rs("�ڶ�����")=psw2
rs("����")="����"
rs("���ɻ���")=0
rs("���")="����"
rs("ְҵ")="�ɱ�"
rs("ʦ��")="��"
rs("ʦ����Ǯ")="��"
rs("��¼")=now()
rs("����")=e_mail
rs("����")=country
rs("ְλ")="����"
rs("oicq")=oicq
rs("����")=reg_ml
rs("����")=reg_dd
rs("ӮǮ")=0
rs("������")=0
rs("�Ĵ���")=0
rs("Ӯ����")=0
rs("�书")=reg_wg
rs("�书��")=0
rs("����")=reg_nl
rs("������")=0
rs("����")=reg_tl
rs("������")=0
rs("����")=reg_gj
rs("����")=reg_fy
rs("֪��")=reg_zz
rs("����")=reg_fl
rs("������")=0
rs("�Ṧ")=reg_qg
rs("�Ṧ��")=0
rs("����")=reg_yl
rs("״̬")="����"
rs("grade")=1
rs("�ȼ�")=reg_dj
rs("times")=1
rs("regtime")=now()
rs("regip")=userip
rs("lasttime")=now()
rs("lastip")=userip
rs("allvalue")=reg_dj*reg_dj*50
rs("mvalue")=0
rs("��Ǯ")=now()
rs("����")="����"
rs("����1")="����"
rs("����2")="����"
rs("���")=reg_ck
rs("��������")=date()
rs("����")="δ֪"
rs("ǩ��")="www.happyjh.com-���ֽ�����"
rs("��Ա�ȼ�")=reg_hydj
rs("��Ա����")=date()+reg_hydate
rs("��Ա��")=reg_hyjk
rs("ɱ����")=0
rs("����")="��"
rs("��������")=0
rs("������")=jsr
rs("�ݶ�����")=0
rs("����ʱ��")=now()
rs("����ʱ��")=now()
rs("��������")="|"&Application("aqjh_user")&"|"
rs("������")=0
rs("��������")=date()
rs("����")="��"
rs("��ż")="��"
rs("�¼�ԭ��")="��"
rs("����")=true
rs("ͨ��")=false
rs("ɱ����")=0
rs("��ɱ��")=0
rs("����ʱ��")=now()-1
rs("���")=reg_jb
rs("��Ա")=false
rs("��Ա����")=date()
rs("��")=0
rs("ľ")=0
rs("ˮ")=0
rs("��")=0
rs("��")=0
rs("���")=0
rs("ľ��")=0
rs("ˮ��")=0
rs("���")=0
rs("����")=0
rs("sl")="��"
rs("slsj")=now()
rs("cw")=""
rs("w1")="��"
rs("w2")="��"
rs("w3")="��"
rs("w4")="��"
rs("w5")=reg_hycard
rs("w6")="��"
rs("w7")="��"
rs("w8")="��"
rs("w9")="��"
rs("w10")="��"
rs("z1")=" "
rs("z2")=" "
rs("z3")=" "
rs("z4")=" "
rs("z5")=" "
rs("z6")=" "
rs("ת��")=0
rs("����")=0
rs("������")=0
rs("������")=0
rs("����")=0
rs("��")=0
rs("����")=0
rs("����")=0
rs("����")=0
rs("ħ��")=0
rs("����")=0
rs("�ƽ�")=0
rs("����")="��"
rs("����")="��ѧ����"
rs.Update
else
  Response.Redirect "../error.asp?id=67"
  response.end
end if
rs.close
rs.open "SELECT * FROM �û� WHERE ����='" & name & "'",conn
id=rs("id")
xinren=rs("����")
sex=rs("�Ա�")
touxiang=rs("����ͷ��")
jsr=rs("������")
regsj=rs("regtime")
rs.close
set rs=nothing
Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
End Function
says="<bgsound src=wav/Global.wav loop=1><img src=../yamen/xinr.gif><font color=red>�����˼��롽���[</font><font color=#cc0000>"& country &"</font><font color=#000000>]��</font><font color=green>����</font><img src="& touxiang &"><font color=#000000>[</font><font color=#cc0000>"& xinren &"</font>("& sex &")<font color=#000000>]</font><font color=green>ID</font><font color=#000000>[</font><font color=#cc0000>"& id &"</font><font color=#000000>]</font><font color=green>���뽭������Ҷ�����Ҫ���չ˰���ע��QQ:</font><font color=#000000>[</font><font color=#cc0000>"& oicq &"</font><font color=#000000>]</font><font color=green>������</font><font color=#000000>[</font><font color=#cc0000>"& jsr &"</font><font color=#000000>]</font>(<font color=#ff00ff>" & regsj & "</font>)"			'��������
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu1(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
Response.Write "<script Language=Javascript>alert('��ϲ{" & name & "}�����Ѿ��ɹ����뱾������������½��\n������Ϣ���μǣ���½��Ӧ��ʱ�������뱣����\n���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ����\n��ID��" & id & "\n������" & name & "\n���룺" & pswc & "\n���ң�" &country& "\n�������������δ��½\n����ID�������û�ȡ����');window.parent.close();</script>"
Response.End
%>