<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then 
	Session.Abandon
	Response.Write "<script language=JavaScript>{alert('��ʾ����ֹʹ�ô���ע�ᣡ');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
if Application("sjjh_closedoor")="1" then
Response.Write "<script Language=Javascript>alert('��ӭ���Ĺ��٣�\n���������������̫�ȶ�������վ��������ʱ�ر�����ע�Ṧ��\n���ڲ��ܽ���ע�ᣬ���Щʱ��������\n���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ����\n��л���Ա�������֧�ֺͺ񰮣�\n���������Ĳ��㣬���б�Ǹ��\nSEE YOU\n\n2003-11-10');window.parent.close();</script>"
Response.End
end if
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yusu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yusu=129 or yusu>192 or (yusu<2 and yusu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yusu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
'���³����ǽ�ֹ����,���Է�ֹ���ͽ�������������
if (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
name=server.HTMLEncode(trim(Request.form("name")))
sex=Request.form("sex")
zz=Request.form("zz")
psw=trim(Request.form("psw"))
pswc=trim(Request.Form("pswc"))
ask=trim(Request.form("ask"))
answer=trim(Request.form("answer"))
e_mail=LCase(Request.form("e_mail"))
oicq=trim(Request.form("oicq"))
jsr=trim(request.form("jsr"))
face=request("face")
if Request.form("face")=00 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������û��ѡ��ͷ���أ���������������ҵ�����Ŷ��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if not isnumeric(oicq) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&oicq&"]�������OICQ��ʹ�����֣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');</script>"
		Response.End
end if
next
name=trim(name)
name=server.HTMLEncode(name)
if jsr=name then response.redirect "../error.asp?id=57"
if name="��" or name="δ��" then response.redirect "../error.asp?id=130"
if chuser(name) then Response.Redirect "../error.asp?id=120"
if chuser(jsr) then Response.Redirect "../error.asp?id=60"
if len(oicq)<4 or len(oicq)>=15 then Response.Redirect "../error.asp?id=50"
if instr(e_mail,"@")=0 then Response.Redirect "../error.asp?id=51"
if len(pswc)<5 then Response.Redirect "../error.asp?id=52"
for i=1 to len(name)
usernamechr=mid(name,i,1)
if asc(usernamechr)>0 then  Response.Redirect "../error.asp?id=120"
next
for i=1 to len(jsr)
usernamechr=mid(jsr,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=60"
next
if instr(name,"or")<>0 or instr(sex,"or")<>0 or instr(psw,"or")<>0 or instr(pswc,"or")<>0 or instr(email,"or")<>0 or instr(oicq,"or")<>0 or instr(ask,"or")<>0 or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if instr(name,"=")<>0 or instr(sex,"=")<>0 or instr(psw,"=")<>0 or instr(pswc,"=")<>0 or instr(email,"=")<>0 or instr(oicq,"=")<>0 or instr(ask,"=")<>0  or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if Instr(name,"��ϯ")>0 or Instr(name,"��������")>0 or Instr(name,"����")>0 or Instr(name,"��������")>0 or Instr(name,Application("sjjh_automanname"))>0 or Instr(name,"���׵���")>0 or Instr(name,"��Ȼ")>0 or Instr(name,"վ��")>0 or Instr(name,"����")>0 or Instr(name,"ʱ��")>0 or Instr(name,"����")>0 or Instr(name,"�й�")>0 or Instr(name,"��")>0 or Instr(name,"��")>0 or Instr(name,"��")>0 or Instr(name,"���")>0 or Instr(name,"��")>0 or Instr(name,"�ڿ�")>0 or Instr(name,"�ٿ�")>0 then Response.Redirect "../error.asp?id=130"
if pswc<>psw then Response.Redirect "../error.asp?id=166"
if trim(request.form("Name"))="" or trim(request.form("psw"))="" or trim(Request.Form("e_mail"))="" or trim(request.form("oicq"))="" then Response.Redirect "../error.asp?id=56"
if trim(request.form("Name"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=129"
if left(name,3)="%20" OR InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name,"��")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 then Response.Redirect "../error.asp?id=120"
if left(jsr,3)="%20" OR InStr(jsr,"=")<>0 or InStr(jsr,"`")<>0 or InStr(jsr,"'")<>0 or InStr(jsr," ")<>0 or InStr(jsr,"��")<>0 or InStr(jsr,"'")<>0 or InStr(jsr,chr(34))<>0 or InStr(jsr,"\")<>0 or InStr(jsr,",")<>0 or InStr(jsr,"<")<>0 or InStr(jsr,">")<>0 then Response.Redirect "../error.asp?id=120"
if len(jsr)=0 then jsr="��"
psw=md5(psw)
zaohui=md5(ask&answer)
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
if y=2 And r>28 Then
	r=28
Else
	if r=31 Then	
		r=30 
	End if
end if
if y<10 then
huiqi=n & "-" & y+3 & "-" & r
else
huiqi=n+1 & "-" & y+3-12 & "-" & r
end if
userip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� WHERE regip=" & SqlStr(userip) & " and DateDiff('s',regtime,#" & sj & "#)<=150",conn,1,1
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=67"
			response.end
end if
rs.close
rs.open "SELECT * FROM �û� WHERE ����='" & name & "'",conn,1,1
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "../error.asp?id=62"
			response.end
end if
rs.close
if sex="��" then
	touxiang="../jhimg/boy00.gif"
else
	touxiang="../jhimg/girl00.gif"
end if
rs.open "select top 1 * FROM �û� WHERE DateDiff('d',lasttime,now())>90 and grade<5",conn,1,2
if rs.eof or rs.bof then
rs.AddNew
end if 
rs("����")=name
rs("�Ա�")=sex
rs("����")=10
rs("����ͷ��")=face
rs("ͷ��")=touxiang
rs("����")=psw
rs("����")="����"
rs("���ɻ���")=0
rs("���")="����"
rs("ְҵ")="�ɱ�"
rs("ʦ��")="��"
rs("ʦ����Ǯ")="��"
rs("��¼")=now()
rs("����")=e_mail
rs("oicq")=oicq
rs("����")=300
rs("����")=300
rs("ӮǮ")=0
rs("������")=0
rs("�Ĵ���")=0
rs("Ӯ����")=0
rs("�书")=200
rs("�书��")=0
rs("����")=200
rs("������")=0
rs("����")=35260
rs("������")=0
rs("����")=0
rs("����")=0
rs("֪��")=100
rs("����")=1000000
rs("״̬")="����"
rs("grade")=1
rs("�ȼ�")=10
rs("times")=1
rs("regtime")=now()
rs("regip")=userip
rs("lasttime")=now()
rs("lastip")=userip
rs("allvalue")=5000
rs("mvalue")=0
rs("��Ǯ")=now()
rs("����")="����"
rs("����1")="����"
rs("����2")="����"
rs("���")=10000
rs("��������")=date()
rs("����")="δ֪"
rs("ǩ��")="WWW.happyjh.com"
rs("��Ա�ȼ�")=1
rs("��Ա����")=huiqi
rs("��Ա��")=0
rs("����")="��"
rs("��������")=0
rs("������")=jsr
rs("�ݶ�����")=0
rs("����ʱ��")=now()-1
rs("����ʱ��")=now()
rs("��������")="|���׵���|"
rs("������")=0
rs("��������")=date()
rs("����")="��"
rs("��ż")="��"
rs("�¼�ԭ��")="��"
rs("����")=true
rs("ͨ��")=false
rs("ɱ����")=0
rs("��ɱ��")=0
rs("����ʱ��")=now()
rs("���")=0
rs("��Ա")=false
rs("��Ա����")=date()
rs("��ֵ��")=0
rs("����")=0
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
rs("��������")="��"
rs("��������")="��"
rs("sl")="��"
rs("slsj")=now()
rs("cw")=""
rs("w1")="a"
rs("w2")="a"
rs("w3")="a"
rs("w4")="a"
rs("w5")="a"
rs("w6")="a"
rs("w7")="a"
rs("w8")="a"
rs("w9")="a"
rs("w10")="a"
rs("z1")=""
rs("z2")=""
rs("z3")=""
rs("z4")=""
rs("z5")=""
rs("z6")=""
rs("hua")=""
rs.Update
rs.close
rs.open "SELECT id,����,�Ա�,����ͷ��,������,regtime FROM �û� WHERE ����='" & name & "'",conn,1,1
id=rs("id")
xinren=rs("����")
sex=rs("�Ա�")
touxiang=rs("����ͷ��")
jsr=rs("������")
regsj=rs("regtime")
rs.close
set rs=nothing
conn.close
set conn=nothing
Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
End Function
says="<bgsound src=wav/Global.wav loop=1><img src=../yamen/xinr.gif><font color=red>�����˼��롽</font><font color=green>����</font><img src=../ico/"& touxiang &"-2.gif width=32 height=32><font color=#000000>[</font><font color=#cc0000>"& xinren &"</font>("& sex &")<font color=#000000>]</font><font color=green>ID</font><font color=#000000>[</font><font color=#cc0000>"& id &"</font><font color=#000000>]</font><font color=green>�����˿��ֽ�����վ����Ҷ�����Ҫ���չ˰���</font><font color=green>������</font><font color=#000000>[</font><font color=#cc0000>"& jsr &"</font><font color=#000000>]</font>(<font color=#ff00ff>" & regsj & "</font>)"			'��������
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('��ϲ{" & name & "}�����Ѿ��ɹ����뱾������������½��\n������Ϣ���μǣ���½��Ӧ��ʱ�������뱣����\n���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ����\n��ID��" & id & "\n������" & name & "\n���룺" & pswc & "\n�������������δ��½\n����ID�������û�ȡ����');window.parent.close();</script>"
Response.End
%>
nse.End
%>