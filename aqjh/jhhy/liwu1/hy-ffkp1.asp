<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../pass.asp"-->
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
name=LCase(trim(Request.form("name")))
psw=trim(Request.form("psw"))
rz=trim(Request.Form("psw2"))
rzjm=trim(request.form("regjm"))
nowinroom=session("nowinroom")
for each element in Request.Form
	if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('������ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
next
name=trim(name)
name=server.HTMLEncode(name)
if name="��" or name="δ��" or name="" or name=null then response.redirect "error.asp?id=130"
if chuser(name) then Response.Redirect "error.asp?id=120"
if len(psw)<5 then Response.Redirect "error.asp?id=52"
for i=1 to len(name)
	usernamechr=mid(name,i,1)
	if asc(usernamechr)>0 then  Response.Redirect "error.asp?id=120"
next
if instr(name,"or")<>0 or instr(psw,"or")<>0 then Response.Redirect "error.asp?id=54"
if instr(name,"=")<>0 or instr(psw,"=")<>0 then Response.Redirect "error.asp?id=54"
if rz<>rzjm then Response.Redirect "error.asp?id=166"
if left(name,3)="%20" OR InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name,"��")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 or instr(name,chr(13)) then Response.Redirect "error.asp?id=120"
psw=md5(psw)
n=Year(date())
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
if sj<>"2015-07-01" then
	Response.Write "<script Language=Javascript>alert('�Բ��𣬽����ڼ�ŷ��ţ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,w5,�ȼ�,��Ա�ȼ� FROM �û� WHERE ����='" & name & "' and ����='"&psw&"'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ڻ��������ǽ������˰ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("�ȼ�")<=40 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('Ϊ��ֹ���ף�ֻ�еȼ���40�����ϵĲſ�����ȡ��Ʒ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
hy=rs("��Ա�ȼ�")
if hy<2 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�㲻�Ƕ�����Ա���ϸ��ѻ�Ա����ذ�!');history.go(-1)';}</script>"
		response.end
end if
myid=rs("id")
wpdate=rs("w5")
hydj=rs("��Ա�ȼ�")
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("sdlw.asp")
rs1.open "select * from a where yhid="&myid,conn1
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���Ѿ������Ƭ�ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs1.close
lw=""
randomize timer
r=int(rnd*2)+1
select case r
	case 1 '����
		n=hydj
		temp=add(wpdate,"����",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="����"&n&"��"
	case 2 '��Ǯ��
		n=hydj
		temp=add(wpdate,"��Ǯ��",n)
		sql="update �û� set w5='"&temp&"' where id="&myid
		lw="��Ǯ��"&n&"��"
       case 3 '����
		n=hydj
		temp=add(wpdate,"����",n)
		sql="update �û� set w5='"&temp&"' where id="&myide
		lw="����"&n&"��"	
	
end select
set rs=conn.execute(sql)
conn1.execute "insert into a(yhid,����,����) values ('"&myid&"','"&name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
conn1.close
set conn1=nothing
says="<font size=2 color=red>����ȡ�������"&name&"</font><font size=2 color=blue>��<font color=red>"&hydj&"</font>����Ա������ȥ����������ȡ��<font color=red>"&lw&"</font>����Ա�ȼ�Խ����ÿ�ƬԽ��!</font>"
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
%>
<html>
<head>
<title>������Ʒ������ί��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FF9900" background="bg.gif">
<table width=401 height="227" border=3 align=center cellpadding="5" cellspacing="10" bordercolor="#6633CC">
  <tr bgcolor="#FFFFFF" align="center"> 
      <td width="367" height="252" bgcolor="#FF9900" background="bg.gif"><div align="center"><b>��ϲ���<font color="#FF0000">���ֽ���</font>��վ����õ���<br>
        <font color="#66FF00"><%=lw%></font></b><br>
<br>
<br>
<br>
<br>
[���ֽ�����]</div></td>
  </tr>
</table>
</body>
</html>
