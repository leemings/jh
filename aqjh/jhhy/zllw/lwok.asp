<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../pass.asp"-->
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
psw=trim(Request.form("psw"))
rz=trim(Request.Form("psw2"))
rzjm=trim(request.form("regjm"))
for each element in Request.Form
	if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('������ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
next
if len(psw)<5 then Response.Redirect "error.asp?id=52"
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
if Weekday(date())<>7 then
	Response.Write "<script Language=Javascript>alert('�Բ�����������ֻ������������ȡ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,w1,mvalue,�ȼ�,��Ա�ȼ� FROM �û� WHERE ����='"&session("aqjh_name")&"' and ����='"&psw&"'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ڻ��������ǽ������˰ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("mvalue")<=5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�»�����5000���ϵĲſ�����ȡ��Ʒ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("�ȼ�")<=60 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('Ϊ��ֹ���ף�ֻ�еȼ���60�����ϵĲſ�����ȡ��Ʒ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if hy=rs("��Ա�ȼ�")and hy<1 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�㲻�ǵȼ��ƻ�Ա����ذ�!');history.go(-1)';}</script>"
		response.end
end if
myid=rs("id")
wpdate=rs("w1")
hydj=rs("��Ա�ȼ�")
aqjh_name=session("aqjh_name")
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
	Response.Write "<script Language=Javascript>alert('���Ѿ������Ʒ�ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs1.close
lw=""
randomize timer
r=int(rnd*2)+1
select case r
	case 1 'ҩ����
		n=int(rnd*5)+6
		temp=add(wpdate,"ҩ����",100)
                sql="update �û� set w1='"&temp&"' where id="&myid
		lw="ҩ����100��"
	case 2 'ҩ��ʥˮ
		n=int(rnd*5)+6
		temp=add(wpdate,"ҩ��ʥˮ",100)
                sql="update �û� set w1='"&temp&"' where id="&myid
		lw="ҩ��ʥˮ100��"
	
end select
set rs=conn.execute(sql)
conn1.execute "insert into a(yhid,����,����) values ('"&myid&"','"&aqjh_name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
says="��<font color=#ff0000>��������</font>��<font color=blue>��ϲ<font color=red><b>��"& aqjh_name &"��</b></font>��վ������õ��ˡ�<font color=red><b>"& lw &"</b></font>������ҿ�ȥ�˵���<font color=red><b>���������</b></font>�������ﰡ��ȥ���˾�û���˰���<font color=red size=+1><b>ף��λ����ڿ��ֽ�����ÿ���!</b></font>"			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
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
set rs1=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FF9900" background="bg.gif">
<table width=401 height="227" border=3 align=center cellpadding="5" cellspacing="10" bordercolor="#6633CC">
  <tr bgcolor="#FFFFFF" align="center"> 
      <td width="367" height="252" bgcolor="#FF9900" background="bg.gif"><div align="center"><b>��ϲ���<font color="#FF0000"><%=application("aqjh_user")%></font>վ������õ���<br>
        <font color="#66FF00"><%=lw%></font></b><br>
<br>
<br>
<br>
<br>
[���ֽ���]</div></td>
  </tr>
</table>
</body>
</html>
