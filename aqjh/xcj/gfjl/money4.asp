<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
nowinroom=session("nowinroom")
id=LCase(trim(request.querystring("id")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
n=Year(date())
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
rs.open "SELECT * FROM �û� where ����='"&aqjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<5 and aqjh_grade<10 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('Ϊ��ֹ���ף�������5���ӵ�����ȡ������нˮ!');location.href = 'javascript:history.back()';}</script>"
		response.end

end if
if rs("����")<>"" then
                rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��û��Ϊ������̳��������!����!');location.href = 'javascript:history.back()';}</script>"
		response.end
end if
if rs("mvalue")<100000 then
                rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('���»���С��10�򣬲����콱��!');location.href = 'javascript:history.back()';}</script>"
		response.end
end if
if aqjh_jhdj>30 then Response.Write "<script language=JavaScript>{alert('��ȼ�̫���ˣ��첻�ˣ����´ΰɣ�');location.href = 'javascript:history.back()';}</script>"
pai=rs("����")
rs.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
tmprs=conn.execute("Select count(*) As ���� from �û� where mvalue>200000 and ������='"&aqjh_name&"'")
jsren=tmprs("����")
rs.open "Select * from �û� where ����='"&aqjh_name&"'",conn
mdate=rs("�ֶ�1")
'if CDate(mdate)=now() then
'if day(mdate)>=day(now()) and month(mdate)=month(now()) and year(mdate)=year(now()) then 
if  DateDiff("d",rs("�ֶ�1"),now())=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ѽ��["&aqjh_name &"]����������������ģ������ˣ�');window.close();</script>"
	response.end
end if
yin=rs("��Ա��")
ying=rs("���")
if yin>40000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ѽ��["&aqjh_name &"]��������4000��Ľ��㻹��Ҫ������');window.close();</script>"
	response.end
end if
if ying>40000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ѽ��["&aqjh_name &"]��������4000��Ľ���㻹��Ҫ������');window.close();</script>"
	response.end
end if
dj=rs("�ȼ�")
zs=rs("ת��")
gl=rs("grade")
if trim(rs("����"))<>"����" then
	'���˰���Ա��10��
	money=(5)*2
	conn.execute("update �û� set ���=���+"&money&",�ֶ�1=now() where ����='"&aqjh_name&"'")
end if
%>
<html>
<head>
<title>������ȡ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "����"; FONT-SIZE: 16pt}
BODY {FONT-FAMILY: "����"; FONT-SIZE: 16pt}
SELECT {FONT-FAMILY: "����"; FONT-SIZE: 16pt}
A {COLOR: #FFC106; FONT-FAMILY: "����"; FONT-SIZE: 16pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "����"; FONT-SIZE: 16pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#CCCCCC" text="#000000" leftmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="center">
<p><font size="2"><font color="#000000"><b>���ֽ�����ȡ��</b></font> <br>
<br>
������ӿ�����ȡ������<b><font color="#FF0000"><%=money%>�����</font></b>��С�ı��棬��Ҫ�һ���        
<% 
rs.close 
conn.close 
set rs=nothing 
set conn=nothing 
fn1="<font size=2 color=red>�������䷢����"&aqjh_name&"��</font><font size=2 color=blue>֧�ִ���������̳��չ��������Ĭ��վ������<font color=red>"&money&"</font>��<font color=brown>���</font>������δ����������»���õ�Ŷ~~</font>" 
says=nuhou(fn1) 
function nuhou(fn1) 
nuhou="<marquee height=80 behavior=alternate loop=100 direction=left >" & fn1 & "" & "</marquee>" 
end function 
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
<br> 
</font> 
</p> 
<p align="center"><font size="2"><br>
<font color="#FF0000" size="+1"><b>������ȡ˵������</b></font><br>
<br>
<font color="#0000FF">��̳���׽���ÿ�������ȡ������ֻ���»��ִﵽ10��Ĳſ�����ȡ��</font></font></p>
<p align="center"><font color="#0000FF" size="2">�޻�Ա��</font></p>
<p align="center"><b><font color="#FF00FF" size="2">�޻�Ա����ȡ��</font></b></p>
<p align="center">&nbsp;</p>
<p align="center">
<input type=button value=�رմ��� onClick='window.close()' name="button">
</p>
</div>
</body>
</html>