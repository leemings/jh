<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
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
if day(date())=28 then
	Response.Write "<script Language=Javascript>alert('����������нˮ��ʱ�䣬����ô�������ﰡ��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
tmprs=conn.execute("Select count(*) As ���� from �û� where allvalue>100 and ������='"&aqjh_name&"'")
jsren=tmprs("����")
rs.open "Select * from �û� where ����='"&aqjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("�ȼ�")>=20 then
   rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('��ĵȼ��Ѿ�����20����,�����������?');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<15 and aqjh_grade<10 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('Ϊ��ֹ���ף�������15���ӵ�����ȡ������нˮ!');location.href = 'javascript:history.back()';}</script>"
		response.end

end if
if rs("��Ա�ȼ�")<>0 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ȼ���Ա�벻Ҫ��������!');location.href = 'javascript:history.back()';}</script>"
response.end
end if
if rs("��Ա")=true and clng(DateDiff("d",now(),rs("��Ա����")))>0  then
   rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('�ݷ��ƻ�Ա�벻Ҫ����������');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
if Session("aqjh_grade")>=5 then
 Response.Write "<script language=JavaScript>{alert('�ⲻ�����ź͹ٸ�Ӧ�����ĵط���');location.href = 'javascript:history.back()';}</script>"
 Response.End
end if
mdate=rs("��Ǯ")
'if CDate(mdate)=now() then
'if day(mdate)>=day(now()) and month(mdate)=month(now()) and year(mdate)=year(now()) then 
if  DateDiff("d",rs("��Ǯ"),now())=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ѽ��["&aqjh_name &"]�����������Ǯ�ģ������ˣ�');window.close();</script>"
	response.end
end if
yin=rs("����")+rs("���")
if yin>(rs("��Ա�ȼ�")+1)*500000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ѽ��["&aqjh_name &"]����������ô��������㻹��Ҫѽ����');window.close();</script>"
	response.end
end if
dj=rs("�ȼ�")
hy=rs("��Ա�ȼ�")
if trim(rs("����"))="����" then
	'��������ͳ������
	tmprs=conn.execute("Select count(*) As ���� from �û� where ����='"&rs("����")&"'")
	renshu=tmprs("����")
	'��Ǯ��
	money=cint(renshu)*100000+jsren*50000+dj*50000+500000
	conn.execute("update �û� set ����=����+"&money&",��Ǯ=now() where ����='"&aqjh_name&"'")
else
	'���˰���Ա��10��
	money=jsren*50000+dj*50000+500000+500000
	conn.execute("update �û� set ����=����+"&money&",��Ǯ=now() where ����='"&aqjh_name&"'")
end if
rs.close
rs.open "Select * from �û� where ����='"&aqjh_name&"'",conn
%>
<html>
<head>
<title>������ȡ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
BODY {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
SELECT {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#CCCCCC" text="#000000" leftmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="center">
<p><font color="#000000"><b><font size="+3">������ȡ��</font></b></font> <br>
<br>
<font size=+1> <b><%=aqjh_name%></b></font>����<font color="#0000FF">[<%=rs("����")%>]</font>
��<font color="#FF0000"> [<%=rs("����")%>] </font>���ˣ�<font color="#FF0000"><b><%=jsren%>��</b></font>    
ս���ȼ�:<font color="#FF0000"><%=rs("�ȼ�")%>��</font>    
<%if trim(rs("����"))="����" then%>    
���е���:<font color="#FF0000"><b><%=renshu%>��</b></font> <%end if%><br>���Ŭ����������ˣ�    
<br>
������ӽ�����ȡ����<b><font color="#FF0000"><%=money%>��</font></b>��С�ı��棬��Ҫ�һ���    
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
says="<font color=red>����ȡ���ʡ�["&aqjh_name&"]</font><font color=blue>���������<font color=red>"&money&"</font>���Ĺ���,��һ�����ȥ,��һ�¾�û����!</font>"
call showchat(says)
%>
<br>
</p>
<p align="center"><br>
<font color="#FF0000" size="+1"><b>���ʼ��㷽��</b></font><br>
<font color="#0000FF">���ӣ�</font>��������x5��+ս���ȼ�x5��+50��</p>
<p align="center">&nbsp;</p>
<p align="center">
<input type=button value=�رմ��� onClick='window.close()' name="button">
</p>
</div>
</body>
</html>