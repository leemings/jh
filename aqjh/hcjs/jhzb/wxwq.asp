<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
lb=lcase(trim(request("lb")))
if lb="" or (lb<>"z1" and lb<>"z2" and lb<>"z3" and lb<>"z4" and lb<>"z5" and lb<>"z6") then
	Response.Write "<script Language=Javascript>alert('��Ʒ��������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if InStr(lb,"or")<>0 or InStr(lb,"=")<>0 or InStr(lb,"`")<>0 or InStr(lb,"'")<>0 or InStr(lb," ")<>0 or InStr(lb,"��")<>0 or InStr(lb,"'")<>0 or InStr(lb,chr(34))<>0 or InStr(lb,"\")<>0 or InStr(lb,",")<>0 or InStr(lb,"<")<>0 or InStr(lb,">")<>0 then Response.Redirect "../../error.asp?id=54"
select case lb
	case z1
		lbmc="ͷ��"
	case z2
		lbmc="װ��"
	case z3
		lbmc="����"
	case z4
		lbmc="����"
	case z5
		lbmc="����"
	case z6
		lbmc="˫��"
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT "&lb&",w6,���� FROM �û� WHERE ����='"&aqjh_name&"'",conn
nnn=trim(rs(lb))
xj=rs("����")
yywpsl=trim(rs("w6"))
rs.close
if yywpsl="" or yywpsl=null then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ʲô��ƷҲû�У�����ά��װ������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(nnn)="" or trim(nnn)="��" or trim(nnn)=Null then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���"&lbmc&"��û��װ������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
data111=split(nnn,"|")
if ubound(data111)<>5 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���"&lbmc&"װ���������󣡣�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if data111(5)>=3 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���"&data111(0)&"��ά�޹�3���ˣ��޷��޸�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.open "select * from b where a='"&data111(0)&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��Ʒ���ݿ����޷��ҵ�["&data111(0)&"]��װ�����ݳ�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
wqlx=rs("b")
nai=rs("j")
newnai=data111(1)
ztz=int((newnai/nai)*100)
rs.close
if ztz>=95 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���["&data111(0)&"]״̬���ã�����Ҫά�ޣ���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if ztz<15 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���["&data111(0)&"]�ѱ��ϣ��޷��޸�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if data111(5)=1 and ztz>=75 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����������һ�Σ����ֻ�ܴﵽԭ�;öȵ�75��������Ҫ������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if data111(5)=2 and ztz>=55 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��������������Σ����ֻ�ܴﵽԭ�;öȵ�55��������Ҫ������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
select case data111(5)
	case 0
		data111(1)=nai
	case 1
		data111(1)=int(nai*0.75)
	case 2
		data111(1)=int(nai*0.55)
end select
yin=(data111(1)-newnai)*10000
if xj<yin then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��û��["&yin&"]�����ӣ�û�˿ϸ�����װ������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
liang=int(((data111(1)-newnai)/10)+1)
xhwpmc=""
if wqlx="����" then
	rs.open "select * from x where a='"&data111(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('�����������ݿ����Ҳ�����Ҫά�޵���������');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
	sxwp=rs("b")
	rs.close
	xdata1=split(sxwp,"|")
	xdata2=UBound(xdata1)
	for ii=0 to xdata2-1
		xadata2=split(xdata1(ii),":")
		myxlwp=trim(xadata2(0))
		yywpsl=abate(yywpsl,myxlwp,liang)
		xhwpmc=xhwpmc&"["&xadata2(0)&"]"&liang&"��  "
	next
	xhwpmc=xhwpmc&","
else
	yywpsl=abate(yywpsl,"����",liang)
	yywpsl=abate(yywpsl,"��ˮ",liang)
	yywpsl=abate(yywpsl,"��ʯ",liang)
	yywpsl=abate(yywpsl,"����",liang)
	yywpsl=abate(yywpsl,"ľͷ",liang)
	xhwpmc="[����]"&liang&"��  [��ˮ]"&liang&"��  [��ʯ]"&liang&"��  [����]"&liang&"��  [ľͷ]"&liang&"��,"
end if
data111(5)=clng(data111(5)+1)
wqnewzt=data111(0)&"|"&data111(1)&"|"&data111(2)&"|"&data111(3)&"|"&data111(4)&"|"&data111(5)

conn.execute "update �û� set "&lb&"='"&wqnewzt&"',w6='"&yywpsl&"',����=����-"&yin&" where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>װ��ά��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="back.gif">
<p align="center"><font color="#FF0000" size="7">װ��ά��</font></p>
<p align="center">װ�����ƣ�<%=data111(0)%> �;öȣ�<font color="#408080"><%=data111(1)%></font>/<font color="#FF0000"><%=nai%></font>  
  �������<font color="#FF0000"><%=data111(5)%></font>��ά�޴�װ����</p> 
<p align="center">������<%=xhwpmc%>����������<%=yin%>����<%=data111(0)%>���;öȻָ���<%=data111(1)%>��<br>
  <br>
  ˵����ÿ��װ��ֻ��ά�����Σ����κ��װ�����޷��޸���</p>
<p align="center"><a href="index.asp">����</a></p>
<p align="center"><font size="2">��Ȩ���С����齭����վ��</font></p>
</body>
</html>
