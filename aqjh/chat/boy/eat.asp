<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=LCase(trim(Request.QueryString("name")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w1,boy,boysex,��ż from �û� where ����='"&aqjh_name&"'",conn,2,2
w1=rs("w1")
if iswp(w1,"�׷�")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����û�׷�����ôι��ѽ��');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
zz=rs("��ż")

myhua=rs("w1")
if isnull(rs("boy")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û��С�����Ͽ�����һ���ɣ�');</script>"
	response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С�����ݳ����������¹���');</script>"
	response.end
end if

mypic=rs("boysex")
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
if zg>=(4+int(DateDiff("d",zt(2),now())/10)) then
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��չ�"&(4+int(DateDiff("d",zt(2),now())/10))&"���ˣ����5Сʱ�ٲ�����');</script>"
	response.end
end if
duyao=abate(myhua,name,1)
'����|���|����|����|����|����|�չ�����|�չ�ʱ��
'����|����|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))+50)&"|"&zt(4)&"|"&zt(5)&"|"&(zg+1)&"|"&now()
conn.execute "update �û� set w1='"&duyao&"',boy='"&temp&"' where  ����='"&aqjh_name&"'"
conn.execute "update �û� set boy='"&temp&"' where  ����='"&zz&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��С��ι��["&name&"]�ɹ���');parent.f3.location.href = 'boy.asp';</script>"
%>