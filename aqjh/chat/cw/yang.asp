<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select cw,w7 from �û� where ����='"&aqjh_name&"'",conn,3,3
if isnull(rs("w7")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���㻹û���ʻ�������ι��!');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
myhua=rs("w7")
if isnull(rs("cw")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û�г����ȥ����');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ��������¹���');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
rs.close
rs.open "select i from b where a='"&zt(1)&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ��������¹���');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
mypic=rs("i")
rs.close
zg=clng(zt(6))
'if zg>=4 then
'	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��չ�4���ˣ�\n���5Сʱ�ٲ�����');location.href = 'javascript:history.go(-1)'</script>"
'	response.end
'end if
%>
<html>
<head>
<title>ι������</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" leftmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<p style="font-size:12pt; color:red" align="center"> <font color="#FFFFFF"><b>ι �� �� ��</b><br>
  </font> <img src="../../hcjs/jhjs/images/<%=mypic%>"></p>
<%
data1=split(myhua,";")
data2=UBound(data1)
Response.Write "<div align='center'><font color=yellow size=-1>����ʻ�ι��!</font><br><br></div>"
Response.Write "<table border='1' cellspacing='0' cellpadding='0' width='145' bordercolorlight='#000000' bordercolordark='#FFFFFF' bgcolor='#CCCCCC' align='center' >"
Response.Write "<tr>"
for y=0 to data2-1
	data3=split(data1(y),"|")
	rs.open "select i FROM b WHERE a='" & data3(0) &"'",conn,3,3
	If not(Rs.Bof and Rs.Eof) then
		Response.Write "<td width='80' ><div align='center'><a href='eat.asp?name="&data3(0)&"' title='������!' target='d'><img src='../../hcjs/jhjs/images/" & rs("i") & "' border=0 alt='"&data3(0)&"���У�"& data3(1)& "��!'></a></div></td>"
		x=x+1
		if x/2=int(x/2) then Response.Write "</tr><tr>"
	end if
	rs.close
next
Response.Write "</tr></table>"
set rs=nothing
conn.close
set conn=nothing
%>  
<br>
<br>
<br>
<div align="center">
<a href=javascript:history.go(-1)><font color=blue size=-1>�����ϼ�</font></a>
</div>
</body>
</html>