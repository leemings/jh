<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="mywp.asp"-->
<%Response.Expires=0
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
rs.open "select boy,boysex,w1 from �û� where ����='"&aqjh_name&"'",conn,3,3
mypic=rs("boysex")
w1=rs("w1")
if iswp(w1,"�׷�")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����û�׷�����ôι��ѽ��');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if

if isnull(rs("boy")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û��С���أ��Ͽ���һ���ɣ�');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С�����ݳ����������¹���');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
rs.close


zg=clng(zt(6))
'if zg>=4 then
'	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��չ�4���ˣ�\n���5Сʱ�ٲ�����');location.href = 'javascript:history.go(-1)'</script>"
'	response.end
'end if
%>
<html>
<head>
<title>ι��С��</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
body{
CURSOR: url('boy.cur');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#006699" bgproperties="fixed" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" background="../<%=chatimage%>">
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td width="100%" height="28"> <div align="center"><font color="#ffffff" size="2">ι��С��</font></div></td>
  </tr>
</table>
<table background="../../bg2.gif" border="1" cellspacing="0" width="140" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr> 
    <td align="center">
<img src="<%=mypic%>" width=40 height=40><br><br>
<%
'data1=split(myhua,";")
'data2=UBound(data1)


mh=mywpsl(w1,"�׷�")
Response.Write "<div align='center'><font color=#ff0000 size=-1>����׷�ι��!</font><br><br></div>"
Response.Write "<table border='1' cellspacing='0' cellpadding='0' width='50' bordercolorlight='#000000' bordercolordark='#FFFFFF' bgcolor='#CCCCCC' align='center' >"
Response.Write "<tr>"

		Response.Write "<td width='50' ><div align='center'><a href='eat.asp?name=�׷�' title='������!' target='d'><img src='images/mh.gif' width='50' height='50' border=0 alt='�׷����У�"& mh & "��!'></a></div></td>"

Response.Write "</tr></table>"
set rs=nothing
conn.close
set conn=nothing
%>  
<br>
</td>
  </tr>
</table>
<br>
<div align="center">&nbsp;&nbsp;<font color="#FFFFFF"><font size="2"><a href=javascript:history.go(-1)><font color=#ffffff size=-1>�����ϼ�</font></a><br>
  <br>
</body>
</html>