<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>��<%=Application("aqjh_chatroomname")%>��</title>
<link rel="stylesheet" href="../../css.css">
</head>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="BACK.GIF" oncontextmenu=self.event.returnValue=false>
<p align="center">[ <%=Application("aqjh_chatroomname")%><font color=blue> <%=aqjh_name%> </font>�ʻ� ]</p>
<%
rs.open "SELECT w7 FROM �û� where ����='"&aqjh_name&"'",conn,2,2
hua=rs("w7")
Response.Write "<div align='center'>"
if hua="" or isnull(hua) then
 Response.Write "<font size=+2 color=red>�Բ�����û���κ��ʻ���</font>"
else
data1=split(hua,";")
data2=UBound(data1)
rs.close
for y=0 to data2-1
	data3=split(data1(y),"|")
rs.open "select i FROM b WHERE a='" & data3(0) &"'",conn,3,3
%>
<img src='../jhjs/images/<%=rs("i")%>' alt="<%=data3(0)%>&nbsp;&nbsp;<%=data3(1)%>��"> 
<%x=x+1
if x/5=int(x/5) then Response.Write "<br>"
rs.close
next
end if
Response.Write "<br><br>"
Response.Write "���ͣҪ��������˵����<br><center>"
Response.Write "</div>"
set rs=nothing
conn.close
set conn=nothing
%><br><br>
<FONT color=#0000ff>&copy; ��Ȩ���� 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>���ֽ�����</FONT></A>