<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name=trim(request.querystring("name"))
for each element in Request.QueryString
if instr(Request.QueryString(element),"'")<>0 or instr(Request.QueryString(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ�����������⣬��鿴��');</script>"
		Response.End
end if
next



Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")


   pddb="pd.mdb"
   Set pdconn = Server.CreateObject("ADODB.Connection")
   Set rspd=Server.CreateObject("ADODB.RecordSet")
   pdconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(pddb)
   pdconn.open pdconnstr
   sqlpd="select * from p where a='"&name&"'"
   rspd.open sqlpd,pdconn,1,3
   if rspd.eof then
   else
      ip=rspd("b")
   end if
   rspd.close
   set rspd=nothing
   pdconn.close
   set pdconn=nothing
%>
<html>
<head>
<title>�ٸ�IP��ѯ</title>
</head>
<body bgcolor=#800000 background="../bgcheetah.gif" text="#000000">
<BR>
<p align="center"><font size="6" face="����">�ٸ�IP��¼</font></p>
<table border=1 align=center width=400 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" style=font-size:9pt>
<tr>
    <td align=center>�ǡ�¼���ա���</td><td height=30 align=center>�ǡ�¼���ɡ���</td><td align=center>������</td>
</tr>
<%
userip=split(ip,"|")
ipn=ubound(userip)
for i=0 to ipn-1
ipi=split(userip(i),"{]")
ipa=split(ipi(0),".")
ipc=ipi(1)
if ipi(1)="" then ipc="δ֪"
%>
<tr>
    <td align=center><%=ipc%></td><td height=30 align=center><%=ipa(0)&"."&ipa(1)&"."&ipa(2)%>.*</td><td align=center><%=ipe(ipi(0))%></td>
</tr>
<%
next
%>
</table>
<p>��</p>
</body>
</html>
<%
function ipe(ipp)

ip=ipp
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
rs.open "select * FROM z WHERE a<="& num &" and b>="&num,conn
if rs.eof or rs.bof then
	country="����"
	city="δ֪"
else
	country=rs("c")
	city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
ipe="����:"& country &"������:"& city
rs.close
end function


set rs=nothing	
conn.close
set conn=nothing
%>