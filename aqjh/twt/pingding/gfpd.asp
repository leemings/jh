<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>�ٸ�����</title>
</head>
<body bgcolor=#800000 background="../bgcheetah.gif" text="#000000">
<BR>
<p align="center"><font size="6" face="����">����Ա����</font></p>
<table border=1 align=center width=700 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" style=font-size:9pt>
<tr bgcolor="#3399FF">
    <td width=80 height=40 align=center><font color=white>�ա���</td>
    <td width=40 align=center><font color=white>����ȼ�</td>
    <td width=40 align=center><font color=white>ս���ȼ�</td>
    <td width=80 align=center><font color=white>����½����</td>
    <td width=80 align=center><font color=white>�վ�����</td>
    <td width=80 align=center><font color=white>��ʧ����</td>
    <td width=40 align=center><font color=white>IP�������</font></td>
    <td width=260 align=center><font color=white>�����ο�</td>
</tr>
<%
date2=split(date(),"-")
rs.open "SELECT * FROM �û� where ����='�ٸ�' and ����<>'��ļ���' order by grade DESC",conn
do while not rs.eof and not rs.bof
ipcname=rs("����")
xsts=int(DateDiff("d",rs("lasttime"),now()))
rjzx=int(rs("mvalue")/date2(2)/10)

if rjzx<30 or xsts>=5 then
 pdck="<font color=red><b>�ù���Ա����ʧְ�����鳷ְ</b></font>"
elseif rjzx>=30 and rjzx<60 then
 pdck="<b>�ù���Ա���ֺܲ�뾡�����</b>"
elseif rjzx>=60 and rjzx<180 and xsts<5 then
 pdck="<b>�ù���Ա����һ�㣬�����Ŭ��</b>"
elseif rjzx>=180 and xsts<3 then
 pdck="<font color=green><b>�ù���Ա���渺�𣬽��齱��</b></font>"
elseif ipc(ipcname)>=5 then
 pdck="<font color=red><b>�ù���Աip�䶯Ƶ������ȥ��̳˵��ԭ��</b></font>"
else
 pdck="<b>�ù���Ա����һ�㣬�����Ŭ��</b>"
end if

%>
<tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';">
    <td align=center width=80><%=rs("����")%></td>
    <td align=center width=40><%=rs("grade")%></td>
    <td align=center width=40><%=rs("�ȼ�")%></td>
    <td align=center width=80><%=rs("lasttime")%></td>
    <td align=center width=80><%=rjzx%> ����</td>
    <td align=center width=80><%=xsts%></td>
    <td align=center width=40><a href=gfip.asp?name=<%=rs("����")%>><%=ipc(ipcname)%></a></td>
    <td align=center width=260><%=pdck%></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing


function ipc(name)


   pddb="pd.mdb"
   Set pdconn = Server.CreateObject("ADODB.Connection")
   Set rspd=Server.CreateObject("ADODB.RecordSet")
   pdconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(pddb)
   pdconn.open pdconnstr
   sqlpd="select * from p where a='"&name&"'"
   rspd.open sqlpd,pdconn,1,3
   if not rspd.eof then
	ipcs=split(rspd("b"),"|")
	ipc=ubound(ipcs)-1
   else
	ipc="NO"
   end if
   rspd.close
   set rspd=nothing
   pdconn.close
   set pdconn=nothing


end function
%>
</table>
</body>
</html>
