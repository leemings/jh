<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
%><LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666><table cellpadding=0 cellspacing=0 border=0 width=95% align=center> 
<tr>
<td>
<table cellpadding=3 cellspacing=1 border=0 width=100%>
<tr>
<td align=center colspan="2">
</td>
</tr>
<tr> 
<td width="20%" valign=top> </td>
<td width="80%" valign=top> 
<table width=100% cellspacing=0 cellpadding=0 >
<tr> 
<td height=25 > &nbsp;&nbsp;<b><br>
���߱��ݽ������ݿ�</b>( ����Ȼ��Ҫ��ʱ�����ص����ر��� 
)<br>
</font></td>
</tr>
<%
if IsObjInstalled("Scripting.FileSystemObject") then
dim method
method=request.querystring("method")
if method="" then 
%>
<form method="post" action="BData.asp?method=Backup">
<tr> 
<td height=100 > &nbsp;&nbsp;<br>
���ݺ�Ľ������ݿ����� 
<input type=text size=30 name=DBpath>
&nbsp;&nbsp;
<input type=submit value="ȷ��">
<br>
<br>
<p><font size="2"> &nbsp;&nbsp;��������д���ݺ�Ľ������ݿ�������������׺����·����<br>
&nbsp;&nbsp;������������������������Ľ��������ݿ⣬�Ա�֤�����û����ݰ�ȫ��<br>
&nbsp;&nbsp;ע�⣺���ݺ�����ݿⱣ����DataBackupĿ¼���档</font></p>
<p>&nbsp;</p>
</td>
</tr>
</form>
<%
elseif method="Backup" then 
Dbpath=trim(request.form("Dbpath"))
if chkfname(Dbpath)=1 then ErrAlt("���ݿ����ֲ��ܰ�����׺����\ /\\|*<>?\"": �ȷ���.")
Dbpath="../DataBackup/#" & Dbpath & ".asp"
Dbpath=server.mappath(Dbpath)
aqjh_data="../aqjh_data/aqjh.asp"
db2=server.mappath(aqjh_data)
Set Fso=server.createobject("scripting.filesystemobject")
fso.copyfile db2,dbpath,true
response.write "<tr><td height=25><br><br><li>�ɹ���������<br><br><br></td></tr>"
Set fso=Nothing
end if 
else
%>
<tr> 
<td height=30 > 
<br>
<li>FSOȨ���ѱ����ķ���������Ա�رգ� 
<p>&nbsp;</p>
</td>
</tr>
<%end if%>
</table>
</td>
<%
Function chkfname(fn)
dim chk2
chk2=0
if instr(fn,"/")<>0 then chk2=1
if instr(fn,"\")<>0 then chk2=1
if instr(fn,"|")<>0 then chk2=1
if instr(fn,"*")<>0 then chk2=1
if instr(fn,"""")<>0 then chk2=1
if instr(fn,":")<>0 then chk2=1
if instr(fn,"?")<>0 then chk2=1
if instr(fn,">")<>0 then chk2=1
if instr(fn,"<")<>0 then chk2=1
if instr(fn,".")<>0 then chk2=1
chkfname=chk2
End Function
Sub ErrAlt(says)
%>
<script language="JavaScript">
alert("<%=says%>");
history.go(-1);
</script>
<%
response.end
end sub
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>
</body>