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
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table cellpadding=0 cellspacing=0 border=0 width=95% align=center> 
<tr>
<td>
<table cellpadding=3 cellspacing=1 border=0 width=100%>
<tr>
<td align=center colspan="2">
</td>
</tr>
<tr>
<td width="20%" valign=top>
</td>
<td width="80%" valign=top>
<table width=100% cellspacing=0 cellpadding=0 > 
<tr>
<td height=25 > &nbsp;&nbsp;<b><br>
���߻ָ��������ݿ�--��ԭ���� <br>
</b></td>
</tr> 
<%
if IsObjInstalled("Scripting.FileSystemObject") then
Set Fso=server.createobject("scripting.filesystemobject")
dim method
method=request.querystring("method")
if method="" then 
FoldName=server.mappath("../DataBackup") 
Set f1 = Fso.GetFolder(FoldName)
Set pFiles = f1.Files
%>
<form method="post" action="RData.asp?method=Backup">
<tr>
<td height=100 > &nbsp;&nbsp; 
<p><font size="2"><br>
��ѡ���Ѿ����ݵ����ݿ�</font>�� 
<select name="DBpath">
<%for each file in f1.Files
tempwrite= tempwrite & "<option value=" & file.name &">" & file.name &"</option>"
next 
response.write tempwrite
set f1=nothing
set fso=nothing
%>
</select>
&nbsp; 
<input type=submit value="ȷ��">
<br>
<br>------------------------------------------------<br>
<br>
&nbsp;&nbsp;<font size="2">�������������������ԭ���Ѿ����ݵĽ��������ݿ⣡<br>
-----------------------------------------<br>���ݵĽ������ݿ�λ�ڽ���Ŀ¼DataBackup�ļ���<br></font>
</p>
</tr> 
</form>
<%
Set f1 = Nothing
Set pFiles = Nothing
elseif method="Backup" then 
Dbpath=trim(request.form("Dbpath"))
Dbpath="../DataBackup/" & Dbpath 
Dbpath=server.mappath(Dbpath) 
db2=server.mappath("../aqjh_data/aqjh_home.asp")
if fso.fileexists(dbpath) then 
fso.copyfile dbpath,db2,true
response.write "<tr><td height=25><br><br><li>�ɹ��ָ����ݣ�</td></tr>"
else
response.write "<tr><td height=25><br><br><li>����Ŀ¼�²������ı����ļ��������Ƿ�������ȷ�����ݿ�����</td></tr>" 
end if 
response.write "<tr><td height=25><br><li>��ɣ�</td></tr>"
end if 
Set Fso=Nothing
else
%> 
<tr>
<td height=30 >
<li>�����ܲ�����,������δ����FSOȨ��!
</td>
</tr>
<%end if%>
</table> 
</td>
<%
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