<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
%>
<!--#include file="../const4.asp"-->
<html>
<head>
<title>ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> �������ҽ�����Ϣ��</div>
<hr noshade size="1" color=009900>
<b>����ʾ��</b>ע�⣬���в��������Ϊ�գ��˴��ı�����һ��Ҫ��ȫ���ַ���
<a href="javascript:history.go(0)">ˢ��</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="text" name="aqjh_userinto1" value="<%=aqjh_userinto1%>" size=80>
//СϺ<br>
<input type="text" name="aqjh_userinto2" value="<%=aqjh_userinto2%>" size=80>
//��Ϻ<br>
<input type="text" name="aqjh_userinto3" value="<%=aqjh_userinto3%>" size=80>
//��Ա<br>
<input type="text" name="aqjh_userinto4" value="<%=aqjh_userinto4%>" size=80>
//���<br>
<input type="text" name="aqjh_userinto5" value="<%=aqjh_userinto5%>" size=80>
//����<br>
<input type="text" name="aqjh_userinto6" value="<%=aqjh_userinto6%>" size=80>
//����<br>
<input type="text" name="aqjh_userinto7" value="<%=aqjh_userinto7%>" size=80>
//��״Ԫ<br>
<input type="text" name="aqjh_userinto8" value="<%=aqjh_userinto8%>" size=80>
//����<br>
<input type="text" name="aqjh_userinto9" value="<%=aqjh_userinto9%>" size=80>
//����<br>
<input type="text" name="aqjh_userinto10" value="<%=aqjh_userinto10%>" size=80>
//����<br>
<input type="text" name="aqjh_userinto11" value="<%=aqjh_userinto11%>" size=80>
//�ٸ�<br>
<input type="text" name="aqjh_userinto12" value="<%=aqjh_userinto12%>" size=80>
//վ��<br>
<input type="text" name="aqjh_userinto13" value="<%=aqjh_userinto13%>" size=80>
//����<br>
<input type="text" name="aqjh_userinto14" value="<%=aqjh_userinto14%>" size=80>
//ة��<br>
<input type="text" name="aqjh_userinto15" value="<%=aqjh_userinto15%>" size=80>
//����<br>
<BR>
<font color=red><b>������һ��Ҫ��ȫ�ǣ�##��ʾ��½����Ϣ[����ID�͵ȼ�]��%%��ʾ��½��������</b></font>
<br><br>
<input type="submit" name="Submit" value="�޸�����">
<input type="hidden" name="set" value="ok">
</p>
</form>
</body>
</html>
<%
response.end
else
call setconst()
end if
sub setconst()
%>
<!--#include file="chkform.asp"-->
<%
writeStr=writeStr & "'*****************�����ҽ�����Ϣ*********************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto1="&""""& request.form("aqjh_userinto1")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto2="&""""& request.form("aqjh_userinto2")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto3="&""""& request.form("aqjh_userinto3")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto4="&""""& request.form("aqjh_userinto4")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto5="&""""& request.form("aqjh_userinto5")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto6="&""""& request.form("aqjh_userinto6")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto7="&""""& request.form("aqjh_userinto7")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto8="&""""& request.form("aqjh_userinto8")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto9="&""""& request.form("aqjh_userinto9")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto10="&""""& request.form("aqjh_userinto10")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto11="&""""& request.form("aqjh_userinto11")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto12="&""""& request.form("aqjh_userinto12")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto13="&""""& request.form("aqjh_userinto13")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto14="&""""& request.form("aqjh_userinto14")&""""& chr(13) & chr(10)
writeStr=writeStr & "aqjh_userinto15="&""""& request.form("aqjh_userinto15")&""""& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write "�ύ�ɹ�"
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const4.asp")
Set fs = CreateObject("scripting.filesystemobject")
If Not Fs.FILEEXISTS(toppath) Then 
Set Ts = fs.createtextfile(toppath, True)
Ts.close
end if
Set Ts= Fs.OpenTextFile(toppath,2)
Ts.writeline (writeStr)
Ts.Close
Set Ts=Nothing
Set Fs=Nothing
%>
<script>
alert("���ķ�����֧��FSOȨ�ޣ��µ�const4.asp���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��const4.asp��ԭ���Ĵ���")
</script>
<%
end if
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
Function rrstr(str)
str="|" & str & "|"
str=replace(str,"||","|")
rrstr=str
End Function 
%>