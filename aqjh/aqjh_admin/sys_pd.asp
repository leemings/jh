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
<!--#include file="../paodian.asp"-->
<html>
<head>
<title>ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> ���ݷ����á�</div>
<hr noshade size="1" color=009900>
<b>����ʾ��</b>ע�⣬���в��������Ϊ�գ���������ȫ���ַ����������
<a href="javascript:history.go(0)">ˢ��</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="text" name="aqjh_paofen" value="<%=aqjh_paofen%>">
�ݷ�ʱ�ӵ�����ֵ(�书Ϊaqjh_paofen/2)<br>
<input type="text" name="aqjh_paofencd" value="<%=aqjh_paofencd%>">
�ݷ�ʱ�ӵĴ�㣬��ֵ��Ҫ̫�ߣ����򽭺����ң�<br>
<input type="text" name="aqjh_paofenyin" value="<%=aqjh_paofenyin%>">
�ݷ�ʱ��������aqjh_paofen*aqjh_paofenyin=<%=aqjh_paofen*aqjh_paofenyin%><br>
<input type="text" name="hy1" value="<%=hy1%>">
һ����Ա�ݷֱ���<br>
<input type="text" name="hy2" value="<%=hy2%>">
������Ա�ݷֱ���<br>
<input type="text" name="hy3" value="<%=hy3%>">
������Ա�ݷֱ���<br>
<input type="text" name="hy4" value="<%=hy4%>">
�ļ���Ա�ݷֱ���<br>
<input type="text" name="paodian" value="<%=paodian%>">
�ݵ��Ա�ı���<br><br>
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
writeStr=writeStr & "'*****************�������*********************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_paofen="& request.form("aqjh_paofen")&"      '�ݷ�ʱ�ӵ�����ֵ(�书Ϊaqjh_paofen/2)"& chr(13)& chr(10)
writeStr=writeStr & "aqjh_paofencd="& request.form("aqjh_paofencd")&"      '�ݷ�ʱ�ӵĴ�㣬��ֵ��Ҫ̫�ߣ����򽭺����ң�"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_paofenyin="& request.form("aqjh_paofenyin")&"      '�ݷ�ʱ��������aqjh_paofen*aqjh_paofenyin"& chr(13) & chr(10)
writeStr=writeStr & "hy1="& request.form("hy1")&"      'һ����Ա�ݷֱ���"& chr(13) & chr(10)
writeStr=writeStr & "hy2="& request.form("hy2")&"      '������Ա�ݷֱ���"& chr(13) & chr(10)
writeStr=writeStr & "hy3="& request.form("hy3")&"      '������Ա�ݷֱ���"& chr(13) & chr(10)
writeStr=writeStr & "hy4="& request.form("hy4")&"      '�ļ���Ա�ݷֱ���"& chr(13) & chr(10)
writeStr=writeStr & "paodian="& request.form("paodian")&"      '�ݵ��Ա�ݷֱ���"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../paodian.asp")
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
alert("���ķ�����֧��FSOȨ�ޣ��µ�paodian.asp���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��paodian.asp��ԭ���Ĵ���")
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