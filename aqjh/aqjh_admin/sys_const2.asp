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
<!--#include file="../const2.asp"-->
<html>
<head>
<title>ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> �������������á�</div>
<hr noshade size="1" color=009900>
<b>����ʾ��</b>ע�⣬���в��������Ϊ�գ���������ȫ���ַ����������
<a href="javascript:history.go(0)">ˢ��</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<p><font color=red>*************ÿ���½�������***********</font><br>
<input type="text" name="jl_dd" value="<%=jl_dd%>">
ÿ������һ����н���(��ǰ��Ϊÿ��28��)<br>
<input type="text" name="jl_allvalue" value="<%=jl_allvalue%>">
��������Ҫ�ﵽ�Ļ���<br>
<input type="text" name="jl_yjf_top" value="<%=jl_yjf_top%>">
�»��ֽ���ǰ������<br>
<input type="text" name="jl_yjf_jk" value="<%=jl_yjf_jk%>">
�»��ֽ�����ϵ��<br>
<input type="text" name="jl_lr_jk" value="<%=jl_lr_jk%>">
���˽�����ϵ��<br>
<input type="text" name="jl_lr_yl" value="<%=jl_lr_yl%>">
���˽�������ϵ��<br>
<p><font color=red>************��ֹע����¼������***********</font><br>
<textarea name="aqjh_disloginname" cols="80" rows="10"><%=aqjh_disloginname%></textarea><br>
ע�⣺����Լ���վ�������ȥ�ˣ�����վ���ͽ���������һ��Ҫ���ո�ʽ������������벻Ҫ�޸�!<br>
<p><font color=red>************�����ҽ�ֹ���໰***********</font><br>
<textarea name="aqjh_badword" cols="80" rows="10"><%=aqjh_badword%></textarea><br>
ע�⣺һ��Ҫ���ո�ʽ������������벻Ҫ�޸�!<br>
<p><font color=red>************������λ��chat/picwords�ļ�����***********</font><br>
<textarea name="aqjh_Zshow" cols="80" rows="10" readonly><%=aqjh_Zshow%></textarea><br>
ע�⣺һ��Ҫ���ո�ʽ������������벻Ҫ�޸�!<br>
<p> 
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
writeStr=writeStr & "'****************ÿ���½�������***************"& chr(13) & chr(10)
writeStr=writeStr & "jl_dd="&request.form("jl_dd")&"      'ÿ������һ����н���(��ǰ��Ϊÿ��28��)"&chr(13)&chr(10)
writeStr=writeStr & "jl_allvalue="&request.form("jl_allvalue")&"      '������Ҫ�ﵽһ���Ļ���"&chr(13)&chr(10)
writeStr=writeStr & "jl_yjf_top="&request.form("jl_yjf_top")&"      '�»��ֽ���ǰ������"&chr(13)&chr(10)
writeStr=writeStr & "jl_yjf_jk="&request.form("jl_yjf_jk")&"      '�»��ֽ�����ϵ��"&chr(13)&chr(10)
writeStr=writeStr & "jl_lr_jk="&request.form("jl_lr_jk")&"      '���˽�����ϵ��"&chr(13)&chr(10)
writeStr=writeStr & "jl_lr_yl="&request.form("jl_lr_yl")&"      '���˽�������ϵ��"&chr(13)&chr(10)
writeStr=writeStr & "'****************��ֹע����¼������***************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_disloginname="&""""& request.form("aqjh_disloginname")&""""&"      '��ֹ��¼������"& chr(13) & chr(10)
writeStr=writeStr & "'****************�����ҽ�ֹ���໰***************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_badword="&""""& request.form("aqjh_badword")&""""&"      '�����ҽ�ֹ���໰"& chr(13) & chr(10)
writeStr=writeStr & "'****************������λ��chat/picwords�ļ�����***************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_Zshow="&""""& request.form("aqjh_Zshow")&""""&"      '����"& chr(13) & chr(10)
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const2.asp")
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
alert("���ķ�����֧��FSOȨ�ޣ��µ�const2���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��const2��ԭ���Ĵ���")
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