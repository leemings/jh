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
<!--#include file="../const1.asp"-->
<html>
<head>
<title>ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> ������״̬���á�</div>
<hr noshade size="1" color=009900>
<b>����ʾ��</b>ע�⣬���в��������Ϊ�գ���������ȫ���ַ����������
<a href="javascript:history.go(0)">ˢ��</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="text" name="reg_ml" value="<%=reg_ml%>">
����<br>
<input type="text" name="reg_dd" value="<%=reg_dd%>">
����<br>
<input type="text" name="reg_wg" value="<%=reg_wg%>">
�书<br>
<input type="text" name="reg_nl" value="<%=reg_nl%>">
����<br>
<input type="text" name="reg_tl" value="<%=reg_tl%>">
����<br>
<input type="text" name="reg_gj" value="<%=reg_gj%>">
����<br>
<input type="text" name="reg_fy" value="<%=reg_fy%>">
����<br>
<input type="text" name="reg_fl" value="<%=reg_fl%>">
����<br>
<input type="text" name="reg_zz" value="<%=reg_zz%>">
֪��<br>
<input type="text" name="reg_qg" value="<%=reg_qg%>">
�Ṧ<br>
<input type="text" name="reg_yl" value="<%=reg_yl%>">
����<br>
<input type="text" name="reg_ck" value="<%=reg_ck%>">
���<br>
<input type="text" name="reg_jb" value="<%=reg_jb%>">
���<br>
<input type="text" name="reg_dj" value="<%=reg_dj%>">
ע��ʱ�ĵȼ�<font color=red>[����ʹ��Ĭ��ֵ15������Ϊϵͳ���õ�18������Ϊ����]</font><br>
<input type="text" name="reg_hyjk" value="<%=reg_hyjk%>">
��Ա��<br>
<input type="text" name="reg_hydj" value="<%=reg_hydj%>">
��Ա�ȼ�[0��4����0Ϊ�ǵȼ���Ա]<br>
<input type="text" name="reg_hydate" value="<%=reg_hydate%>">
��Ա����[30Ϊһ���µ�ʱ��]<br>
<input type="text" name="reg_hycard" value="<%=reg_hycard%>">
��Ƭ[��ȷ��ʽ(��ɫ����)��<font color=red><u>��Ƭ��1|����;</u>��Ƭ��2|����;</font>]<br><br>
<input type="text" name="reg_yp" value="<%=reg_yp%>">
ҩƷ[��ȷ��ʽ(��ɫ����)��<font color=red><u>ҩƷ��1|����;</u>ҩƷ��2|����;</font>]<br><br>
<input type="text" name="reg_xh" value="<%=reg_xh%>">
�ʻ�[��ȷ��ʽ(��ɫ����)��<font color=red><u>�ʻ���1|����;</u>�ʻ���2|����;</font>]<br><br>
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
writeStr=writeStr & "'*****************��������*********************"& chr(13) & chr(10)
writeStr=writeStr & "reg_ml="& request.form("reg_ml")&"      '����"& chr(13)& chr(10)
writeStr=writeStr & "reg_dd="& request.form("reg_dd")&"      '����"&chr(13) & chr(10)
writeStr=writeStr & "reg_wg="& request.form("reg_wg")&"      '�书"&chr(13) & chr(10)
writeStr=writeStr & "reg_nl="& request.form("reg_nl")&"      '����"& chr(13) & chr(10)
writeStr=writeStr & "reg_tl="& request.form("reg_tl")&"      '����"& chr(13) & chr(10)
writeStr=writeStr & "reg_gj="& request.form("reg_gj")&"      '����"& chr(13) & chr(10)
writeStr=writeStr & "reg_fy="& request.form("reg_fy")&"      '����"& chr(13) & chr(10)
writeStr=writeStr & "reg_fl="& request.form("reg_fl")&"      '����"& chr(13) & chr(10)
writeStr=writeStr & "reg_zz="& request.form("reg_zz")&"      '֪��"& chr(13) & chr(10)
writeStr=writeStr & "reg_qg="& request.form("reg_qg")&"      '�Ṧ"& chr(13) & chr(10)
writeStr=writeStr & "reg_yl="& request.form("reg_yl")&"      '����"& chr(13) & chr(10)
writeStr=writeStr & "reg_ck="& request.form("reg_ck")&"      '���"& chr(13) & chr(10)
writeStr=writeStr & "reg_jb="& request.form("reg_jb")&"      '���"& chr(13) & chr(10)
writeStr=writeStr & "reg_dj="& request.form("reg_dj")&"      'ע��ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "reg_hyjk="& request.form("reg_hyjk")&"      '��Ա��"& chr(13) & chr(10)
writeStr=writeStr & "reg_hydj="& request.form("reg_hydj")&"      '��Ա�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "reg_hydate="& request.form("reg_hydate")&"      '��Ա����"& chr(13) & chr(10)
writeStr=writeStr & "reg_hycard="&""""& request.form("reg_hycard")&""""&"      '��Ƭ"& chr(13) & chr(10)
writeStr=writeStr & "reg_yp="&""""& request.form("reg_yp")&""""&"      'ҩƷ"& chr(13) & chr(10)
writeStr=writeStr & "reg_xh="&""""& request.form("reg_xh")&""""&"      '�ʻ�"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const1.asp")
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
alert("���ķ�����֧��FSOȨ�ޣ��µ�const1.asp���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��const1.asp��ԭ���Ĵ���")
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