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
<!--#include file="../const5.asp"-->
<html>
<head>
<title>ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> ����̳����������</div>
<hr noshade size="1" color=009900>
<b>����ʾ��</b>ע�⣬���в��������Ϊ�գ���������ȫ���ַ����������
<a href="javascript:history.go(0)">ˢ��</a> 
<hr noshade size="1" color=009900>
***��̳��ز���[1000����=1����=2Ԫ��]��˲�Ҫ����̫��***<br>
���½����ͳͷ��ľ�Ϊ�ݶ�����
<form name="form1" method="post" action="">
<font color=red>************ ���� ***********</font>
<BR>
<input type="text" name="bbs_add_dj" value="<%=bbs_add_dj%>">
�������������Ҫ�Ľ����ȼ�<br>
<input type="text" name="de_grade" value="<%=de_grade%>">
��������͹���ȼ�<br>
<font color=red>************ ���� ***********</font>
<BR>
<input type="text" name="bbs_add1" value="<%=bbs_add1%>">
����+��<br>
<input type="text" name="bbs_add2" value="<%=bbs_add2%>">
�ظ�+��<br>
<input type="text" name="bbs_add3" value="<%=bbs_add3%>">
��ǰ����+��<br>
<input type="text" name="bbs_add4" value="<%=bbs_add4%>">
�̶ܹ�+��<br>
<input type="text" name="bbs_add5" value="<%=bbs_add5%>">
�̶�+��<br>
<input type="text" name="bbs_add6" value="<%=bbs_add6%>">
������+��<br>
<font color=red>************ �۳� ***********</font>
<BR>
<input type="text" name="bbs_del1" value="<%=bbs_del1%>">
ɾ������-��<br>
<input type="text" name="bbs_del2" value="<%=bbs_del2%>">
ɾ���ظ�-��<br>
<input type="text" name="bbs_del3" value="<%=bbs_del3%>">
�ر�����-��<br>
<input type="text" name="bbs_del4" value="<%=bbs_del4%>">
�ƶ�����-��<br>
<font color=red>************ ϵͳ ***********</font>
<BR>
<input type="text" name="clubname" value="<%=clubname%>">
��̳����<br>
<input type="text" name="allclass" value="<%=allclass%>">
��������ҳ�Զ�չ��������̳0=�� 1=��<br>
<input type="text" name="badwords" value="<%=badwords%>">
���������֣��������á�|���ָ���<br>
<input type="text" name="kickwords" value="<%=kickwords%>">
�����û���ǰ,��,�м��á�|���ָ���<br>
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
writeStr=writeStr & "'***��̳��ز���[1000����=1����=2Ԫ��]��˲�Ҫ����̫��***"& chr(13) & chr(10)
writeStr=writeStr & "'���½����ͳͷ��ľ�Ϊ�ݶ�����"& chr(13) & chr(10)
writeStr=writeStr & "'-------------����-----------------"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add_dj="& request.form("bbs_add_dj")&"      '�������������Ҫ�Ľ����ȼ�"& chr(13)& chr(10)
writeStr=writeStr & "de_grade="& request.form("de_grade")&"      '��������͹���ȼ�"&chr(13) & chr(10)
writeStr=writeStr & "'-------------����-----------------"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add1="& request.form("bbs_add1")&"      '����+��"&chr(13) & chr(10)
writeStr=writeStr & "bbs_add2="& request.form("bbs_add2")&"      '�ظ�+��"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add3="& request.form("bbs_add3")&"      '��ǰ����+��"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add4="& request.form("bbs_add4")&"      '�̶ܹ�+��"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add5="& request.form("bbs_add5")&"      '�̶�+��"& chr(13) & chr(10)
writeStr=writeStr & "bbs_add6="& request.form("bbs_add6")&"      '������+��"& chr(13) & chr(10)
writeStr=writeStr & "'-------------�۳�-----------------"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del1="& request.form("bbs_del1")&"      'ɾ������-��"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del2="&""""& request.form("bbs_del2")&""""&"      'ɾ���ظ�-��"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del3="&""""& request.form("bbs_del3")&""""&"      '�ر�����-��"& chr(13) & chr(10)
writeStr=writeStr & "bbs_del4="&""""& request.form("bbs_del4")&""""&"      '�ƶ�����-��"& chr(13) & chr(10)
writeStr=writeStr & "'-------------ϵͳ-----------------"& chr(13) & chr(10)
writeStr=writeStr & "clubname="&""""& request.form("clubname")&""""&"      '��̳����"& chr(13) & chr(10)
writeStr=writeStr & "cluburl="&""""& request.form("cluburl")&""""&"      '��̳url��ַ"& chr(13) & chr(10)
writeStr=writeStr & "badwords="&""""& request.form("badwords")&""""&"      '���������֣��������á�|���ָ���"& chr(13) & chr(10)
writeStr=writeStr & "allclass="& request.form("allclass")&"      '��������ҳ�Զ�չ��������̳0=�� 1=��"& chr(13) & chr(10)
writeStr=writeStr & "kickwords="&""""& request.form("kickwords")&""""&"      '�����û���ǰ,��,�м��á�|���ָ���"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const5.asp")
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
alert("���ķ�����֧��FSOȨ�ޣ��µ�const5.asp���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��const5.asp��ԭ���Ĵ���")
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