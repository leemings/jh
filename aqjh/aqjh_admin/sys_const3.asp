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
<!--#include file="../const3.asp"-->
<html>
<head>
<title>ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> ���������á�</div>
<hr noshade size="1" color=009900>
<b>����ʾ��</b>ע�⣬���в��������Ϊ�գ���������ȫ���ַ����������
<a href="javascript:history.go(0)">ˢ��</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="text" name="aqjh_ifnpc" value="<%=aqjh_ifnpc%>">
NPC����,1Ϊ��,0Ϊ��<br>
<input type="text" name="aqjh_chatkptime" value="<%=aqjh_chatkptime%>">
��������ÿ�ʱ�����ƣ�ÿСʱǰxx���ӣ����ܳ���60<br>
<input type="text" name="aqjh_onlinekill" value="<%=aqjh_onlinekill%>">
��������������ޣ����ڴ���������ֹ����<br>
<input type="text" name="chatshuaxin_time" value="<%=chatshuaxin_time%>">
�����ҷ�ˢ��ʱ��Ϊxx��<br>
<input type="text" name="aqjh_dongtai" value="<%=aqjh_dongtai%>">
�Ƿ���ʾ��̬��棬�������chat/gg.js�ļ����ݣ�0Ϊ����ʾ��1Ϊ��ʾ<br>
<input type="text" name="aqjh_disproxy" value="<%=aqjh_disproxy%>">
�����������1Ϊ���ƣ�0Ϊ������<br>
<input type="text" name="aqjh_yjdh" value="<%=aqjh_yjdh%>">
�Ƿ�����һ�����,1Ϊ����,0Ϊ������<br>
<input type="text" name="aqjh_myie" value="<%=aqjh_myie%>">
�Ƿ�����myie�ര�������,1Ϊ����,0Ϊ������<br>
<input type="text" name="aqjh_sfkf" value="<%=aqjh_sfkf%>">
�Ƿ񿪷����Ŵ������书�ܣ�0Ϊ���ţ�1Ϊ�ر�<br>
<input type="text" name="aqjh_chat1" value="<%=aqjh_chat1%>">
��Թ��������<br>
<input type="text" name="aqjh_chat2" value="<%=aqjh_chat2%>">
��ս��������<br>
<input type="text" name="aqjh_chat3" value="<%=aqjh_chat3%>">
npc�������ƣ�Ĭ��Ϊ�������<br>
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
writeStr=writeStr & "'*****************��������*********************"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_ifnpc="& request.form("aqjh_ifnpc")&"      'NPC����,1Ϊ��,0Ϊ��"& chr(13)& chr(10)
writeStr=writeStr & "aqjh_chatkptime="& request.form("aqjh_chatkptime")&"      '��������ÿ�ʱ�����ƣ�ÿСʱǰxx���ӣ����ܳ���60"&chr(13) & chr(10)
writeStr=writeStr & "aqjh_onlinekill="& request.form("aqjh_onlinekill")&"      '��������������ޣ����ڴ���������ֹ����"&chr(13) & chr(10)
writeStr=writeStr & "chatshuaxin_time="& request.form("chatshuaxin_time")&"      '�����ҷ�ˢ��ʱ��Ϊxx��"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_dongtai="& request.form("aqjh_dongtai")&"      '�Ƿ���ʾ��̬��棬�������chat/gg.js�ļ����ݣ�0Ϊ����ʾ��1Ϊ��ʾ"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_disproxy="& request.form("aqjh_disproxy")&"      '�����������1Ϊ���ƣ�0Ϊ������"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_yjdh="& request.form("aqjh_yjdh")&"      '�Ƿ�����һ�����,1Ϊ����,0Ϊ������"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_myie="& request.form("aqjh_myie")&"      '�Ƿ�����myie��½,1Ϊ����,0Ϊ������"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_sfkf="& request.form("aqjh_sfkf")&"      '�Ƿ񿪷����Ŵ������书�ܣ�0Ϊ���ţ�1Ϊ�ر�"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_chat1="&""""& request.form("aqjh_chat1")&""""&"      '��Թ��������"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_chat2="&""""& request.form("aqjh_chat2")&""""&"      '��ս��������"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_chat3="&""""& request.form("aqjh_chat3")&""""&"      'npc�������ƣ�Ĭ��Ϊ�������"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../const3.asp")
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
alert("���ķ�����֧��FSOȨ�ޣ��µ�const3.asp���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��const3.asp��ԭ���Ĵ���")
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