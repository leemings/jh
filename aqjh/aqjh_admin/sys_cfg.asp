<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
%>
<!--#include file="../config.asp"-->
<html>
<head>
<title>ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%if not request.form("set")="ok" then%>
<div align="center"> ��ϵͳ������</div>
<hr noshade size="1" color=009900>
<b>����ʾ��</b>ע�⣬���в��������Ϊ�գ���������ȫ���ַ����������
<a href="javascript:history.go(0)">ˢ��</a> 
<hr noshade size="1" color=009900>
<form name="form1" method="post" action="">
<input type="hidden" name="aqjh_chatbgcolor" value="<%=Session("afa_chatbgcolor")%>">
<input type="hidden" name="aqjh_chatimage" value="<%=Application("aqjh_chatimage")%>">
<input type="hidden" name="aqjh_chatcolor" value="<%=Application("aqjh_chatcolor")%>">
<input type="hidden" name="aqjh_iplocktime" value="<%=Application("aqjh_iplocktime")%>">
<p><font color=red>************ ������������***********</font><br>
<input type="text" name="aqjh_maxtimeout" value="<%=Application("aqjh_maxtimeout")%>">
�ݵ㳬ʱʱ�䣬��Ҫ����̫�����������۷�����[xx����]<br>
<input type="text" name="aqjh_maxnpc" value="<%=Application("aqjh_maxnpc")%>">
npc����������<br>
<input type="text" name="aqjh_npcwp" value="<%=application("aqjh_npcwp")%>">
npc���ǱȽ�ֵ[���鲻Ҫ�Ķ�]<br>
<input type="text" name="aqjh_npcff" value="<%=application("aqjh_npcff")%>">
npc������Ҫ�ĵȼ�[����,������Ĭ��10������]<br>
<input type="text" name="aqjh_baowuyin" value="<%=Application("aqjh_baowuyin")%>">
ÿһ�������������õ�������<br>
<input type="text" name="aqjh_baowuxl" value="<%=Application("aqjh_baowuxl")%>">
ÿ�������ļ�������100����������4�Σ������õ������޾���4�����ֵ����ս���ȼ�ȡ���ģ�<br>
<input type="text" name="aqjh_killman" value="<%=Application("aqjh_killman")%>">
����һ���ɱ�������ᱦһ����һ��ɱ��<br>
<input type="text" name="aqjh_wgsx" value="<%=aqjh_wgsx%>">
ÿ��һ�����书��ֵ,��ֵ��aqjh_sx=1ʱ��������<br>
<input type="text" name="aqjh_nlsx" value="<%=aqjh_nlsx%>">
ÿ��һ����������ֵ,��ֵ��aqjh_sx=1ʱ��������<br>
<input type="text" name="aqjh_tlsx" value="<%=aqjh_tlsx%>">
ÿ��һ����������ֵ,��ֵ��aqjh_sx=1ʱ��������<br>
<input type="text" name="aqjh_gjsx" value="<%=aqjh_gjsx%>">
ÿ��һ���Ĺ�����ֵ,��ֵ��aqjh_sx=1ʱ��������<br>
<input type="text" name="aqjh_fysx" value="<%=aqjh_fysx%>">
ÿ��һ���ķ�����ֵ,��ֵ��aqjh_sx=1ʱ��������<br>
<input type="hidden" name="aqjh_chat_maxpeople" value="<%=Application("aqjh_chat_maxpeople")%>" readonly>
<p><font color=red>************ ������������***********</font><br>
<input type="text" name="aqjh_closedoor" value="<%=Application("aqjh_closedoor")%>">
�����ҹ���,0Ϊ����,1Ϊ����<br>
<input type="text" name="aqjh_baowu" value="<%=Application("aqjh_baowu")%>">
�Ƿ��б��ﵱΪ0ʱû�б����������ǻ�Ӱ���������ٶȵģ����������ùرգ�<br>
<input type="text" name="aqjh_sx" value="<%=aqjh_sx%>">
�Ƿ�������0û�����ޣ�1������<br>
<input type="text" name="aqjh_automanname" value="<%=Application("aqjh_automanname")%>">
������������<br>
<input type="text" name="aqjh_userout" value="<%=Application("aqjh_userout")%>">
�˳������ҵ���Ϣ(ע�������"һ��ҪΪȫ���ַ�)<br>
<input type="text" name="aqjh_dieip" value="<%=aqjh_dieip%>">
ip���÷���ʹ��ǰ����ʽ��д�������;�Ÿ���,�����뱣���зֺ�<br>
<input type="text" name="aqjh_baowuname" value="<%=Application("aqjh_baowuname")%>">
������������˵��<br>
<input type="text" name="aqjh_baowusm" value="<%=Application("aqjh_baowusm")%>">
�õ���������Ϣ <br>
<hr noshade size="1" color=009900>
<p><font color=red>************ ����ϵͳ����***********</font><br>
<input type="hidden" name="aqjh_chatroomname" value="<%=Application("aqjh_chatroomname")%>" readonly>
<input type="text" name="aqjh_homepage" value="<%=Application("aqjh_homepage")%>">
��������<br>
<input type="hidden" name="aqjh_sn" value="<%=Application("aqjh_sn")%>" readonly>
<input type="text" name="aqjh_qq" value="<%=Application("aqjh_qq")%>">
վ��qq<br>
<input type="text" name="aqjh_email" value="<%=Application("aqjh_email")%>">
վ��email<br>
<input type="hidden" name="aqjh_user" value="<%=Application("aqjh_user")%>" readonly>
<input type="text" name="aqjh_admin" value="<%=Application("aqjh_admin")%>">
��������10����Ա:���֮���԰�Ƕ���,�ָ��磺����,����<br>
<input type="text" name="hidden_admin" value="<%=Application("hidden_admin")%>">
����Ĺ���Ա�����磺����,����<br>
<input type="text" name="aqjh_slbox" value="<%=Application("aqjh_slbox")%>">
���Բ鿴˽�ĵ����磺����,����<br>
<input type="text" name="aqjh_admin_send" value="<%=Application("aqjh_admin_send")%>">
����ү�������|��|����������ҲҪ����|<br>
<input type="text" name="aqjh_guibin" value="<%=Application("aqjh_guibin")%>">
�������(����������κν�����Թ)�������|��|����������ҲҪ����|<br>
<input type="text" name="aqjh_adminuser" value="<%=Application("aqjh_adminuser")%>">
��̨�����˺�<br>
<input type="password" name="aqjh_adminkey" value="<%=Application("aqjh_adminkey")%>">
��̨��������</p>
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
'*************************************************************
'*����Ϊ����ϵͳ�����밴�Լ�����Ҫ���ã�����ǰ��ñ���һ��ԭ�ļ���*
'*************************************************************
writeStr=writeStr & "Application("&""""&"aqjh_chatbgcolor"&""""&")=" &"""" & request.form("aqjh_chatbgcolor")& """"&"      '�����ұ�����ɫ����"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chatimage"&""""&")=" &"""" & request.form("aqjh_chatimage")& """"&"      '�������ⲿ�߿�ͼƬ"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chatcolor"&""""&")=" &"""" & request.form("aqjh_chatcolor")& """"&"      '����������������ɫ"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_iplocktime"&""""&")="& request.form("aqjh_iplocktime")&"      '��IPʱ��"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_maxtimeout"&""""&")="& request.form("aqjh_maxtimeout")&"      '�ݵ㳬ʱʱ�䣬��Ҫ����̫��"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_maxnpc"&""""&")="& request.form("aqjh_maxnpc")&"      'npc����������"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_npcwp"&""""&")="& request.form("aqjh_npcwp")&"      'npc���ǱȽ�ֵ"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_npcff"&""""&")="& request.form("aqjh_npcff")&"      'npc������Ҫ�ĵȼ���������"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowuyin"&""""&")="& request.form("aqjh_baowuyin")&"      'ÿһ���������õ�������"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowuxl"&""""&")="& request.form("aqjh_baowuxl")&"      'ÿ�������ļ�������100����������4�Σ������õ������޾���4�����ֵ����ս���ȼ�ȡ����"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_killman"&""""&")="& request.form("aqjh_killman")&"      '����һ���ɱ�������ᱦһ����һ��ɱ��"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_wgsx="&request.form("aqjh_wgsx")&"      'ÿ��һ�����书��ֵ,��ֵ��aqjh_sx=1ʱ��������"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_nlsx="&request.form("aqjh_nlsx")&"      'ÿ��һ����������ֵ,��ֵ��aqjh_sx=1ʱ��������"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_tlsx="&request.form("aqjh_tlsx")&"      'ÿ��һ����������ֵ,��ֵ��aqjh_sx=1ʱ��������"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_gjsx="&request.form("aqjh_gjsx")&"      'ÿ��һ���Ĺ�����ֵ,��ֵ��aqjh_sx=1ʱ��������"&chr(13)&chr(10)
writeStr=writeStr & "aqjh_fysx="&request.form("aqjh_fysx")&"      'ÿ��һ���ķ�����ֵ,��ֵ��aqjh_sx=1ʱ��������"&chr(13)&chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chat_maxpeople"&""""&")="& request.form("aqjh_chat_maxpeople")&"      '���������������"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_closedoor"&""""&")="& request.form("aqjh_closedoor")&"      '�����ҹ��ţ�1Ϊ����"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowu"&""""&")="& request.form("aqjh_baowu")&"      '�Ƿ��б��ﵱΪ0ʱû�б����������ǻ�Ӱ���������ٶȵģ����������ùرգ�"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_sx="&request.form("aqjh_sx")&"      '�Ƿ�������0û�����ޣ�1������"&chr(13)&chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_automanname"&""""&")=" &"""" & request.form("aqjh_automanname")& """"&"      '������������"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_userout"&""""&")=" &"""" & request.form("aqjh_userout")& """"&"      '�˳�ʱ��������ʾ����Ϣ"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_dieip=" &"""" & request.form("aqjh_dieip")& """"&"      'ip���÷���ʹ��ǰ����ʽ�飬����÷ֺŸ����������뱣���зֺ�"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowuname"&""""&")=" &"""" & request.form("aqjh_baowuname")& """"&"      '������������˵��"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_baowusm"&""""&")=" &"""" & request.form("aqjh_baowusm")& """"&"      '����������������ʾ��Ϣ"& chr(13) & chr(10)
writeStr=writeStr & "'***************************************************"& chr(13) & chr(10)
writeStr=writeStr & "'*���½����������ã��мǲ�Ҫ�Ĵ������޷��������� *"& chr(13) & chr(10)
writeStr=writeStr & "'***************************************************"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_chatroomname"&""""&")=" &"""" & request.form("aqjh_chatroomname")& """"&"      '��������"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_homepage"&""""&")=" &"""" & request.form("aqjh_homepage")& """"&"      '������ҳ"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_sn"&""""&")=" &"""" & request.form("aqjh_sn")& """"&"      '�������кţ������־���޸ĺ����Ը�"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_user"&""""&")=" &"""" & request.form("aqjh_user")& """"&"      '����վ��"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_qq"&""""&")=" &"""" & request.form("aqjh_qq")& """"&"      'վ��qq"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_email"&""""&")=" &"""" & request.form("aqjh_email")& """"&"      'վ��email"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_admin"&""""&")=" &"""" & request.form("aqjh_admin")& """"&"      '����10���Ĺܣ�����ö��Ÿ���"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"hidden_admin"&""""&")=" &"""" & request.form("hidden_admin")& """"&"      '������Ա������ö��Ÿ���"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_slbox"&""""&")=" &"""" & request.form("aqjh_slbox")& """"&"      '���Բ鿴˽�ĵ��ˣ�����ö��Ÿ���"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_admin_send"&""""&")=" &"""" & request.form("aqjh_admin_send")& """"&"      '����ү�������|��|����"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_guibin"&""""&")=" &"""" & request.form("aqjh_guibin")& """"&"      '�������(����������κν�����Թ)�������|��|����"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_adminuser"&""""&")=" &"""" & request.form("aqjh_adminuser")& """"&"      'վ�������û���"& chr(13) & chr(10)
writeStr=writeStr & "Application("&""""&"aqjh_adminkey"&""""&")=" &"""" & request.form("aqjh_adminkey")& """"&"      'վ����������"& chr(13) & chr(10)
writeStr=writeStr & "'*********************************************************"& chr(13) & chr(10)
writeStr=writeStr & "'*���Ҫ���ù�����������chat/jhchat.asp�ļ���ʼ��"& chr(13) & chr(10)
writeStr=writeStr & "'*�ݷֵĵ����������Լ����ã������paodian.asp�ļ���ʼ��"& chr(13) & chr(10)
writeStr=writeStr & "'�޸���ʽ�ȼ��޸�chat/sjfunc/czdj.asp"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../config.asp")
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
alert("���ķ�����֧��FSOȨ�ޣ��µ�config.asp���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��config.asp��ԭ���Ĵ���")
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