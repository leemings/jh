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
<!--#include file="../chat/sjfunc/czdj.asp"-->
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
*****************������Ҫ����ս���ȼ�*****************<Br>
<input type="text" name="jhdj_qlcy" value="<%=jhdj_qlcy%>">
ǧ�ﴫ���ȼ�<br>
<input type="text" name="jhdj_duyao" value="<%=jhdj_duyao%>">
�¶��ȼ�<br>
<input type="text" name="jhdj_tq" value="<%=jhdj_tq%>">
͵Ǯ�ȼ�<br>
<input type="text" name="jhdj_xx" value="<%=jhdj_xx%>">
���Ǵ�<br>
<input type="text" name="jhdj_aq" value="<%=jhdj_aq%>">
�����ȼ�<br>
<input type="text" name="jhdj_ai" value="<%=jhdj_ai%>">
ħ���ȼ�<br>
<input type="text" name="jhdj_fz" value="<%=jhdj_fz%>">
���й���<br>
<input type="text" name="jhdj_cs" value="<%=jhdj_cs%>">
��������<br>
<input type="text" name="jhdj_sq" value="<%=jhdj_sq%>">
��Ǯ�ȼ�<br>
<input type="text" name="jhdj_so" value="<%=jhdj_so%>">
�ͷ����ȼ�<br>
<input type="text" name="jhdj_zs" value="<%=jhdj_zs%>">
���ͺ��Ķ��ȼ�<br>
<input type="text" name="jhdj_pm" value="<%=jhdj_pm%>">
�����ȼ�<br>
<input type="text" name="jhdj_bs" value="<%=jhdj_bs%>">
��ʦ�ȼ�<br>
<input type="text" name="jhdj_jj" value="<%=jhdj_jj%>">
������ĵȼ�<br>
<input type="text" name="jhdj_zz" value="<%=jhdj_zz%>">
ת����Ҫ�ĵȼ�<br>
<input type="text" name="jhdj_nh" value="<%=jhdj_nh%>">
ŭ����Ҫ�ȼ�<br>
<input type="text" name="jhdj_xt" value="<%=jhdj_xt%>">
������Ҫ�ȼ�<br>
<input type="text" name="jhdj_hy" value="<%=jhdj_hy%>">
�������<br>
<input type="text" name="jhdj_db" value="<%=jhdj_db%>">
�Ĳ���ׯ<br>
<input type="text" name="jhdj_qr" value="<%=jhdj_qr%>">
���ˡ�������Ҫ�ȼ�<br>
<input type="text" name="jhdj_dt" value="<%=jhdj_dt%>">
������Ҫ�ȼ�<br>
<input type="text" name="jhdj_chi" value="<%=jhdj_chi%>">
��ǩ�ȼ�<br>
<input type="text" name="jhdj_hua" value="<%=jhdj_hua%>">
�ͻ��ȼ�<br>
<input type="text" name="jhdj_zl" value="<%=jhdj_zl%>">
���뿪�ȼ�<br>
<input type="text" name="jhdj_xq" value="<%=jhdj_xq%>">
���鹦�ܵȼ�<br>
<input type="text" name="jhdj_ti" value="<%=jhdj_ti%>">
����ȼ�<br>
<input type="text" name="jhdj_nj" value="<%=jhdj_nj%>">
���Ҫ�ȼ�<br>
<input type="text" name="jhdj_jhtw" value="<%=jhdj_jhtw%>">
���������ȼ�<br>
<input type="text" name="jhdj_frag" value="<%=jhdj_frag%>">
��ҹ��ȼ�<br>
<input type="text" name="aqjh_jhcz" value="<%=aqjh_jhcz%>">
���ҵȼ�<br>
*****************������Ҫ���ǹ���ȼ�*****************<Br>
<input type="text" name="grade_dx" value="<%=grade_dx%>">
��Ѩ�����ȼ�<br>
<input type="text" name="grade_jx" value="<%=grade_jx%>">
��Ѩ�����ȼ�<br>
<input type="text" name="grade_db" value="<%=grade_db%>">
���������ȼ�<br>
<input type="text" name="grade_zl" value="<%=grade_zl%>">
���β����ȼ�<br>
<input type="text" name="grade_jj" value="<%=grade_jj%>">
��������ȼ�<br>
<input type="text" name="grade_jc" value="<%=grade_jc%>">
�������ȼ�<br>
<input type="text" name="grade_zs" value="<%=grade_zs%>">
ն�ײ����ȼ�<br>
<input type="text" name="grade_jg" value="<%=grade_jg%>">
��������ȼ�<br>
<input type="text" name="grade_see" value="<%=grade_see%>">
�鿴״̬<br>
<input type="text" name="grade_cf" value="<%=grade_cf%>">
���ȼ��̶�<br>
<input type="text" name="grade_gg" value="<%=grade_gg%>">
������ĵȼ�<br>
<input type="text" name="grade_tr" value="<%=grade_tr%>">
���˵ȼ�<br>
<input type="text" name="grade_fd" value="<%=grade_fd%>">
�Ŵ�����<br>
<input type="text" name="grade_ip" value="<%=grade_ip%>">
��ip�ȼ�<br>
<input type="text" name="grade_bo" value="<%=grade_bo%>">
��ը<br>
<input type="text" name="grade_jr" value="<%=grade_jr%>">
���ջ���<br>
<input type="text" name="grade_tjf" value="<%=grade_tjf%>">
ͨ�����ȼ�<br>
**********������****************************************<br>
<input type="text" name="pyyl" value="<%=pyyl%>">
���������Ҫ�Ľ����<br>
<input type="text" name="pyyls" value="<%=pyyls%>">
���������Ҫ��������<br>
<input type="text" name="xyyl" value="<%=xyyl%>">
���������Ҫ�����<br>
<input type="text" name="xyyls" value="<%=xyyls%>">
���������Ҫ������<br>
<input type="text" name="xzsj" value="<%=xzsj%>">
�Ƿ���ʱ������(trueΪ��,falseΪ��)<br>
<input type="text" name="xzsjs" value="<%=xzsjs%>">
�����ʱ������<br>
<br>
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
writeStr=writeStr & "'*********������Ҫ����ս���ȼ�*************************"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_qlcy="& request.form("jhdj_qlcy")&"      'ǧ�ﴫ���ȼ�"& chr(13)& chr(10)
writeStr=writeStr & "jhdj_duyao="& request.form("jhdj_duyao")&"      '�¶��ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_tq="& request.form("jhdj_tq")&"      '͵Ǯ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_xx="& request.form("jhdj_xx")&"      '���Ǵ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_aq="& request.form("jhdj_aq")&"      '�����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_ai="& request.form("jhdj_ai")&"      'ħ���ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_fz="& request.form("jhdj_fz")&"      '���й���"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_cs="& request.form("jhdj_cs")&"      '��������"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_sq="& request.form("jhdj_sq")&"      '��Ǯ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_so="& request.form("jhdj_so")&"      '�ͷ����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_zs="& request.form("jhdj_zs")&"      '���ͺ��Ķ��ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_pm="& request.form("jhdj_pm")&"      '�����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_bs="& request.form("jhdj_bs")&"      '��ʦ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_jj="& request.form("jhdj_jj")&"      '������ĵȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_zz="& request.form("jhdj_zz")&"      'ת����Ҫ�ĵȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_nh="& request.form("jhdj_nh")&"      'ŭ����Ҫ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_xt="& request.form("jhdj_xt")&"      '������Ҫ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_hy="& request.form("jhdj_hy")&"      '�������"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_db="& request.form("jhdj_db")&"      '�Ĳ���ׯ"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_qr="& request.form("jhdj_qr")&"      '���ˡ�������Ҫ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_dt="& request.form("jhdj_dt")&"      '������Ҫ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_chi="& request.form("jhdj_chi")&"      '��ǩ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_hua="& request.form("jhdj_hua")&"      '�ͻ��ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_zl="& request.form("jhdj_zl")&"      '���뿪�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_xq="& request.form("jhdj_xq")&"      '���鹦�ܵȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_ti="& request.form("jhdj_ti")&"      '����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_nj="& request.form("jhdj_nj")&"      '���Ҫ�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_jhtw="& request.form("jhdj_jhtw")&"      '���������ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "jhdj_frag="& request.form("jhdj_frag")&"      '��ҹ��ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "aqjh_jhcz="& request.form("aqjh_jhcz")&"      '���ҵȼ�"& chr(13) & chr(10)
writeStr=writeStr & "'*********������Ҫ���ǹ���ȼ�*************************"& chr(13) & chr(10)
writeStr=writeStr & "grade_dx="& request.form("grade_dx")&"      '��Ѩ�����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_jx="& request.form("grade_jx")&"      '��Ѩ�����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_db="& request.form("grade_db")&"      '���������ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_zl="& request.form("grade_zl")&"      '���β����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_jj="& request.form("grade_jj")&"      '��������ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_jc="& request.form("grade_jc")&"      '�������ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_zs="& request.form("grade_zs")&"      'ն�ײ����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_jg="& request.form("grade_jg")&"      '��������ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_see="& request.form("grade_see")&"      '�鿴״̬"& chr(13) & chr(10)
writeStr=writeStr & "grade_cf="& request.form("grade_cf")&"      '���ȼ��̶�"& chr(13) & chr(10)
writeStr=writeStr & "grade_gg="& request.form("grade_gg")&"      '������ĵȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_tr="& request.form("grade_tr")&"      '���˵ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_fd="& request.form("grade_fd")&"      '�Ŵ�����"& chr(13) & chr(10)
writeStr=writeStr & "grade_ip="& request.form("grade_ip")&"      '��ip�ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "grade_bo="& request.form("grade_bo")&"      '��ը"& chr(13) & chr(10)
writeStr=writeStr & "grade_jr="& request.form("grade_jr")&"      '���ջ���"& chr(13) & chr(10)
writeStr=writeStr & "grade_tjf="& request.form("grade_tjf")&"      'ͨ�����ȼ�"& chr(13) & chr(10)
writeStr=writeStr & "'*********������***********************************"& chr(13) & chr(10)
writeStr=writeStr & "pyyl="& request.form("pyyl")&"      '���������Ҫ�Ľ����"& chr(13) & chr(10)
writeStr=writeStr & "pyyls="& request.form("pyyls")&"      '���������Ҫ��������"& chr(13) & chr(10)
writeStr=writeStr & "xyyl="& request.form("xyyl")&"      '���������Ҫ�����"& chr(13) & chr(10)
writeStr=writeStr & "xyyls="& request.form("xyyls")&"      '���������Ҫ������"& chr(13) & chr(10)
writeStr=writeStr & "xzsj="& request.form("xzsj")&"      '�Ƿ���ʱ������"& chr(13) & chr(10)
writeStr=writeStr & "xzsjs="& request.form("xzsjs")&"      '�����ʱ������"
%>
<%
tpstr=replace(writestr,chr(13) & chr(10) ,"<br>")
tpstr="&lt;%<br>" & tpstr & "<br>%&gt;"
writestr="<" & "%" & chr(13) & chr(10) & writestr & chr(13) & chr(10) & "%" & ">"
response.write tpstr
if IsObjInstalled("Scripting.FileSystemObject") then
toppath = Server.Mappath("../chat/sjfunc/czdj.asp")
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
alert("���ķ�����֧��FSOȨ�ޣ��µ�czdj.asp���Զ�����")
</script>
<%
else
%>
<script>
alert("���ķ�������֧��FSOȨ�ޣ����ֹ��������븲��czdj.asp��ԭ���Ĵ���")
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