<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
p=Trim(Request("id"))
if instr(p,"or")<>0 or instr(p,"=")<>0 then
	Response.Redirect "../error.asp?id=54"
	response.end
end if
if left(p,3)="%20" OR InStr(p,"=")<>0 or InStr(p,"`")<>0 or InStr(p,"'")<>0 or InStr(p," ")<>0 or InStr(p,"��")<>0 or InStr(p,"'")<>0 or InStr(p,chr(34))<>0 or InStr(p,"\")<>0 or InStr(p,",")<>0 or InStr(p,"<")<>0 or InStr(p,">")<>0 then
	Response.Redirect "../error.asp?id=120"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT id FROM �û� WHERE  ����='" & sjjh_name &"'",conn,2,2
sjjh_id=rs("id")
rs.close
sql="select * from ͶƱ"
set rs=conn.execute(sql)
ks=rs("��ʼʱ��")
jz=rs("��ֹʱ��")
dj=rs("�ȼ�")
rs.close
if now()<ks then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>ͶƱ���δ��ʼ������ͶƱ��"
	Response.End
end if
if now()>jz then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>ͶƱ��ѽ���������ͶƱ��"
	Response.End
end if
if sjjh_jhdj<dj then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>���ĵȼ�û�дﵽ" & dj & "��������ͶƱ��"
	Response.End
end if
if p="" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>��ѡ������֧�ֵĽ���Ӣ�ۡ�<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.End
end if
sql="select ���� from ��ѡ�� where ����='" & p & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>��֧�ֵĽ���Ӣ�ۣ�<font color=red>" & p & "</font> �����ڡ�<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.End
end if
rs.close
sql="select * from ͶƱ�� where ͶƱID="&sjjh_id
set rs=conn.execute(sql)
if not(rs.eof or rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>���Ѿ�Ͷ��Ʊ�ˣ�������Ͷ�ˡ�<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.End
end if
rs.close
sql="update ��ѡ�� set ��Ʊ��=��Ʊ��+1 where ����='" & p & "'"
set rs=conn.execute(sql)
sql="insert into ͶƱ��(ͶƱID,����) values ('" & sjjh_id & "','" & sjjh_name & "')"
set rs=conn.execute(sql)
set rs=nothing
conn.close
set conn=nothing
%><html>
<head>
<title>ͶƱ��wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
<!--
.p9 {line-height: 150%; font-size: 9pt;}
.p12 {line-height: 150%; font-size: 12pt;}
body {line-height: 150%;font-size : 12pt;}
td {line-height: 135%;font-size : 10.5pt;}
A  {text-decoration: none;}
A:Hover  {text-decoration : none;}
a:visited {  color: #0000FF}
-->
</style>
<script language="JavaScript">
function load()
{var name=navigator.appName
var vers=navigator.appVersion
if(name=="Netscape")
{window.location.reload()
}else
{history.go(0)
}}</script>
</head>
<body bgcolor="#FFFFEC" topmargin="0" leftmargin="0">
<table border="1" cellpadding="2" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1">
<tr>
    <td width="100%" height="1"><font size="2">����λ��&gt;&gt;<font color="#0000FF"><a href="poll.asp">����ͶƱ</a></font>&gt;&gt;</font><a href="pollview.asp">ͶƱ/�鿴ͶƱ���</a>&gt;&gt;<font color="#0000FF">ͶƱ���</font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">��ͶƱ�����</font></h1>
<hr noshade size="1" color=009900>
<blockquote>
<p><b><font color="#FF0000">[�������]</font></b></p>
<blockquote>
<p><font size="2">��л��Ͷ������ʥ��һƱ����֧�ֵ��ǣ�<font color="#FF0000"><%=p%></font> <a href=javascript:history.go(-1)>�����ء�</a></font></p>
</blockquote>
</blockquote>
<hr noshade size="1" color=009900>
</body>
</html>

