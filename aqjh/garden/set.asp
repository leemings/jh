<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from ��԰����"
rs.open sql,connt,1,1
if rs.recordcount<>0 then
bm_start=rs("������ʼ")
bm_end=rs("��������")
tp_start=rs("ͶƱ��ʼ")
tp_end=rs("ͶƱ����")
tp_dj=rs("�ȼ�")
end if
%>
<HTML><HEAD>
<TITLE>����԰������</TITLE>
<link rel="stylesheet" href="../css.css">
</HEAD>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table width=500 border="0" align=center cellspacing="0" bordercolor="#FF0000">
        <tr bordercolor="#FFFFFF">
          <td align=center><font color=red><b>�� ԰ �� �� �� ��</b></font><BR><BR></td>
        </tr>
      </table>
      <table width=500 border="0" align=center cellspacing="0" bordercolorlight="#800080"
bordercolordark="#FFFFFF">
<tr><td><center>
<p>��ʽΪ����-��-�� ʱ:��:�롱��������2000-02-17 08:00:00����ȫ��Ϊ����ַ�����</p>
<p><a href=cl.asp><font color=red>��������</font></a>(ע�⣺����󽫲��ָܻ��������С��)</p>
</blockquote>
<table border="1" align="center" bgcolor="#E0E0E0" cellpadding="4" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0">
<form method="post" action="setok.asp">
<tr>
<td>ͶƱ����ȼ���</td>
<td>
<input type="text" name="tp_dj" maxlength="19" size="19" value="<%=tp_dj%>">
</td>
</tr>
<tr>
<td>��ʼ����ʱ�䣺</td>
<td>
<input type="text" name="bm_start" maxlength="19" size="19" value="<%=bm_start%>">
</td>
</tr>
<tr>
<td>��ֹ����ʱ�䣺</td>
<td>
<input type="text" name="bm_end" maxlength="19" size="19" value="<%=bm_end%>">
</td>
</tr><tr>
<td>��ʼͶƱʱ�䣺</td>
<td>
<input type="text" name="tp_start" maxlength="19" size="19" value="<%=tp_start%>">
</td>
</tr>
<tr>
<td>��ֹͶƱʱ�䣺</td>
<td>
<input type="text" name="tp_end" maxlength="19" size="19" value="<%=tp_end%>">
</td>
</tr>
<tr>
<td align="right">ѡ�������</td>
<td align="center">
<input type="submit" name="Submit" value="����">
<input type="button" value="����" onclick=javascript:history.go(-1)>
</td>
</form>
</table>
</td></tr></table></body></html>