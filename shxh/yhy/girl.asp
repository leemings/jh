
<%
nickname=Session("Ba_jxqy_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("Ba_jxqy_connstr")
conn.open connstr
sql="select * from �û� where ����='" & nickname & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	  Response.Redirect "../error.asp?id=016"
	conn.close
	response.end
else
dj=rs("�ȼ�")
yl=rs("����")
xingbie=rs("�Ա�")
if dj<3 then
	response.write "�㻹�ǽ���С�������������ֵط�"
	conn.close
	response.end
else
if xingbie="Ů" then
	response.write "����Ժ�Ĺ��ﲻ�Ӵ�Ů�ģ���ֹ����"
	conn.close
	response.end

else %>
<!--#include file="jiu.asp"-->
<% id=request("id")
sql="select * from ���� where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
response.write "����û�и��ѽ����Ժ�����������ѽ���ǲ�������������"
	connt.close
	response.end
else
mingji=rs("����")
meimao=rs("��ò��")

if yl<meimao*1.5  then
	response.write "ûǮҲ�������ָ߼��ĵط�ѽ����ֹ����"
            conn.close
	response.end
else
sql="update �û� set ����=����+"&meimao&"/2,����=����-100 where ����='"&nickname&"' "
conn.execute sql
sql1="update �û� set ����=����-"&meimao&"*1.5 where ����='"&nickname&"' "
conn.execute sql1
connt.close
conn.close


end if
end if
end if
end if
end if%>
<html>
<title>�������-�������</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>

<body bgcolor="#FFB366">
<p>��</p>
<table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="17" bgcolor="#996633" align="center">��</td>
  </tr>
  <tr bgcolor="#66FF66"> 
    <td align="center" height="378" bgcolor="#FFCC66"> 
      <p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400">
          <param name=movie value="girl.swf">
          <param name=quality value=high>
          <embed src="girl.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
          </embed></object><font> </font></p>
    </td>
  </tr>
  <tr bgcolor="#0033CC"> 
    <td align="center" height="26" class="unnamed1" bgcolor="#FFCC66"><b></b></td>
  </tr>
</table>
<p>��</p>
</body>
</html>