<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ����ܽ��������书"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
mp=rs("����")
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');window.close();}</script>"
	Response.End
end if
if rs("����")<500000 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������50��');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("����")<10000 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������1��~���߻���ħ������');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
gg=rs("����")
if rs("����")<10000 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������1��,���߻���ħ������');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
rs.close
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
yy=15+gg*2
conn.execute "update �û� set ����=����-10000,����=����-10000,����=����-500000,allvalue=allvalue+"&yy&",����ʱ��=now() where ����='"&aqjh_name&"'"
%>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>��Ͷ�ݸ��ˡ�["&aqjh_name&"]</font><font color=blue>������������ҹ�Ŀ���ʹ��������������1��,����50����������������������[<font color=red>����</font>]��Ӧָ���书��������<font color=red>["&aqjh_name&"]</font>֪�����ޣ���������ϣ�ֻ���Լ��ܻ�������<font color=red>"&yy&"</font>�㣬���ž��������񻰵����Ľ̵����վ�Ҳ�ܳ�Ϊ���ָ��֣���</font>"
call showchat(says)
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr>
          <td height="37"> <strong><font color="#800080">���ֽ�������--��������:</font></strong> 
        <tr>
          <td height="182" valign="top"><font color="#FF0000">������������ʥ�������书������8���ӵ�������С�гɾͣ�ս�������20�㣨���֣���</font><font color="#FF00FF"><Br>
            </font><Br>
          </td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="�� ��" onclick="location.href='index.asp'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
