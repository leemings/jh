<%@ LANGUAGE=VBScript%>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
jifen=rs("allvalue")
myzs=rs("ת��")
ztdj=200+myzs*20
ztjf=ztdj^2*50
if jifen<ztjf or aqjh_jhdj<ztdj then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��������׼��"&myzs+1&"ת,��Ҫ"&ztdj&"������');window.close();}</script>"
		response.end
end if
neili=int(jifen*.5)
tili=int(jifen*.5)
wg=int(jifen*.25)
conn.execute "update �û� set ������=������+"&neili&",������=������+"&tili&",�书��=�书��+"&wg&",allvalue=allvalue-"&ztjf&",�ȼ�=�ȼ�-"&ztdj&",ת��=ת��+1 where ����='"&aqjh_name&"'"
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','����','����',now(),'ת��Ͷ̥')"
zstt=rs("ת��")
rs.close
set rs=nothing	
set conn=nothing
says="<font color=green>��ת��Ͷ̥��<font color=#ff0000>" & aqjh_name & "</b>����ת��Ͷ̥�ˣ������ӡ������Ӹ���"&neili&"�㣬�书������"&wg&"�㣡�����ǵ�"&myzs+1&"��ת��</font>"			'��������
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
call showchat(says)
Response.Write "<script language=JavaScript>{alert('��ʾ����ת��Ͷ̥���ˣ����ȷ�����µ�½������');top.location.href='../../exit.asp';}</script>"
response.end
%>