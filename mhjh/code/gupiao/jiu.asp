<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->
<% 
if session("yx8_mhjh_username")="" then
response.redirect "../login.asp"
else
 	 
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then
call endinfo("��û���ʸ����")
else
  
sql = "select count(*) as num from �¼� "
set rs=conn.execute(sql)
num=rs("num")
if num>=40 then '*
Set rs = Server.CreateObject("ADODB.recordset")
sql = "select top 1 * from �¼� order by ԭ��ʱ�� asc"
rs.open sql,conn,3,2
rs.delete
end if '*

tt=now()
sql= "select * from ��Ʊ"
set rs=conn.execute(sql) 
DO UNTIL rs.EOF 
if (rs("��ǰ�۸�")-rs("���̼۸�"))/rs("���̼۸�")<0.15 then 
sql="update ��Ʊ set ��ǰ�۸�=��ǰ�۸�*1.04 where sid="&rs("sid")
conn.execute sql 
end if
rs.movenext
loop
mess="�������и�Ԥ�����Ʊ�г�Ͷ��޶��ʽ����й�Ʊ�۸����� 4 %"
sql="insert into �¼�(ԭ��,ԭ��ʱ��) values('"&mess&"','"&tt&"' )"
conn.execute sql  


 sub endinfo(message) 
'-------------------------------��Ϣ��ʾ-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>��Ϣ��ʾ</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="stock.asp">����</a></td></tr></table>
<%end sub

CloseDatabase		'�ر����ݿ�

Response.Redirect "stock.asp"

end if
end if
%>