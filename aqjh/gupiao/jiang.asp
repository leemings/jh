<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->
<% 
if   session("aqjh_name") ="" then
response.redirect "../login.asp"
else
 
if not master then
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
sql2="select count(*) as toto from ��Ʊ"   
set rsst=conn.execute(sql2)  
toto=rsst("toto")
rsst.close
for sj=1 to toto
sql= "select * from ��Ʊ"
set rs=conn.execute(sql) 
if (rs("��ǰ�۸�")-rs("���̼۸�"))/rs("���̼۸�")>-0.15 then 
sql="update ��Ʊ set ��ǰ�۸�=��ǰ�۸�*0.97 where sid="&sj
conn.execute sql 
end if
rs.movenext
next 
mess="���÷�չ���ȣ�ͨ���������أ���Ʊ�г����й�Ʊ�۸��´� 3 %"
sql="insert into �¼�(ԭ��,ԭ��ʱ��) values('"&mess&"','"&tt&"' )"
conn.execute sql  



	sj=formatdatetime(time(),3)
	sj="<font color=#FF00FF size=1>(" & sj & ")</font>"

CloseDatabase	
Response.Redirect "stock.asp"

end if
end if

 sub endinfo(message) 
'-------------------------------��Ϣ��ʾ-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>��Ϣ��ʾ</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="stock.asp">����</a></td></tr></table>
<%end sub
%>