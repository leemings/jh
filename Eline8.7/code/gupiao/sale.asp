<!--#include file="conn.asp"-->
<% 

if session("sjjh_name")="" then
response.redirect "../login.asp"
response.end
else
%>
 
<%sid=Request.Form ("sid")
ushare=abs(Request.Form ("ushare2"))
sql= "select * from ��Ʊ where sid="&sid        
set rs=conn.execute(sql) 
if rs("��ǰ�۸�")<=1  or (rs("��ǰ�۸�")-rs("���̼۸�"))/rs("���̼۸�")<=-0.15 then
call endinfo("����ͣ�ơ���ͣ����ߵ�ͣ��Ĺ�Ʊ�ǲ��ܽ�����������Ŷ����")

else
set rs=nothing
 
session("sjjh_name")=session("sjjh_name")
sql="select * from �ͻ� where �ʺ�='"&session("sjjh_name")&"'"
set rs=conn.execute(sql)
username=rs("�ʺ�")
nowmoney=rs("�ʽ�")

set rs_s=conn.execute ("select ���̼۸�,��ͨ��Ʊ,��ǰ�۸�,��ҵ from ��Ʊ where sid="&sid)
set rs_u=conn.execute ("select �ֹ���,����۸�,ƽ���۸� from �� where sid="&sid&" and  �ʺ�='"&username&"'")

dsshare=rs_s("��ͨ��Ʊ")+ushare
nowp=rs_s("��ǰ�۸�")

addmoney=ushare*rs_s("��ǰ�۸�")
nowp=rs_s("��ǰ�۸�")
tot=rs_s("��ͨ��Ʊ")

sql = "select count(*) as num from �¼� "
set rs=conn.execute(sql)
num=rs("num")
if num>=40 then '*
Set rs = Server.CreateObject("ADODB.recordset")
sql = "select top 1 * from �¼� order by ԭ��ʱ�� asc"
rs.open sql,conn,3,2
rs.delete
end if '*

if rs_u.eof then 
call endinfo("��û�����ֹ�Ʊ��")
else
dushare=rs_u("�ֹ���")-ushare
if dushare<0 then
call endinfo("���Ĺ�Ʊ�����㣡")
response.end

elseif dushare=0 then
Randomize
s=Rnd


sprice=nowp*(1-ushare/(tot*10))


shou=(rs_s("��ǰ�۸�")-rs_u("ƽ���۸�"))*ushare
conn.execute "update ��Ʊ set ����="&date()&",������=������+"&ushare&",��ǰ�۸�="&sprice&", ��ͨ��Ʊ="&dsshare&" where sid="&sid
conn.execute "delete from �� where �ʺ�='"&username&"' and sid="&sid
sql="update �ͻ� set �ʽ�=�ʽ�+"&addmoney&"*0.99 where �ʺ�='"&username&"'"
conn.execute sql
Set rs2 = Server.CreateObject("ADODB.recordset")
sql2="select ���ʽ� from ���ʽ� where �ʺ�='"&username&"'"
rs2.open sql2,conn
if (rs2("���ʽ�")-addmoney*1.01)<=0 then 
sql="update ���ʽ� set ���ʽ�=0 where  �ʺ�='"&username&"'"
conn.execute sql
else
sql="update ���ʽ� set ���ʽ�=���ʽ�-"&addmoney&"*1.01 where  �ʺ�='"&username&"'"
conn.execute sql
end if
rs2.close
set rs2=nothing

zd=ccur((formatcurrency(nowp,3)-formatcurrency(sprice,3))/formatcurrency(rs_s("��ǰ�۸�"),3))
mess="<font color=#006600>"&username&"����"&rs_s("��ҵ")&" "&ushare&" �ɣ��ּ��»� "&formatpercent(zd,2,-1)&"</font>"

sql="insert into �¼�(ԭ��,ԭ��ʱ��) values('"&mess&"','"&now()&"' )"
conn.execute sql  
else
Randomize
s=Rnd

sprice=nowp*(1-ushare/(tot*10))

shou=(rs_s("��ǰ�۸�")-rs_u("ƽ���۸�"))*ushare
conn.execute "update ��Ʊ set ����="&date()&",������=������+"&ushare&",��ǰ�۸�="&sprice&", ��ͨ��Ʊ="&dsshare&" where sid="&sid
conn.execute "update �� set �ֹ���="&dushare&" where �ʺ�='"&username&"' and sid="&sid
sql="update �ͻ� set �ʽ�=�ʽ�+"&addmoney&"*0.99 where �ʺ�='"&username&"'"
conn.execute sql

sql2="select ���ʽ� from ���ʽ� where �ʺ�='"&username&"'"
Set rs2 = Server.CreateObject("ADODB.recordset") 
rs2.open sql2,conn
if (rs2("���ʽ�")-addmoney*1.01)<=0 then 
sql="update ���ʽ� set ���ʽ�=0 where  �ʺ�='"&username&"'"
conn.execute sql
else
sql="update ���ʽ� set ���ʽ�=���ʽ�-"&addmoney&"*1.01 where  �ʺ�='"&username&"'"
conn.execute sql
end if
rs2.close
set rs2=nothing

zd=ccur((formatcurrency(nowp,3)-formatcurrency(sprice,3))/formatcurrency(rs_s("��ǰ�۸�"),3))
mess="<font color=#006600>"&username&"����"&rs_s("��ҵ")&" "&ushare&" �ɣ��ּ��»� "&formatpercent(zd,2,-1)&" </font>"

sql="insert into �¼�(ԭ��,ԭ��ʱ��) values('"&mess&"','"&now()&"' )"
conn.execute sql  

end if
call endinfo("��Ʊ���������ѳɹ��ύ��ȷ���󷵻�")
end if
end if
end if

 sub endinfo(message) 
'-------------------------------��Ϣ��ʾ-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>��Ϣ��ʾ</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="stock.asp">����</a></td></tr></table>
<%end sub
%>
<%conn.close
set conn=nothing

%>