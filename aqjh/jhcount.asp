<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("aqjh_data/aqjh_home.asp")

function boy()
     sql=conn.execute("select count(*) as boy from �û� where �Ա�='��'") 
     boy=sql(0) 
 set sql=nothing 
 if isnull(boy) then boy=0
end function 
function girl()
     sql=conn.execute("select count(*) as girl from �û� where �Ա�='Ů'") 
     girl=sql(0) 
 set sql=nothing 
 if isnull(girl) then girl=0
end function 
function mp()
     sql=conn.execute("select count(*) as mp from p where id") 
     mp=sql(0) 
 set sql=nothing 
 if isnull(mp) then mp=0
end function

function keeper()
     sql=conn.execute("select count(*) as keeper from �û� where ����='�ٸ�'") 
     keeper=sql(0) 
 set sql=nothing 
 if isnull(keeper) then keeper=0
end function
'����������
dim aass
call jsu()
function jsu()
sql=conn.execute("update [count] set manber=manber+1 where count='������'")
set sql=nothing
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "select * from [count] where count='������'",conn,1,3
sjs=split(rs("ndate"),"-")
sj=sjs(0)
sj=sj&"��"
sj=sj&sjs(1)
sj=sj&"��"
sj=sj&sjs(2)
sj=sj&"��"
aass=rs("manber")&"|"&rs("num")&"|"&rs("newuser")&"|"&sj
rs("ndate")=date()
rs.update
rs.close
set rs=nothing
end function
he=boy()+girl()
lsss=split(aass,"|")
%>
document.write("ע������:<%=he%> �����û�:<%=boy()%> Ů���û�:<%=girl()%> �����ʴ�:<%=lsss(0)%> �������:<%=lsss(1)%> ���ע��:<%=lsss(2)%> ������:<%=lsss(3)%>");
<%
conn.close
set rs=nothing
set conn=nothing
%>