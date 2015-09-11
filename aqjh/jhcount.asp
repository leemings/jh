<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("aqjh_data/aqjh_home.asp")

function boy()
     sql=conn.execute("select count(*) as boy from 用户 where 性别='男'") 
     boy=sql(0) 
 set sql=nothing 
 if isnull(boy) then boy=0
end function 
function girl()
     sql=conn.execute("select count(*) as girl from 用户 where 性别='女'") 
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
     sql=conn.execute("select count(*) as keeper from 用户 where 门派='官府'") 
     keeper=sql(0) 
 set sql=nothing 
 if isnull(keeper) then keeper=0
end function
'计数器处理
dim aass
call jsu()
function jsu()
sql=conn.execute("update [count] set manber=manber+1 where count='计数器'")
set sql=nothing
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "select * from [count] where count='计数器'",conn,1,3
sjs=split(rs("ndate"),"-")
sj=sjs(0)
sj=sj&"年"
sj=sj&sjs(1)
sj=sj&"月"
sj=sj&sjs(2)
sj=sj&"日"
aass=rs("manber")&"|"&rs("num")&"|"&rs("newuser")&"|"&sj
rs("ndate")=date()
rs.update
rs.close
set rs=nothing
end function
he=boy()+girl()
lsss=split(aass,"|")
%>
document.write("注册人数:<%=he%> 大侠用户:<%=boy()%> 女侠用户:<%=girl()%> 被访问次:<%=lsss(0)%> 被进入次:<%=lsss(1)%> 最后注册:<%=lsss(2)%> 最后更新:<%=lsss(3)%>");
<%
conn.close
set rs=nothing
set conn=nothing
%>