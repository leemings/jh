<%
username=session("yx8_mhjh_username")
if username="" then
response.write("对不起，你还没有<a href=../exit.asp>登录</a>")
response.end
end if
sex=request.form("sex")
if sex="" then
response.write("对不起，你还没有<a href=../exit.asp>登录</a>")
else
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 银两,体力,结算日期 FROM 用户 where 姓名='"&username&"' and 性别= '" &sex&"'"
Set Rs=conn.Execute(sql)
if rs.bof then
Response.Redirect "../error.asp?id=501"
else
if rs("银两")<50000 then Response.Redirect "../error.asp?id=502"
if day(rs("结算日期"))<day(now()) or month(rs("结算日期"))<month(now()) or year(rs("结算日期"))<year(now()) or isnull(rs("结算日期")) then
tini=50000
tempdate=date
timax=rs("体力")
tinow=rs("体力")
if tinow+tini>timax then tini=timax-tinow
sql="update 用户 set 银两=银两-50000,结算日期='"&tempdate&"',体力=体力+" & tini & " where 姓名='"&username&"'"
set rs=conn.execute(sql)
Response.Redirect "../ok.asp?id=504"
else
rs.close
Response.Redirect "../error.asp?id=503"
end if
end if
end if
rs.Close                                     
set rs=nothing                                     
conn.Close                                     
set conn=nothing    
%>


