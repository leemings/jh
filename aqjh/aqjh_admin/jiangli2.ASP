<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
mp=trim(request.form("mp"))
zm=trim(request.form("zm"))
sm=trim(request.form("sm"))
xzsm=trim(request.form("xzsm"))
bds=trim(request.form("bds"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
' 初始化
number=trim(request.form("number"))
if number="" or not(isnumeric(number)) then response.write"奖励金额必须填数字"
xm=trim(request.form("xm"))
dj=trim(request.form("dj"))
sl=trim(request.form("sl"))

'打开数据库
sql="Select top "&sl&" "&xm&",姓名 from 用户 where "&dj&" order by mvalue desc"
rs.open sql,conn,3,2
'stop
do while not rs.eof
i=0
rs(xm)=rs(xm)+number
name=rs("姓名")
rs.update
conn.execute "insert into jl(a,b,c,d,e,f) values ('"&name&"',date(),'"&xm&"','"&number&"','"&aqjh_name&"','"&dj&"')"
rs.movenext
i=i+1
if i>sl then exit do
loop

rs.close
set rs=nothing
Response.Write "<script Language=Javascript>alert('提示：奖励完成！！');location.href = 'javascript:history.go(-1)';</script>"
%>

