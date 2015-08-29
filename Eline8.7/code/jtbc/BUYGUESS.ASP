<%@ LANGUAGE=VBScript codepage ="936" %>
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("sjjh_name")=""  then response.end
sjjh_name=Session("sjjh_name")
grade=Session("sjjh_grade")
guessnumber=request.form("guessnumber")
if guessnumber="" then response.redirect"../error.asp?id=484"
if not isnumeric(guessnumber) then response.redirect"../error.asp?id=484"
intguessnumber=guessnumber-int(guessnumber)
if gussnumber<0 or intguessnumber<>0 then response.redirect"../error.asp?id=484"
if len(guessnumber)<>5 then response.redirect"../error.asp?id=484"
guessdate=date
guess="DBQ="+server.mappath("guess.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set cn=Server.CreateObject("ADODB.CONNECTION")
Set rst=Server.CreateObject("ADODB.RecordSet")
cn.open guess
sql="select * from 购买者 where 号码="&guessnumber&" and datediff('d','"&guessdate&"',购买日期)=0 "
set rst=cn.execute(sql)
If Rst.Bof OR Rst.Eof Then
sql="select * from 购买者 where 购买者='"&sjjh_name&"' and datediff('d','"&guessdate&"',购买日期)=0 "
set rst=cn.execute(sql)
If Rst.Bof OR Rst.Eof Then
sql="SELECT 购买者 FROM 购买者 WHERE 购买者='"&sjjh_name&"'"
set rst=cn.Execute(sql)
If not(Rst.Bof OR Rst.Eof) Then
sql="update 购买者 set 号码="&guessnumber&",购买日期='"&guessdate&"' where 购买者='"&sjjh_name&"'"
set rst=cn.execute(sql)
'累计金加100
sql="update 中奖号码 set 累计金额=累计金额+1000"
set rst=cn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="update 用户 set 银两=银两-100 where 姓名='"&sjjh_name&"'"
set rs=conn.execute(sql)
response.redirect "buy.asp"
else
sql="INSERT into 购买者(购买者,号码,购买日期)values('"&sjjh_name&"',"&guessnumber&",'"&guessdate&"')"
Set Rst=cn.Execute(sql)
'累计金加100
sql="update 中奖号码 set 累计金额=累计金额+1000"
set rst=cn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="update 用户 set 银两=银两-100 where 姓名='"&usernmae&"'"
set rs=conn.execute(sql)
response.redirect "buy.asp"
end if
else
response.redirect "../error.asp?id=484"
end if
else
response.redirect "../error.asp?id=485"
end if
%>
