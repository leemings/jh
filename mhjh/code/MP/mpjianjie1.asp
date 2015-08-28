
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
pai=session("yx8_mhjh_usercorp")
username=session("yx8_mhjh_username")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from 用户 where 姓名='"&username&"' and 门派='"&pai&"' and 身份='掌门'",conn
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=046"
rs.Close
comment=request.form("comment")
nomment=request.form("nomment")
if InStr(comment,"=")<>0 or InStr(comment,"`")<>0 or InStr(comment,"'")<>0  or InStr(comment,"　")<>0 or InStr(comment,"'")<>0 or InStr(comment,chr(34))<>0 or InStr(comment,"\")<>0 or InStr(comment,"<")<>0 or InStr(comment,">")<>0 or InStr(comment,"&")<>0 then Response.Redirect "../error.asp?id=512"
if len(comment)>50 then Response.Redirect "error.asp?id=513"
if InStr(nomment,"=")<>0 or InStr(nomment,"`")<>0 or InStr(nomment,"'")<>0  or InStr(nomment,"　")<>0 or InStr(nomment,"'")<>0 or InStr(nomment,chr(34))<>0 or InStr(nomment,"\")<>0 or InStr(nomment,"<")<>0 or InStr(nomment,">")<>0 or InStr(nomment,"&")<>0 then Response.Redirect "../error.asp?id=512"
if len(nomment)>50 then Response.Redirect "error.asp?id=513"
sql="Update 门派 Set 公告='"&comment&"',简介='"&nomment&"' Where 门派='"&pai&"' "
conn.execute(sql)
set rs=nothing
Response.Redirect "mpjianjie.asp"
rs.close
conn.close
set rs=nothing
set conn=nothing
%>

