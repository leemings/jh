<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
cpass=Request.Form("cpass")
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 where 姓名='"&cpass&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要查找的人我们找不到！请查看是否正确！');history.back();</script>"
else
sql="update 用户 set 密码='e10adc3949ba59abbe56e057f20f883e' where 姓名='"&cpass&"'"
Set Rs=conn.Execute(sql)
Response.Write "<script language=javascript>alert('用户：["&cpass&"]的密码修改成123456成功！');history.back();</script>"
end if
%>