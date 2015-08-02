<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
id=Request.QueryString("id")
uname=Request.QueryString("uname")
ucur=Request.QueryString("ucur")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
conn.Execute "update 物品 set 数量=0 where id="&id
conn.Close
set conn=nothing
%>
<script language=javascript>
location.replace('showcurbyname.asp?uname=<%=uname%>&ucur=<%=ucur%>');
</script>
