<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
id=Request.QueryString("id")
uname=Request.QueryString("uname")
ucur=Request.QueryString("ucur")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
conn.Execute "delete  from ����  where id="&id&""
conn.Close
set conn=nothing
%>
<script language=javascript>
location.replace('showwubyname.asp?uname=<%=uname%>&ucur=<%=ucur%>');
</script>
