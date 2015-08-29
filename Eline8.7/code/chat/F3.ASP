<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
nowinroom=session("nowinroom")
online=Application("sjjh_onlinelist"&nowinroom)
onlineno=ubound(online)
%>
<script Language="JavaScript">
if(parent.document.URL.indexOf("file:")>=0||parent.f2.document.URL.indexOf("file:")>=0){parent.location.href='chaterr.asp?id=001';}
parent.md1(<%=onlineno%>);
<%
for o = 1 to onlineno 
Response.Write "parent.md2("&chr(34)& online(o) &chr(34)&");"& chr(13) & chr(10)
next
%>
parent.md3();
</script>