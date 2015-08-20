<%
username=session("Ba_jxqy_username")
session.Abandon
Application("Ba_jxqy_allonline")=replace(Application("Ba_jxqy_allonline"),";"&username&";",";")
Application("Ba_jxqy_allonlinenum")=0
%>
<script language=javascript>top.location.replace('index.asp');</script>