<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
%>
<head><title>¶Ä÷»×Ó</title></head>
<frameset rows="*,60" >
<frame name='imgfrm' src='pdimg.asp' noresize frameborder=no>
<frame name='optfrm' src='pdbet.asp' noresize frameborder=no>
</frameset>
