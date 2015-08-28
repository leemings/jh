<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
%>
<head><title>¶Ä÷»×Ó</title></head>
<frameset rows="*,60" framespacing="0" border="0" frameborder="0" >
<frame name='imgfrm' src='pdimg.asp' noresize frameborder=no scrolling="no">
<frame name='optfrm' src='pdbet.asp' noresize frameborder=no>
</frameset>
