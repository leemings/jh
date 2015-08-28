<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
%>
<head><title>二十一点</title></head>
<frameset rows="*,60"  frameborder=0 framespacing="0" border="0">
<frameset cols="*,157"  frameborder=0>
<frame name='imgfrm' src='pcimg.asp' noresize frameborder=no framespacing=0 scrolling=no>
<frame name='cirfrm' src='net.asp' noresize frameborder=no framespacing=0 scrolling=no>
</frameset>
<frame name='optfrm' src='pcbet.asp' noresize frameborder=no framespacing=0 scrolling=no>
</frameset>
