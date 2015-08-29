<%
if session("UserName")="" then
response.write"<script>alert('µÇÂ½³¬Ê±£¬ÇëÖØĞÂµÇÂ½£¡');location='../index.asp'</script>"
response.end
end if
%>