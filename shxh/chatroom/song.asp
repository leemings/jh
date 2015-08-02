<%
set fs=server.CreateObject("scripting.filesystemobject")
set ffolder=fs.GetFolder(server.Mappath("../mid/song"))
set ffiles=ffolder.Files
dim sng()
i=0
for each fname in ffiles
	redim preserve sng(i)
	sng(i)=fname.name
	i=i+1
next
set ffiles=nothing
set ffolder=nothing
set fs=nothing
msg="<html><head><link rel='stylesheet' href='css.css'></head><body bgcolor='"&Application("Ba_jxqy_backgroundcolor")&"' background="&Application("Ba_jxqy_backgroundimage")&" ><p align=center><font color=0000FF size=4>ÌùÍ¼</font><hr></p>"
for i=0 to ubound(sng)
msg=msg&"<a href='#' onclick="&chr(34)&"javascript:parent.talkfrm.addtalk('[sng]"&sng(i)&"[/sng]');"&chr(34)&" onmouseover="&chr(34)&"window.status='"&sng(i)&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='"&sng(i)&"'>"&sng(i)&"</a><br>"
next
Response.Write msg&"</body></html>"
%>