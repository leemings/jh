<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
advertisemen=Application("Ba_jxqy_advertisemen")
if advertisemen="adrotator" then
	set ad=server.CreateObject("mswc.AdRotator")
	ad.border=0
	ad.TargetFrame="target=_blank"
	msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel=stylesheet href='css.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' topmargin=0 rightmargin=0 leftmargin=0 bottommargin=0><p align=center>"&Ad.GetAdvertisement("../adrotator.txt")&"</p></body>"
	set ad=nothing
elseif advertisemen="html" then
	Response.Redirect "../banner.htm"
else
	msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel=stylesheet href='css.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' topmargin=0 rightmargin=0 leftmargin=0 bottommargin=0>"&advertisemen&"</body>"
end if
Response.Write msg
%>