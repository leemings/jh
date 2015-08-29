<%
if Application("MaxShangpinlist")="" then
 Application.Lock
	Application("MaxSpecialList")="30"
 Application.UnLock
end if
	MaxSpecialList=Application("MaxSpecialList")
%>


