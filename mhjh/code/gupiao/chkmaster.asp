<!--#include file="connbbs.asp"-->
<%
'判断是否管理员  我来了  2002.10.11
	dim username
	username=session("yx8_mhjh_username")
	
        if session("yx8_mhjh_usercorp")="官府" and session("yx8_mhjh_usergrade")=120 then
           master=true
        else
	   master=false        
        end if 

%>
