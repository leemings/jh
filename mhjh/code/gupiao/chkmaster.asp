<!--#include file="connbbs.asp"-->
<%
'�ж��Ƿ����Ա  ������  2002.10.11
	dim username
	username=session("yx8_mhjh_username")
	
        if session("yx8_mhjh_usercorp")="�ٸ�" and session("yx8_mhjh_usergrade")=120 then
           master=true
        else
	   master=false        
        end if 

%>
