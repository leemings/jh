<!--#include file="connbbs.asp"-->
<%
'�ж��Ƿ����Ա  ������  2002.10.11
	dim membername
	membername=session("aqjh_name")
	
        if Session("aqjh_grade")=10 and instr(Application("aqjh_admin"),Session("aqjh_name"))<>0 then
           master=true
        else
	   master=false        
        end if 

%>
