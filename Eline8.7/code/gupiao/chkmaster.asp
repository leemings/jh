<!--#include file="connbbs.asp"-->
<%
'判断是否管理员  ♀wWw.happyjh.com♀
	dim membername
	membername=session("sjjh_name")
	
        if Session("sjjh_grade")=10 and instr(Application("sjjh_admin"),Session("sjjh_name"))<>0 then
           master=true
        else
	   master=false        
        end if 

%>
