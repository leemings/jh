<%Response.Expires=0
if Session("aqjh_name")="" then 
	Response.Write "<script language=JavaScript>{alert('提示：系统已经超时，请重新登录！');top.location.href ='../exit.asp';}</script>"
	Response.End 
end if
if session("aqjh_adminok")<>true then
	Response.Write "<script language=JavaScript>{alert('提示：你没有登录管理,请进行登录操作！');top.location.href ='login.asp';}</script>"
	Response.End
end if
if Session("aqjh_grade")<10 or instr(Application("aqjh_admin"),Session("aqjh_name"))=0  then 
	Response.Write "<script language=JavaScript>{alert('提示：你无权进入系统管理选项！');window.close();}</script>"
	Response.End 
end if
%>
