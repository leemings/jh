<%
'检查代理服务器
Sub Chkproxy()
if (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then
 call stopex("禁止使用代理服务器")
end if
End Sub
'一机多号
Sub Chkyjdh()
 aqjh_sid=trim(request.cookies("aqjh")("aqjh_sid"))
 if (aqjh_sid="" or aqjh_sid<>session.sessionid) and Session("aqjh_grade")<6 then
  Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('您一台计算机上了多个帐号，被系统请出！');}</script>"
  Response.End
 end if
End Sub
'站外提交数据
sub ChkPost()
	dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		 call stopex("禁止站外提交数据")
	end if
end sub
'Myie多窗口判断
sub Chkmyie()
response.write "<script>contents='<body>欢迎朋友们，请不要使用myie 等多窗口浏览器登陆本江湖!</body>';"
response.write "pop=window.open('','_blank','width=200,height=100');"
response.write "pop.document.writeln(contents);"
response.write "if(pop.document.body.clientWidth!=200||pop.document.body.clientHeight!=100){alert('友情提示:请不要使用MyIe、MwIe等多窗口浏览器登陆本江湖');top.location.href='exit.asp';}else{top.location.href='AQJH.HTM';}"
response.write "pop.close();"
response.write "</script>"
end sub

'==========反馈信息
Sub stopEx(mess)
	Response.Write "<script language=JavaScript>{alert('提示："&mess&"');location.href='javascript:history.go(-1)';}</script>"
	Response.End 
end sub
%>

