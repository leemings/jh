<%
'�����������
Sub Chkproxy()
if (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then
 call stopex("��ֹʹ�ô��������")
end if
End Sub
'һ�����
Sub Chkyjdh()
 aqjh_sid=trim(request.cookies("aqjh")("aqjh_sid"))
 if (aqjh_sid="" or aqjh_sid<>session.sessionid) and Session("aqjh_grade")<6 then
  Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('��һ̨��������˶���ʺţ���ϵͳ�����');}</script>"
  Response.End
 end if
End Sub
'վ���ύ����
sub ChkPost()
	dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		 call stopex("��ֹվ���ύ����")
	end if
end sub
'Myie�ര���ж�
sub Chkmyie()
response.write "<script>contents='<body>��ӭ�����ǣ��벻Ҫʹ��myie �ȶര���������½������!</body>';"
response.write "pop=window.open('','_blank','width=200,height=100');"
response.write "pop.document.writeln(contents);"
response.write "if(pop.document.body.clientWidth!=200||pop.document.body.clientHeight!=100){alert('������ʾ:�벻Ҫʹ��MyIe��MwIe�ȶര���������½������');top.location.href='exit.asp';}else{top.location.href='AQJH.HTM';}"
response.write "pop.close();"
response.write "</script>"
end sub

'==========������Ϣ
Sub stopEx(mess)
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&mess&"');location.href='javascript:history.go(-1)';}</script>"
	Response.End 
end sub
%>

