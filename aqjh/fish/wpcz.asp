<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
err_num=0
wp_name=Replace(Trim(Request.Form("wp_name")), "'", "''")
'wp_name="С��"
if wp_name="" then 
Response.Write("<script language=""JavaScript"">alert ('û���Ʋ���');window.location.href=""mywp.asp"";</script>")
Response.end
end if

if IsNumeric(Request.Form("wp_num")) then 
	wp_num=clng(Request.Form("wp_num"))
else
Response.Write("<script language=""JavaScript"">alert ('û��������');window.location.href=""mywp.asp"";</script>")
Response.end
end if

myid= Replace(Session("aqjh_name"), "'", "''")
'myid="asdf"
if myid="" then
Response.Write("<script language=""JavaScript"">alert ('û����');window.location.href=""mywp.asp"";</script>")
Response.end
end if

wpadd_name=wp_name
wpadd_user=myid
wpadd_sl=wp_num
	wpadd=0
	Set conn = Server.CreateObject("ADODB.CONNECTION")
	conn.Open Application("aqjh_usermdb")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
if err_num=0 then	
	rs.open "select * from b where a='"& wpadd_name &"'",conn,1,1
	If rs.Bof and rs.Eof then
		info="�㹺����"&wp_name&"������"
		err_num=err_num+1
	else
		h=rs("h")
		info=wpadd_name&"����"&h&"����"
	end if
	rs.close
end if

if err_num=0 then
	rs.open "select * from �û� where ����='"& myid &"'",conn,1,3
	If rs.Bof and rs.Eof then
		info="�㹺���ǽ�������"&info
		err_num=err_num+1
	else
		jh_yl=rs("����")
		if (h*wpadd_sl)>jh_yl then
		info="����Ǯ"&jh_yl&"����,��Ҫ"&(h*wpadd_sl>jh_yl)&"��"
		err_num=err_num+1
		else
		rs("����")=rs("����")-h*wpadd_sl
		rs.update
		info="�㻨�� "&(h*wpadd_sl)&" �����ӡ�"
		end if
	end if
	rs.close
end if
if err_num=0 then
	rs.open "select * from [wpname] where wp_user='" & wpadd_user & "' and wp_name='"& wpadd_name &"' and wp_sl>0" ,conn,1,3
	If rs.Bof and rs.Eof then
		rs.close
		'�������ݿ�����ɾ����¼ʹ��ɾ��������������Ʒ
		rs.open "select * from [wpname] where wp_sl=0" ,conn,1,3
		If rs.Bof and rs.Eof then
			rs.addnew
		end if
		wpadd=wpadd_sl				'��������
		rs("wp_sl")=wpadd_sl
	else
		wpadd=rs("wp_sl")+wpadd_sl	'��������
		rs("wp_sl")=rs("wp_sl")+wpadd_sl
	end if
	rs("wp_user")=wpadd_user
	rs("wp_name")=wpadd_name
	rs.update
	rs.close
	info=info&"������ "&wp_name&wp_num&" ����"
end if

set rs=nothing
conn.close
set conn=nothing

Response.Write("<script language=""JavaScript"">alert ('"&info&"');location.href=""mywp.asp"";</script>")
%>



</body>
</html>