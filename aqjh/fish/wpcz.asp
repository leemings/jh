<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
err_num=0
wp_name=Replace(Trim(Request.Form("wp_name")), "'", "''")
'wp_name="小鱼"
if wp_name="" then 
Response.Write("<script language=""JavaScript"">alert ('没名称参数');window.location.href=""mywp.asp"";</script>")
Response.end
end if

if IsNumeric(Request.Form("wp_num")) then 
	wp_num=clng(Request.Form("wp_num"))
else
Response.Write("<script language=""JavaScript"">alert ('没数量参数');window.location.href=""mywp.asp"";</script>")
Response.end
end if

myid= Replace(Session("aqjh_name"), "'", "''")
'myid="asdf"
if myid="" then
Response.Write("<script language=""JavaScript"">alert ('没登入');window.location.href=""mywp.asp"";</script>")
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
		info="你购买了"&wp_name&"不存在"
		err_num=err_num+1
	else
		h=rs("h")
		info=wpadd_name&"单价"&h&"银两"
	end if
	rs.close
end if

if err_num=0 then
	rs.open "select * from 用户 where 姓名='"& myid &"'",conn,1,3
	If rs.Bof and rs.Eof then
		info="你购不是江湖中人"&info
		err_num=err_num+1
	else
		jh_yl=rs("银两")
		if (h*wpadd_sl)>jh_yl then
		info="你有钱"&jh_yl&"不够,需要"&(h*wpadd_sl>jh_yl)&"。"
		err_num=err_num+1
		else
		rs("银两")=rs("银两")-h*wpadd_sl
		rs.update
		info="你花了 "&(h*wpadd_sl)&" 两银子。"
		end if
	end if
	rs.close
end if
if err_num=0 then
	rs.open "select * from [wpname] where wp_user='" & wpadd_user & "' and wp_name='"& wpadd_name &"' and wp_sl>0" ,conn,1,3
	If rs.Bof and rs.Eof then
		rs.close
		'查找数据库中已删除记录使用删除亡不保存新物品
		rs.open "select * from [wpname] where wp_sl=0" ,conn,1,3
		If rs.Bof and rs.Eof then
			rs.addnew
		end if
		wpadd=wpadd_sl				'返回数量
		rs("wp_sl")=wpadd_sl
	else
		wpadd=rs("wp_sl")+wpadd_sl	'返回数量
		rs("wp_sl")=rs("wp_sl")+wpadd_sl
	end if
	rs("wp_user")=wpadd_user
	rs("wp_name")=wpadd_name
	rs.update
	rs.close
	info=info&"购买了 "&wp_name&wp_num&" 个。"
end if

set rs=nothing
conn.close
set conn=nothing

Response.Write("<script language=""JavaScript"">alert ('"&info&"');location.href=""mywp.asp"";</script>")
%>



</body>
</html>