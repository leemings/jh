<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="jmconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

yin=jmyin	'银两
jb=jmjb		'金币
id=request("id")
if id="" then
	response.write "参数错误！"
	response.end
end if
if not isnumeric(id) then
	response.write "<script language=JavaScript>{alert('错误：ID请使用数字');window.close();}</script>"
	response.end
end if
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("e_zgzm.asp")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from zgjm where id="&id,conn
if rs.eof or rs.bof then
	jmsay="没有找到你指定的梦类别。"
	lb="无"
else
	jmbt=rs("b")
	jmsay=rs("c")
	jmlb=rs("a")
	rs.close
	rs.open "select * from lb where id="&jmlb,conn
	if not (rs.eof or rs.bof) then
		lb=rs("lb")
	else
		lb="未分类"
	end if
end if
rs.close
conn.close
if lb<>"无" then
	conn.open application("sjjh_usermdb")
	rs.open "select 银两,金币 from 用户 where 姓名='"&sjjh_name&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "你不是江湖中人。"
		response.end
	end if
	newyin=rs("银两")-yin
	newjb=rs("金币")-jb
	rs.close
	if newyin<0 or newjb<0 then
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script language=javascript>{alert('周公替人解梦也不是免费的，一次收取银两："&yin&",金币："&jb&"，你没这么多钱，不会给你解梦的。');window.close();}</script>"
		response.end
	end if
	conn.execute "update 用户 set 银两="&newyin&",金币="&newjb&" where 姓名='"&sjjh_name&"'"
	conn.close
end if
set rs=nothing
set conn=nothing
%>
<html><head><title>周公解梦</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body background="bg.gif"><CENTER><TABLE cellSpacing=0 cellPadding=0 width=80% border=0><TBODY> 
<TR style="FONT-SIZE: 9pt"><TD colspan="2" bgColor=#D7EBFF></TD><TD width=8 height=16></TD></TR>
<TR><TD width=16 bgColor=#D7EBFF></TD><TD style="FONT-SIZE: 9pt; LINE-HEIGHT: 160%" bgColor=#D7EBFF><BLOCKQUOTE>
<P align=center class="3D"><B>周公解梦</B></P>
<BR>
&nbsp;<b>类别：</b><font color=red><%=lb%></font> <b>你梦到了：</b><font color=red><%=jmbt%></font>
<HR color=red SIZE=1><P style="word-break:break-all"><%=jmsay%></p>
<HR color=red SIZE=1><P align=center><input type=button value="关闭窗口" onclick="window.close()" class="input">
<br><br>
            <font color="#0000FF"><%=application("sjjh_chatroomname")%></font>之<font color="#FF0000">周公解梦</font></P>
        </BLOCKQUOTE>
        <P align=center>『E线江湖』&#8482; 2003-2004<br><br><br>
        </P>
      </TD><TD align=middle width=8 bgColor=#808080></TD></TR>
<TR style="FONT-SIZE: 3pt" align=middle><TD width=16 height=8></TD><TD bgColor=#808080 height=8></TD>
<TD width=8 bgColor=#808080 height=8></TD></TR></TBODY></TABLE></CENTER></body></html>
