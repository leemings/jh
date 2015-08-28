<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="jmconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=session("yx8_mhjh_username")
sjjh_grade=session("yx8_mhjh_usergrade")
sjjh_jhdj=session("yx8_mhjh_usergrade")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

yin=jmyin	'银两
id=request("id")
if id="" then
	response.write "参数错误！"
	response.end
end if
if not isnumeric(id) then
	response.write "<script language=JavaScript>{alert('错误：ID请使用数字');window.close();}</script>"
	response.end
end if
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("zgzm.asp")
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
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
	rs.open "select 银两 from 用户 where 姓名='"&sjjh_name&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "你不是江湖中人。"
		response.end
	end if
	newyin=rs("银两")-yin
	rs.close
	if newyin<0 then
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script language=javascript>{alert('周公替人解梦也不是免费的，一次收取银两："&yin&",你没这么多钱，不会给你解梦的。');window.close();}</script>"
		response.end
	end if
	conn.execute "update 用户 set 银两="&newyin&" where 姓名='"&sjjh_name&"'"
	kl="<font color=0000FF>"&sjjh_name&"</font>做了个怪梦：<font color=0000FF>【"&jmbt&"】</font>,在娱乐中心找到传说中的周公，花了"&yin&"两白银,终于得到了答案：<font color=red>"&jmsay&"！</font>"
	conn.close
end if
set rs=nothing
set conn=nothing
says="<font color=red><b>【周公解梦】</b></font>"&kl	
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=name
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
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
        <P align=center>游戏吧制作<br><br><br>
        </P>
      </TD><TD align=middle width=8 bgColor=#808080></TD></TR>
<TR style="FONT-SIZE: 3pt" align=middle><TD width=16 height=8></TD><TD bgColor=#808080 height=8></TD>
<TD width=8 bgColor=#808080 height=8></TD></TR></TBODY></TABLE></CENTER></body></html>
