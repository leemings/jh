<%@ LANGUAGE = VBScript %>
<!--#include file="jha.asp"-->
<%
nl=request.form("nl")
if nl=0 then
	mess="没有内力怎么出招?"
else
	Jname=session("Ba_jxqy_username")
	sql="select * from 用户 where 姓名='" & Jname & "' and 内力>" & nl & " and 体力>1000 and 状态<>'死' and 状态<>'狱'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		mess=Jname & "，你不能攻击,内力不够或者是体力不够1000"
	else
			randomize timer
			r=int(rnd*nl)
			m=int(rnd*800)
			rou=int(rnd*4)
			if r>m then
				mess=Jname&"打败了狗,得到防御" & rou & "点"
				sql="update 用户 set 内力=内力-" & nl & ",防御=防御+" & rou & " where 姓名='" & Jname & "'"
				conn.execute sql
			elseif r<m then
				mess=Jname&"被狼狗乱咬一阵，趴在地上"
				sql="update 用户 set 内力=内力-" & nl & ",体力=体力-" & m & " where 姓名='" & Jname & "'"
				conn.execute sql
			end if
        end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
%>
<TITLE>新手训练室</TITLE>
<style>td{font-size:14}</style>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=xuetang.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>