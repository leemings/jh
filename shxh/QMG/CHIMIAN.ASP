<%
name=session("Ba_jxqy_username")
%>
<!--#include file="dadata.asp"-->
<% sql="select * from 用户 where 姓名='" & name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	response.write "你不是剑侠中人，不能吃面"
	conn.close
	response.end

else    
js=rs("积分")    
sql="select * from 巧面 where 拥有者='" & name & "'"
		set rs=connt.execute(sql)
              
if rs.eof or rs.bof then%>

<script language=vbscript>
 MsgBox "你没有面啊，是不是想骗我，快去面馆作面吧。"
 location.href = "qmg.htm"
 </script>

<%     connt.close
	response.end

else

mw=rs("美味度")

zjtl=mw*10+js
conn.execute("update 用户 set 体力=体力+" & zjtl & "  where 姓名='"&name&"'")
conn.execute("update 用户 set 积分=积分 +1 where 姓名='"&name&"'")


connt.execute("delete * from 巧面 where 拥有者='" & name & "'" )


Response.write"你一口气把香气扑鼻的面吃下去。增加体力" & zjtl& ""




rs.close
conn.close
connt.close%>
<%end if%>
<%end if%>