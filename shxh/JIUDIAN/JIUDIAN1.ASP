<%
id=request("id")
my=session("Ba_jxqy_username")
%><!--#include file="dadata.asp"-->
<%
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	response.write "你不是剑侠中人，不能订购酒宴"
	conn.close
	response.end
else

sql="SELECT * FROM 宴会 where ID=" & id 
	Set Rs=connt.Execute(sql)
        wu=rs("宴会名")
	yin=rs("售价")
	lx=rs("类型")
        nl=rs("内力")
        tl=rs("体力")
        jb=rs("金币")
        sl=rs("数量")%>

<%
		sql="select * from 用户 where 姓名='" & my & "'"
		rs=conn.execute(sql)
		if yin<=rs("银两") then
			sql="update 用户 set 银两=银两-" & yin & " where 姓名='" & my & "'"
			rs=conn.execute(sql)%>
			
<%
			sql="select * from 宴会列表 where 宴会名='" & wu & "' and 拥有者='" & my & "'"
			set rs=connt.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into 宴会列表(宴会名,拥有者,类型,内力,体力,金币,数量,售价,时间) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now())"
			rs=connt.execute(sql)
                  
			connt.close
			Response.Redirect "jd.asp"
			else
            
				
					response.write "不能定酒宴，原因：你已定了同样规格的酒席！"
					connt.close
					response.end
				
			end if
		else
			response.write "不能定酒宴，原因：你的银两不够了"
			connt.close
			response.end
		end if
   rs.close
   set rs=nothing
end if
%>