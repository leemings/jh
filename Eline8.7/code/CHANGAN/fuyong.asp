<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="SELECT * FROM 用户 WHERE id="&sjjh_jhid
              Set Rs=conn.Execute(sql)
              if rs("内力")>=rs("grade")*500000 or rs("体力")>=rs("grade")*500000 then %>
<script language=vbscript>
MsgBox "你吃的太多了，小心撑着"
location.href = "xiaowu3.asp"
</script>
<%else
              sql="select * from 物品 where 类型='食品' and 拥有者='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("数量")<=0 then 
			sql="delete * from 物品 where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu3.asp"
			else
			ti=rs("体力")
			sql="update 物品 set 数量=数量-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update 用户 set 体力=体力+" & ti & " where id="&sjjh_jhid
			set rs=conn.execute(sql)
                     conn.close
                     Response.Redirect "xiaowu3.asp"
                     response.end
    conn.close
    set rs=nothing
		end if
end if 
%>
<html><script language="JavaScript">                                                                  </script></html>