<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from ��Ʒ where ����='�鼮' and ӵ����='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("����")<=0 then 
			sql="delete * from ��Ʒ where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu9.asp"
			else
                     dd=rs("����")
			ml=rs("����")
			sql="update ��Ʒ set ����=����-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update �û� set ����=����+" & dd & " ,����=����+" & ml & " where ����='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                     Response.Redirect "xiaowu9.asp"
                     response.end
                                   conn.close
                                   set rs=nothing
		end if
%>