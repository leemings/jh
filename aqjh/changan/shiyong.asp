<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from ��Ʒ where ����='������Ʒ' and ӵ����='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("����")<=0 then 
			sql="delete * from ��Ʒ where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu4.asp"
			else
			ti=rs("����")
			sql="update ��Ʒ set ����=����-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update �û� set ����=����+" & ti & " where ����='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                                      set rs=nothing
                     Response.Redirect "xiaowu4.asp"
                     response.end
                                    
		end if
%>