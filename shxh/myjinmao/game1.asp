<%If Session("usepro") = true Then%>
          <%
          
          id=request("id")
          win=request("win")
          my=session("Ba_jxqy_username")
          Set conn=Server.CreateObject("ADODB.CONNECTION")
          Set rs=Server.CreateObject("ADODB.RecordSet")
          connstr=Application("Ba_jxqy_connstr")
          conn.open connstr
          sql="select * from �û� where ����='" & my & "'"
          set rs=conn.execute(sql)
          if rs.eof or rs.bof then
	%>
                     <script language=vbscript>
                         MsgBox "�㲻�ǽ������ˣ����ܴ�����!"
                         window.close()
                    </script>
                    <%
                    	conn.close
                    	response.end
          else
	          if win=1 then
                              sql="select * from �û� where ����='" & my & "'"
                              conn.execute sql
                              sql="update �û� set ����=����+500,����=����+500,����=����+20 where ����='"&my&"' "
                              conn.execute sql  %>
                              <script language=vbscript>
                         MsgBox "���������������й�������������������500������20��"
                         window.close()
                    </script><%
                     else 
                              sql="select * from �û� where ����='" & my & "'"
                              conn.execute sql
                              sql="update �û� set ����=����-300,����=����-300,����=����-20 where ����='"&my&"' "
                              conn.execute sql
                              %>
                              <script language=vbscript>
                         MsgBox "����ƽʱ�����������ڴ���ˣ����������ң�����������Ŭ��300������20��"
                         window.close()
                    </script><%
  
                     end if
          end if
                    conn.close
                    set rs=nothing
   Session("usepro")= false 
                     %>
  <%else 
  response.write "��ͨ������;����������." 
  response.end 
  end if
 %>