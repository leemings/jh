<script lanaguage=javascript>
if(window.name!="aqjh_win"){ var i=1;while (i<=50){window.alert("������ʲôѽ���㵹��ˢ��������������50�Σ���");i=i+1;}top.location.href="../../exit.asp"}
</script>
<%If Session("usepro") = true Then%>
<%
if Session("aqjh_name")="" then response.end
my=Session("aqjh_name")
grade=Session("aqjh_grade")
myid=session("aqjh_id")
          id=request("id")
          win=request("win")
	  Set conn=Server.CreateObject("ADODB.CONNECTION")
	  Set rs=Server.CreateObject("ADODB.RecordSet")
	  conn.open Application("aqjh_usermdb")
          sql="select ����,����,������,������ from �û� where ����='" & my &"'"
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
                              sql="update �û� set ������=������+50,������=������+50 where ����='"&my&"' "
                              conn.execute sql  
                              %>
                              <script language=vbscript>
                         MsgBox "���������������й�����������������������50"
                         window.close()
                    </script><%
                     else 
                              sql="update �û� set ����=����-1000,����=����-1000 where ����='"&my&"' "
                              conn.execute sql
                              %>
                              <script language=vbscript>
                         MsgBox "����ƽʱ�����������ڴ���ˣ����������ң���������������1000��"
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