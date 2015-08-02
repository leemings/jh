<%If Session("usepro") = true Then%>
          <%
          
          id=request("id")
          win=request("win")
          my=session("Ba_jxqy_username")
          Set conn=Server.CreateObject("ADODB.CONNECTION")
          Set rs=Server.CreateObject("ADODB.RecordSet")
          connstr=Application("Ba_jxqy_connstr")
          conn.open connstr
          sql="select * from 用户 where 姓名='" & my & "'"
          set rs=conn.execute(sql)
          if rs.eof or rs.bof then
	%>
                     <script language=vbscript>
                         MsgBox "你不是江湖中人，不能打倭寇!"
                         window.close()
                    </script>
                    <%
                    	conn.close
                    	response.end
          else
	          if win=1 then
                              sql="select * from 用户 where 姓名='" & my & "'"
                              conn.execute sql
                              sql="update 用户 set 内力=内力+500,体力=体力+500,银两=银两+20 where 姓名='"&my&"' "
                              conn.execute sql  %>
                              <script language=vbscript>
                         MsgBox "大侠，你打击倭寇有功，盟主奖内力和体力500，银两20个"
                         window.close()
                    </script><%
                     else 
                              sql="select * from 用户 where 姓名='" & my & "'"
                              conn.execute sql
                              sql="update 用户 set 内力=内力-300,体力=体力-300,银两=银两-20 where 姓名='"&my&"' "
                              conn.execute sql
                              %>
                              <script language=vbscript>
                         MsgBox "菜鸟，平时不练功，现在打败了，还敢来见我，罚你体力和努力300，银两20个"
                         window.close()
                    </script><%
  
                     end if
          end if
                    conn.close
                    set rs=nothing
   Session("usepro")= false 
                     %>
  <%else 
  response.write "请通过正常途径来打倭寇." 
  response.end 
  end if
 %>