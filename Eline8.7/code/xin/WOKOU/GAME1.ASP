<%If Session("usepro") = true Then%>
<%
if Session("sjjh_name")="" then response.end
my=Session("sjjh_name")
grade=Session("sjjh_grade")
myid=session("sjjh_id")
          id=request("id")
          win=request("win")
	  Set conn=Server.CreateObject("ADODB.CONNECTION")
	  Set rs=Server.CreateObject("ADODB.RecordSet")
	  conn.open Application("sjjh_usermdb")
          sql="select 内力,体力,内力加,体力加,金,木 from 用户 where 姓名='" & my &"'"
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
                              sql="update 用户 set 内力加=内力加+50,体力加=体力加+50,金=金+20,木=木+20 where 姓名='"&my&"' "
                              conn.execute sql  
                              %>
                              <script language=vbscript>
                         MsgBox "大侠，你打击倭寇有功，盟主奖内力和体力上限50，金加20，木加20！"
                         window.close()
                    </script><%
                     else 
                              sql="update 用户 set 内力=内力-1000,体力=体力-1000 where 姓名='"&my&"' "
                              conn.execute sql
                              %>
                              <script language=vbscript>
                         MsgBox "菜鸟，平时不练功，现在打败了，还敢来见我，罚你体力和内力1000！"
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