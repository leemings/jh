<%
function reply(un,st,mg)
        if Application("yx8_mhjh_answer"&chatroomsn)<>"stop" then
        if  cint(datediff("n",application("questiontime"&chatroomsn),now()))<=1 then        
           if instr(1,application("answer"&chatroomsn),Session("yx8_mhjh_username"))>0 then
            reply="<font color=FF0000>【答题】</font>##您已经答过题了，不能再答！" 
           else
        silver=Application("yx8_mhjh_asksilver"&chatroomsn)
        application("answer"&chatroomsn)=application("answer"&chatroomsn)&Session("yx8_mhjh_username")&" "
        if mg=Application("yx8_mhjh_reply"&chatroomsn) then
       if Application("yx8_mhjh_askqiang"&chatroomsn)=true then
  	application("answername"&chatroomsn)=application("answername"&chatroomsn)&Session("yx8_mhjh_username")&" "  	        
        rst.Open "select 银两,道德 from 用户 where 姓名='"&un&"'",conn        
        conn.Execute"update 用户 set 银两=银两+'"&silver&"',道德=道德+10 where 姓名='"&un&"'"       
		rst.close        
        reply="<font color=FF0000>【答题】</font>##感谢您参与这次的答题，剑侠欢迎你"       
        Application("yx8_mhjh_answer"&chatroomsn)="stop"
        Application("yx8_mhjh_answerqiang"&chatroomsn)="ok"
        else        application("answername"&chatroomsn)=application("answername"&chatroomsn)&Session("yx8_mhjh_username")&" "  	        
        rst.Open "select 银两,道德 from 用户 where 姓名='"&un&"'",conn        
        conn.Execute"update 用户 set 银两=银两+'"&silver&"',道德=道德+3 where 姓名='"&un&"'"       
		rst.close
        reply="<font color=FF0000>【答题】</font>##感谢您参与这次的答题，剑侠欢迎你" 
        end if  
		else
        reply="<font color=FF0000>【答题】</font>##感谢您参与这次的答题，剑侠欢迎你" 		
        end if
          end if
        else
        reply="<font color=FF0000>【答题】</font>答题时间是一分钟，##您已经过了答题时间" 
        end if
      else
      reply="<font color=FF0000>【答题】</font>##现在暂时无题可答，请稍候！"           
      end if
end function
%>