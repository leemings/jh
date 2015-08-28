<%
function buyquesiton(un,st)
       if Application("yx8_mhjh_answer"&chatroomsn)="stop"  then        
        rst.Open "select 银两,会员 from 用户 where 姓名='"&un&"' and 会员=true and 银两>"&Application("yx8_mhjh_asksilver"&chatroomsn)*2&"",conn
        if rst.eof or rst.bof then
        buyquesiton="<font color=FF0000>【买题】</font>对不起，##您的银两不够或者不是会员，不能买答案。"
        else
        buyquesiton="<b><font color=red>【买题】</font>上道题:<font color=balck>" & Application("yx8_mhjh_ask"&chatroomsn) & "</font>答案是:<font color=red>" &Application("yx8_mhjh_reply"&chatroomsn)&"</font>,##花费了<font color=red>"&Application("yx8_mhjh_asksilver"&chatroomsn)*2&"</font>两买下。</b>"
        conn.Execute"update 用户 set 银两=银两-"&Application("yx8_mhjh_asksilver"&chatroomsn)*2&" where 姓名='"&un&"'"
        rst.close
        end if
        else
        buyquesiton="<font color=FF0000>【买题】</font>答题时，##，您不能购买答案。" 
       end if        
end function
%>