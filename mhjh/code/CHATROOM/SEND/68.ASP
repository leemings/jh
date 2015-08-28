<%
function jiangli(un,st,co,mg)
    if not isnumeric(mg) then
     Response.Write "<script language=JavaScript>{alert('"&un&",请输入数字!');}</script>"
	 Response.End
	 exit function
	 end if
	 if mg<=0 then
      Response.Write "<script language=JavaScript>{alert('"&un&",请输入大于0的数字!');}</script>"
	 Response.End
	 exit function
	 end if
rst.Open "select 身份 from 用户 where 姓名='"&un&"'",conn
if not(rst.EOF or rst.BOF) then
sf=rst("身份")
if sf<>"掌门" and sf<>"副掌门" then 
jiangli="<font color=FF0000>【门派奖励】</font>,##不可以使用奖励功能，只有<font color=#000000><b>门派里的掌门，副掌门</b></font>才可以使用本功能" 
else
rst.Close
rst.Open "select 门派,身份 from 用户 where 姓名='"&st&"'",conn
menpai=rst("门派")
shenfen=rst("身份")
if co<>menpai then 
jiangli="<font color=FF0000>【门派奖励】</font>,##,你的门派管理范围也太大了吧，门派奖励只能对自己的门派弟子使用!"
elseif shenfen="掌门" or shenfen="副掌门" then
   jiangli="<font color=FF0000>【门派奖励】</font>,对方是门派中有头有脸的人物，##应该把机会留给别人吧!（掌门，副掌门你不能获得门派奖励）"
elseif mg>1000000 then 
   jiangli="<font color=FF0000>【门派奖励】</font>,想奖励" & st & "大洋" & mg & "两，但由于数目超过江湖规定，奖励失败！江湖规定奖励最大限制为100万！" 
else
rst.Close
rst.Open "select 基金 from 门派 where 门派='"&co&"'",conn
jj=rst("基金")
if jmg>jj then 
jiangli="<font color=FF0000>【门派奖励】</font>##,你们的门派基金<b>不够"&mg&"两银子！</b>" 
else
conn.execute "update 用户 set 银两=银两+"&mg&" where 姓名='"&st&"'"
conn.execute "update 门派 set 基金=基金-"&mg&" where 门派='"&co&"'"
jiangli="<font color=FF0000>【门派奖励】</font><font color=red><b>"&co&"的##认为%%确实做的不错，决定奖励</b></font>" & mg & "<font color=red><b>两白银，本门的门派库金减少"& mg &"两</b></font>" 
end if
end if
end if
end if
rst.Close
end function
%>
