<%
'请杀手
function zuisha(un,st,co,gr,mg)
if InStr(mg,"'")<>0 or InStr(mg,"`")<>0 or InStr(mg,"=")<>0 or InStr(mg,"-")<>0 or InStr(mg,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
	exit function
end if
zt=split(mg,"|")
if ubound(zt)<>1 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为//追杀 你要杀的人的名字|需要在多少天内把他杀死(天数) 例//追杀 小花|5');}</script>"
	Response.End 
	exit function
end if
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，名字后面的天数请使用数字！');}</script>"
	Response.End 
	exit function
end if
    if co="十八地狱" then       
        Response.Write "<script language=JavaScript>{alert('官府的人还请杀手?');}</script>"
        Response.End
       exit function
    end if
if co="天网" then       
        Response.Write "<script language=JavaScript>{alert('你自己就是杀手啊');}</script>"
       Response.End
       exit function
    end if
  if gr<10 then
       Response.Write "<script language=JavaScript>{alert('你的等级太低了，需要10级才能请杀手！');}</script>"
       Response.End
       exit function
    end if
mz=trim(zt(0))
ts=abs(int(clng(zt(1))))
if ts=0 or ts>100 then
	Response.Write "<script language=JavaScript>{alert('提示：杀人天数应大于0天小于100天！');}</script>"
	Response.End 
	exit function
end if
rst.Open "select * from s where ss='"&st&"' and wc<>'1'",conn
if rst.eof or rst.bof then
rst.close
rst.Open "select 门派,等级,头像,攻击,防御 from  用户 where 姓名='"&mz&"'",conn
  if rst.eof or rst.bof then
Response.Write "<script language=JavaScript>{alert('江湖里没这个人啊.正确格式为://追杀 你要杀的人的名字|天数');}</script>"
Response.End
exit function 
  end if
 if rst("门派")="官府" then
Response.Write "<script language=JavaScript>{alert('官府的人是不能杀的哦');}</script>"
Response.End
exit function 
 end if
 if rst("门派")="天网" then
Response.Write "<script language=JavaScript>{alert('不可能叫他自己人杀自己人吧');}</script>"
Response.End
exit function 
 end if
 if rst("等级")<10 then
Response.Write "<script language=JavaScript>{alert('这个小于10级的新手也值得你杀吗?');}</script>"
Response.End
exit function 
 end if
tx=rst("头像")
gj=rst("攻击")
wg=rst("防御")
yl=(gj+wg)*5
rst.close
rst.Open "select 门派,道德,杀人 from  用户 where 姓名='"&st&"'",conn
 if rst("门派")<>"天网" then
       Response.Write "<script language=JavaScript>{alert('他是杀手吗?');}</script>"
       Response.End
       exit function
    end if
 if rst("道德")<10000 then            
	       Response.Write "<script language=JavaScript>{alert('他道德这么低你还请他?小心拿了你的钱就闪人哦');}</script>"
        Response.End
        exit function
   end if  
 if rst("杀人")>=10 then             
	       Response.Write "<script language=JavaScript>{alert('他杀人的数量已经超过了江湖规定的10个，所以你不能再请了!');}</script>"
        Response.End
        exit function
 end if  
 rst.close
 rst.Open "select 银两 from  用户 where 姓名='"&un&"' and 银两>"&yl&"",conn
	if rst.eof or rst.bof then
       Response.Write "<script language=JavaScript>{alert('你所杀的人不止这个价哦! 你的钱不够啊!所需银两=(对方攻击+防御)X5');}</script>"
       Response.End
       exit function
    end if      				
			 zuisha="【雇佣杀手】##从口袋里掏出"&yl&"两<img src=IMAGE/IMAGE/dz01.gif>对%%说:帮我杀一个人,这是他的照片<img src="&tx&">,<font color=blue>"&ts&"天</font>内杀了它这些钱就是你的<input type=button value='接受' onclick=""javascript:window.open('qssok.asp?un="&un&"','sendfrm','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button2231""><input type=button value='拒绝' onclick=""javascript:window.open('qssno.asp?un="&un&"','sendfrm','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223"">"	
conn.execute "insert into s(gyr,sj,bsr,ss,bc,wc,ts) values ('"& un &"',now(),'"& mz &"','"&st&"',"& yl &",'3',"&ts&")"

else
Response.Write "<script language=JavaScript>{alert('他还有任务没完成呢');}</script>"
Response.End
exit function 
end if
	rst.close
end function
%>
