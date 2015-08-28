<%
function shoushi(un,co,mg)
if InStr(mg,"'")<>0 or InStr(mg,"`")<>0 or InStr(mg,"=")<>0 or InStr(mg,"-")<>0 or InStr(mg,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
elseif co<>"天网" then       
        Response.Write "<script language=JavaScript>{alert('你是杀手吗?不是的话收什么啊?');}</script>"
        Response.End
end if
  rst.Open "select * from  用户 where 姓名='"&mg&"'",conn
 if rst.eof or rst.bof then
  Response.Write "<script language=JavaScript>{alert('江湖里没这个人啊.正确格式为://收尸 你杀的人的名字');}</script>"
  Response.End
end if
rst.close
  rst.Open "select top 1 时间 from 英烈堂 where 凶手='"&un&"' and 死者='"&mg&"' order by 时间 desc",conn
shijian=rst("时间")
 if rst.eof or rst.bof then
   Response.Write "<script language=JavaScript>{alert('你还没有杀死你要杀的人');}</script>"
  Response.End
 end if
rst.close
 rst.Open "select * from s where ss='"&un&"' and bsr='"&mg&"' and wc='0'",conn
 if rst.eof or rst.bof then
	   Response.Write "<script language=JavaScript>{alert('要么不是你的追杀对象,要么你没有接受这单任务');}</script>"
        Response.End
      else
      yin=rst("bc")
      gy=rst("gyr")
      sj=rst("sj")
      ts=rst("ts")
      preday=dateadd("d",ts,sj)
     if shijian<sj then
       Response.Write "<script language=JavaScript>{alert('你还没有完成任务');}</script>"
        Response.End
     end if
    if preday>shijian then
     conn.execute "update 用户 set 银两=银两-"&yin&" where 姓名='"&gy&"'"
     conn.execute "update 用户 set 银两=银两+"&yin&" where 姓名='"&un&"'"
     conn.execute "update s set wc='1' where ss='"&un&"' and wc='0' and bsr='"&mg&"'"
     shoushi="【杀手收尸】##杀死了"&mg&"，开始收尸了，想不到##还真有两下子！"
    else
     Response.Write "<script language=JavaScript>{alert('你收尸过了追杀的最后时间了,该任务无效,无钱可收,去毁掉协议吧,这样你才能继续接活干');}</script>"
     Response.End
end if
end if
rst.close
end function
%>
