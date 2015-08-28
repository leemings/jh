<%@ LANGUAGE=VBScript%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
nickname=session("yx8_mhjh_username")
if nickname="" then response.end
from1=LCase(trim(request.querystring("from")))
xiazhu=LCase(trim(request.querystring("xiaozhu")))
jiayin=int(xiazhu*0.9)
to1=LCase(trim(request.querystring("to1")))
yn=LCase(trim(request.querystring("yn")))
if InStr(from1,"or")<>0 or InStr(from1,"'")<>0 or InStr(from1,"`")<>0 or InStr(from1,"=")<>0 or InStr(from1,"-")<>0 or InStr(from1,",")<>0 or InStr(from1,"&")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
if InStr(to1,"or")<>0 or InStr(to1,"'")<>0 or InStr(to1,"`")<>0 or InStr(to1,"=")<>0 or InStr(to1,"-")<>0 or InStr(to1,",")<>0 or InStr(to1,"&")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
if InStr(xiazhu,"or")<>0 or InStr(xiazhu,"'")<>0 or InStr(xiazhu,"`")<>0 or InStr(xiazhu,"=")<>0 or InStr(xiazhu,"-")<>0 or InStr(xiazhu,",")<>0 or InStr(xiazhu,"&")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
if InStr(yn,"or")<>0 or InStr(yn,"'")<>0 or InStr(yn,"`")<>0 or InStr(xiazhu,"=")<>0 or InStr(yn,"-")<>0 or InStr(yn,",")<>0 or InStr(yn,"&")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
if nickname<>to1 then
      msg="人家没有邀请你呀！"
      abc=0
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from d where xm='"&from1&"' and df='"&nickname&"' and yl="&xiazhu&" and c='对'"
set rs=conn.execute(sql)
if rs.EOF or rs.BOF then
      msg="没有此对局，也有可能时间已经过了！"
      abc=0
      conn.close
      set conn=nothing
      set rs=nothing
else      
if yn=0 then
       msg="你拒绝了"&from1&"的对赌！"
       conn.execute "delete * from d where yl="& xiazhu &" and xm='"&from1&"' and df='"&nickname&"' and c='对'"
       conn.close
       set conn=nothing
       set rs=nothing
       abc=1
       duidu="<b>"&nickname&"</b>拒绝了<b>"&from1&"</b>的赌博挑战！"
else
       sql="select 姓名,银两 from 用户 where 银两>"&xiazhu&" and 姓名='"&from1&"'"
       set rs=conn.execute(sql)
       if rs.EOF or rs.BOF then
       conn.execute "delete * from d where yl="& xiazhu &" and xm='"&from1&"' and df='"&nickname&"' and c='对'"
       conn.close
       set conn=nothing
       set rs=nothing
       duidu=""&from1&"银两不够，赌博挑战失败！"
       msg=""&from1&"银两不够，赌博挑战失败！"
       abc=1
       else
        sql="select 姓名,银两 from 用户 where 银两>"&xiazhu&" and 姓名='"&nickname&"'"
        set rs=conn.execute(sql)
       if rs.EOF or rs.BOF then
       conn.execute "delete * from d where yl="& xiazhu &" and xm='"&from1&"' and df='"&nickname&"' and c='对'"
       conn.close
       set conn=nothing
       set rs=nothing
       duidu=""&nickname&"银两不够，赌博挑战失败！"
       msg=""&nickname&"银两不够，赌博挑战失败！"
       abc=1
       else
       conn.execute "delete * from d where yl="& xiazhu &" and xm='"&from1&"' and df='"&nickname&"' and c='对'"
      	Randomize
	f6 = Int(6 * Rnd + 1)
        Randomize
	f1 = Int(6 * Rnd + 1)
	Randomize
	f4 = Int(6 * Rnd + 1)
	Randomize
	f3 = Int(6 * Rnd + 1)
	Randomize
	f5 = Int(6 * Rnd + 1)
        Randomize
	f2 = Int(6 * Rnd + 1)
        jiutian_duiduf=f1+f2+f3
	jiutian_duidut=f4+f5+f6
if f1=f2=f3 and f4=f5=f6 then
        abc=1
        msg="你同意"&from1&"的对赌！"
        duidu="["&nickname&"]开：<img src=../chatroom/image/"&f4&".gif><img src=../chatroom/image/"&f5&".gif><img src=../chatroom/image/"&f6&".gif>计："& jiutian_duidut &"点!开出豹子！["&from1&"]开：<img src=../chatroom/image/"&f1&".gif><img src=../chatroom/image/"&f2&".gif><img src=../chatroom/image/"&f3&".gif>计："& jiutian_duiduf &"点!开出豹子！平局！！"
set conn=nothing
set rs=nothing
else
if f1=f2=f3 then
msg="你同意"&from1&"的对赌！"
abc=1
duidu="["&nickname&"]开：<img src=../chatroom/image/"&f4&".gif><img src=../chatroom/image/"&f5&".gif><img src=../chatroom/image/"&f6&".gif>计："& jiutian_duidut &"点!！["&from1&"]开：<img src=../chatroom/image/"&f1&".gif><img src=../chatroom/image/"&f2&".gif><img src=../chatroom/image/"&f3&".gif>计："& jiutian_duiduf &"点!开出豹子！["&from1&"]赢，赢了"&xiazhu&"两！胜者扣10%官税！"
                conn.execute "update 用户 set 银两=银两+'"&jiayin&"' where 姓名='"&from1&"'"
		conn.execute "update 用户 set 银两=银两-'"&xiazhu&"' where 姓名='"&nickname&"'"	
		conn.close
                set conn=nothing
                set rs=nothing
		else
		    if f4=f5=f6 then
		     msg="你同意"&from1&"的对赌！"
		    abc=1
		    duidu="["&nickname&"]开：<img src=../chatroom/image/"&f4&".gif><img src=../chatroom/image/"&f5&".gif><img src=../chatroom/image/"&f6&".gif>计："& jiutian_duidut &"点!开出豹子！["&from1&"]开：<img src=../chatroom/image/"&f1&".gif><img src=../chatroom/image/"&f2&".gif><img src=../chatroom/image/"&f3&".gif>计："& jiutian_duiduf &"点!["&nickname&"]从["&from1&"]赢到"&xiazhu&"两银子！胜者扣10%官税！"
                   conn.execute "update 用户 set 银两=银两-'"&xiazhu&"' where 姓名='"&from1&"'"
		           conn.execute "update 用户 set 银两=银两+'"&jiayin&"' where 姓名='"&nickname&"'"	
                    conn.close
                      set conn=nothing
                      set rs=nothing
                      else 
                      if jiutian_duiduf=jiutian_duidut then
                       msg="你同意"&from1&"的对赌！"
                       abc=1
                      duidu="["&nickname&"]开：<img src=../chatroom/image/"&f4&".gif><img src=../chatroom/image/"&f5&".gif><img src=../chatroom/image/"&f6&".gif>计："& jiutian_duidut &"点! ["&from1&"]开：<img src=../chatroom/image/"&f1&".gif><img src=../chatroom/image/"&f2&".gif><img src=../chatroom/image/"&f3&".gif>计："& jiutian_duiduf &"点!点数相同，平局！"            
                      conn.close
                      set conn=nothing
                      set rs=nothing
                       else 
                        if jiutian_duiduf<jiutian_duidut then
                         msg="你同意"&from1&"的对赌！"
                        abc=1
                        duidu="["&nickname&"]开：<img src=../chatroom/image/"&f4&".gif><img src=../chatroom/image/"&f5&".gif><img src=../chatroom/image/"&f6&".gif>计："& jiutian_duidut &"点! ["&from1&"]开：<img src=../chatroom/image/"&f1&".gif><img src=../chatroom/image/"&f2&".gif><img src=../chatroom/image/"&f3&".gif>计："& jiutian_duiduf &"点!["&nickname&"]从["&from1&"]赢了"&xiazhu&"两银子！胜者扣10%官税！"            
                        conn.execute "update 用户 set 银两=银两-'"&xiazhu&"' where 姓名='"& from1&"'"
		                conn.execute "update 用户 set 银两=银两+'"& jiayin &"' where 姓名='"& nickname&"'"
		                msg="你同意"&from1&"的对赌！"
                        conn.close
                        set conn=nothing
                        set rs=nothing
                        else
                         msg="你同意"&from1&"的对赌！"
                        abc=1
                        duidu="["&nickname&"]开：<img src=../chatroom/image/"&f4&".gif><img src=../chatroom/image/"&f5&".gif><img src=../chatroom/image/"&f6&".gif>计："& jiutian_duidut &"点! ["&from1&"]开：<img src=../chatroom/image/"&f1&".gif><img src=../chatroom/image/"&f2&".gif><img src=../chatroom/image/"&f3&".gif>计："& jiutian_duiduf &"点!["&from1&"]从["&nickname&"]赢了"&xiazhu&"两银子！胜者扣10%官税！"            
                        conn.execute "update 用户 set 银两=银两+'"&jiayin&"' where 姓名='"& from1 &"'"
		                conn.execute "update 用户 set 银两=银两-'"& xiazhu &"' where 姓名='"& nickname &"'"
		                msg="你同意"&from1&"的对赌！"
                        conn.close
                        set conn=nothing
                        set rs=nothing
                        end if
                      end if
                      end if
                 end if
   end if
   end if
   end if
end if
end if
end if
%>
<html>
<head>
<style>
a:link {text-decoration:none; font-size:12px;color:#000000}
a:hover {text-decoration:underline;font-size:12px; color:#000000; background:#ffffff}
a:visited {text-decoration:none;font-size:12px; color:#000000}
td {font-size:12px; color:#ff0000; line-height:16pt}
input{font-size:12px; color:#ff0000;}
</style>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>对赌银两</title>
<script>
setTimeout('self.close()',3000)
</script>
</head>
<body bgcolor=#FFFFFF  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<table>
<tr>
  <td> <font color="red"> </font>
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#FFFFFF">
      <tr> 
        <td height="115"> 
          <table border="0" bgcolor="#009900" cellspacing="0" cellpadding="2" width="361">
            <tr> 
              <td width="324" height="9"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">对赌银两提示</font><font color="red" size=2> 
                </font></td>
              <td width="29" height="9"> 
                <table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
                  <tr> 
                    <td width="16"><b><a href="javascript:window.close();" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="返回"><font color="000000">×</font></a></b></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table border="0" width="359" cellpadding="4">
            <tr> 
              <td width="59" align="center" valign="top" height="29"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
              <td width="278" height="29">
                <%=msg%>
</td>
            </tr>
            <tr> 
              <td colspan="2" align="center" valign="top" height="58"> 
                <input type=button value='关闭' onClick="javascript:window.close()" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button223">
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>

<%
if abc=1 then
    talkarr=Application("yx8_mhjh_talkarr")
	talkpoint=int(Application("yx8_mhjh_talkpoint"))
	Application("yx8_mhjh_talkpoint")=talkpoint+1
	Dim newtalkarr(600)
	j=1
	for i=11 to 600 step 10
	 newtalkarr(j)=talkarr(i)
	 newtalkarr(j+1)=talkarr(i+1)
	 newtalkarr(j+2)=talkarr(i+2)
	 newtalkarr(j+3)=talkarr(i+3)
	 newtalkarr(j+4)=talkarr(i+4)
	 newtalkarr(j+5)=talkarr(i+5)
	 newtalkarr(j+6)=talkarr(i+6)
	 newtalkarr(j+7)=talkarr(i+7)
	 newtalkarr(j+8)=talkarr(i+8)
     newtalkarr(j+8)=talkarr(i+9)
	 j=j+10
	next
	newtalkarr(591)=talkpoint+1
	newtalkarr(592)=1
	newtalkarr(593)=0
	newtalkarr(594)=nickname
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
	newtalkarr(599)="<b><font color=FF0000>【对赌银两】</font></b>"&duidu&"<font class=timsty>("&time()&")</font>"
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
end if
%>


















