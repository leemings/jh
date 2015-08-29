<%money = clng(request.form("money"))
my=session("myname")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=120"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><script language=vbscript>
            MsgBox "你不是江湖中人!"
          location.href = "../welcome.asp"
       </script><%
	conn.close
	response.end
else%>
<!--#include file="data1.asp"--><%
randomize()
         waittime=(rnd()*10000000)
         for i=0 to waittime
         next
sql="select * from 房屋 where 户主='" & my & "' or  伴侣='" & my & "'"
set rs=conn.execute(sql)
yin=rs("伴侣银两")
if money>yin then%>
<script language=vbscript>
                         MsgBox "错误！您的金库里没有那么多银两。"
                         location.href = "xiaowu5.asp"
                    </script>
	<%elseif money>50000 or money<1 then%>
       <script language=vbscript>
                         MsgBox "错误！您填写的数量不能小于1。"
                         location.href = "xiaowu5.asp"
                    </script><%
		else
       sql="update 房屋 set 伴侣银两=伴侣银两-'" & money & "' where 户主='" & my & "' or  伴侣='" & my & "'"
       rs=conn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
       sql="update 用户 set 银两=银两+'" & money & "' where 姓名='" & my & "'"
       rs=conn.execute(sql)
	conn.close
       if yin<0 then
       sql="update 用户 set 银两=银两+'" & yin & "' where 姓名='" & my & "'"
       rs=conn.execute(sql)
	conn.close
       end if
	Response.Redirect "xiaowu5.asp"
			conn.close
			response.end
		end if
   rs.close
   set rs=nothing
end if
%>

<html><script language="JavaScript">                                                                  </script></html>