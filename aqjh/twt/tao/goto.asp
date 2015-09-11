<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
  s=now()-session("tao_time")
  if int(s*1000000)<30 then
  response.write "请不要刷呀，悠着点，江湖上不少人呢！"
  response.end
  end if
%>
<!--#include file="data.asp"-->
<%
  randomize timer
  t=int(rnd*100)
    if t<30 then
    session("tao_time")=now()
  randomize timer
  jin=int(rnd*100)
  sql="select * from 淘金者 where 拥有者='" & aqjh_name & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into 淘金者(拥有者,金,银,铜) values('" & aqjh_name & "'," & jin & ",0,0)"
       conn.execute(sql)
       else
       sql="update 淘金者 set 金=金+" & jin & " where 拥有者='" & aqjh_name & "'"
       conn.execute(sql)
       end if
       conn.execute("update 矿石 set 总点=总点+1,总流=总流+" & jin& " where 矿种='金'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=aqjh_name%>" & "恭喜您!获得" & <%=jin%> & "点金矿!"
  location.href="index.asp"
</script>
<%
    elseif t<65 and t>=30 then
  session("tao_time")=now()
  randomize timer
  yin=int(rnd*200)
  sql="select * from 淘金者 where 拥有者='" & aqjh_name & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into 淘金者(拥有者,金,银,铜) values('" & aqjh_name & "',0," & yin & ",0)"
       conn.execute(sql)
       else
       sql="update 淘金者 set 银=银+" & yin & " where 拥有者='" & aqjh_name & "'"
       conn.execute(sql)
       end if
       conn.execute("update 矿石 set 总点=总点+1,总流=总流+" & yin& " where 矿种='银'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=aqjh_name%>" & "恭喜您!获得" & <%=yin%> & "点银矿!"
  location.href="index.asp"
</script>
<%
    else
  session("tao_time")=now()
  randomize timer
  tong=int(rnd*300)
  sql="select * from 淘金者 where 拥有者='" & aqjh_name & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into 淘金者(拥有者,金,银,铜) values('" & aqjh_name & "',0,0," & tong & ")"
       conn.execute(sql)
       else
       sql="update 淘金者 set 铜=铜+" & tong & " where 拥有者='" & aqjh_name & "'"
       conn.execute(sql)
       end if
       conn.execute("update 矿石 set 总点=总点+1,总流=总流+" & tong & " where 矿种='铜'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=aqjh_name%>" & "恭喜您!获得" & <%=tong%> & "点铜矿!"
  location.href="index.asp"
</script>
<%
    end if
%>