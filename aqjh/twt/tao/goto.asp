<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
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
  response.write "�벻Ҫˢѽ�����ŵ㣬�����ϲ������أ�"
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
  sql="select * from �Խ��� where ӵ����='" & aqjh_name & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into �Խ���(ӵ����,��,��,ͭ) values('" & aqjh_name & "'," & jin & ",0,0)"
       conn.execute(sql)
       else
       sql="update �Խ��� set ��=��+" & jin & " where ӵ����='" & aqjh_name & "'"
       conn.execute(sql)
       end if
       conn.execute("update ��ʯ set �ܵ�=�ܵ�+1,����=����+" & jin& " where ����='��'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=aqjh_name%>" & "��ϲ��!���" & <%=jin%> & "����!"
  location.href="index.asp"
</script>
<%
    elseif t<65 and t>=30 then
  session("tao_time")=now()
  randomize timer
  yin=int(rnd*200)
  sql="select * from �Խ��� where ӵ����='" & aqjh_name & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into �Խ���(ӵ����,��,��,ͭ) values('" & aqjh_name & "',0," & yin & ",0)"
       conn.execute(sql)
       else
       sql="update �Խ��� set ��=��+" & yin & " where ӵ����='" & aqjh_name & "'"
       conn.execute(sql)
       end if
       conn.execute("update ��ʯ set �ܵ�=�ܵ�+1,����=����+" & yin& " where ����='��'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=aqjh_name%>" & "��ϲ��!���" & <%=yin%> & "������!"
  location.href="index.asp"
</script>
<%
    else
  session("tao_time")=now()
  randomize timer
  tong=int(rnd*300)
  sql="select * from �Խ��� where ӵ����='" & aqjh_name & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into �Խ���(ӵ����,��,��,ͭ) values('" & aqjh_name & "',0,0," & tong & ")"
       conn.execute(sql)
       else
       sql="update �Խ��� set ͭ=ͭ+" & tong & " where ӵ����='" & aqjh_name & "'"
       conn.execute(sql)
       end if
       conn.execute("update ��ʯ set �ܵ�=�ܵ�+1,����=����+" & tong & " where ����='ͭ'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=aqjh_name%>" & "��ϲ��!���" & <%=tong%> & "��ͭ��!"
  location.href="index.asp"
</script>
<%
    end if
%>