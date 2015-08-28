<%if session("yx8_mhjh_username")=""  then %>
<script language="vbscript">
  alert("您已经超时与服务器断开联接，请重新登陆！")
window.close()
</script>
<%
response.end
end if
my=session("yx8_mhjh_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr= Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language="vbscript">
  alert("您不是剑侠中人，我们不欢迎你！")
</script>
<%
set rs=nothing	
set conn=nothing
response.end
end if
set rs=nothing	
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr= Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 物品 where 名称='矿石' and 所有者='" & my & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
%>
<script language="vbscript">
  alert("您的现在还没有矿石，所以你不可以进行练铁操作！")
  window.close()
</script>
<%
response.end
else
if rs("数量")<2 then
%>
<script language="vbscript">
  alert("您现在的矿石小于2块，所以你不可以进行练铁操作！")
  window.close()
</script>
<%
response.end
end if
Session("tiesl")=rs("数量")
end if
set rs=nothing	
set conn=nothing
%>
<HTML>
<HEAD>
<TITLE>高炉练铁</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<SCRIPT>
function go(n){
        top.fs.waternum.innerHTML=n+"桶";
}
function go1(m){
        top.fs.iceform.itneed.value=m;
}
</SCRIPT>

<META content="MSHTML 5.00.2920.0" name=GENERATOR>
</HEAD>
<%
Session("tiets")=0
Session("tiesj")=true
%>
<FRAMESET border=0 cols=* rows=*,160 frameBorder=no frameSpacing=0>
<FRAME name=fs scrolling=no src="tie.asp" frameSpacing=0 scrolling=no frameBorder=no>
<FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=50,50,0>
<FRAME border=0 frameBorder=no frameSpacing=0 name=ig scrolling=no src="tiejl.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="tiecz.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz1 scrolling=no src="about:blank">
</FRAMESET>
</FRAMESET>
</HTML>
