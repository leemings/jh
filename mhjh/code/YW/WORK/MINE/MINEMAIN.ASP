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
connstr=Application("yx8_mhjh_connstr")
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
if trim(rs("门派"))="十八地狱" then%>
<script language="vbscript">
  alert("鬼魂不能在这里挖矿，请去凤凰蘖盘进化成人！")
  window.close()
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
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 物品 where 名称='矿石' and 所有者='" & my & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
Session("minesl")=0
else
Session("minesl")=rs("数量")
end if
set rs=nothing	
set conn=nothing
%>
<HTML>
<HEAD>
<TITLE>深山采矿</TITLE>
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
Session("minets")=0
Session("minesj")=true
%>
<FRAMESET border=0 cols=* rows=*,180 frameBorder=no frameSpacing=0>
<FRAME name=fs scrolling=no src="mine.asp" frameSpacing=0 scrolling=no frameBorder=no>
<FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=50,50,0>
<FRAME border=0 frameBorder=no frameSpacing=0 name=ig scrolling=no src="minejl.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="minecz.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz1 scrolling=no src="about:blank">
</FRAMESET>
</FRAMESET>
</HTML>
