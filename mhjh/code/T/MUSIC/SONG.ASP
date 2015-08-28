<!--#include file="data.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
	id=request("id")
	sql="select * from 歌曲 where id="&id
	Set Rs=conn.Execute(sql)   
	if rs.eof then
		errmsg="<li>对不起！该歌曲不存在，可能已经被管理员删除。</li>"
		call error()
		Response.End 
	else
	    sql="update 歌曲 set 点击=点击+1 where id="&id
	    Set Rs=conn.Execute(sql) 
	    sql="select * from 歌曲 where id="&id
	    Set Rs=conn.Execute(sql)  
		source=rs("地址")
		MusicName=rs("歌名")
		geshou=rs("歌手")
	end if
	rs.Close
	response.write "<html><head><title>歌曲播放："&MusicName&"--"&geshou&"</title></head>"
set rs=nothing
conn.close
set conn=nothing
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")                         
conn.Open Application("yx8_mhjh_connstr")                         
set rs=server.CreateObject("adodb.recordset")                         
rs.Open "select 银两,体力 from 用户 where 姓名='"&username&"'",conn  
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能听歌"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
tl=rs("银两")
nl=rs("体力")
if tl<100 or nl<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "错误！你目前的银两或内力值不够，回去多多拿点钱然后练练马步蹲裆再来吧！"
location.href = "javascript:close()"
</script>
<%
else
conn.execute "update 用户 set 银两=银两-20,体力=体力-10 where 姓名='"&username&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<body topmargin=0 leftmargin=0 bgcolor=#006AAA oncontextmenu=self.event.returnValue=false>
<table border="0" width="275" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%"><img border="0" src="player/top_left.gif" width="6" height="18"><img border="0" src="player/top.gif" width="248" height="18"><a href='javascript:window.close();'><img border="0" src="player/close.gif" width="21" height="18"></a></td>
  </tr>
  <tr>
    <td width="100%">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td width="6"><img border="0" src="player/left.gif" width="6" height="66"></td>
          <td width="248">
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50%">
                  <table border="0" width="100%" cellspacing="0" cellpadding="0">
                    <tr>
                      <td colspan="2"></td>
                    </tr>
                    <tr>
                      <td width="100%">
<object id=video1 classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" height=66 width=248>
              <param name="controls" value="StatusBar,ControlPanel">
              <param name="console" value="Clip1">
              <param name="autostart" value="true">
              <param name="src" value="<%=source%>">
              <embed src="<%=source%>" type="audio/x-pn-realaudio-plugin" console="Clip1" controls="ControlPanel,StatusBar" height=66 width=248 autostart=true>
              </embed></object></td></tr>
</table></td></tr></table></td><td width="21"><img border="0" src="player/right.gif" width="21" height="66"></td>
</tr></table></td></tr><tr><td width="100%"><table border="0" width="100%" cellspacing="0" cellpadding="0"><tr>
<td width="33%"><img border="0" src="player/below_left.gif" width="6" height="32"></td>
<td width="33%"><img border="0" src="player/below.gif" width="248" height="32"></td>
<td width="34%"><img border="0" src="player/below_right.gif" width="21" height="32"></td></tr></table></td></tr></table>
</body></html>
<%end if%>




