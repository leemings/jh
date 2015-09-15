<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Buffer =true
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
roomsn=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
chatroomname=Application("aqjh_chatroomname")%>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<meta http-equiv=refresh content='300;url=selectchatroom.asp?roomsn=<%=roomsn%>'>
<title>选择区</title>
<style type='text/css'>
body{font-size:9pt;}
input{font-size:9pt;background-color:FF9900;color:FFFFFF;border: 1 double}
a{font-size:9pt;color:FFFF00;text-decoration:none;}
a:hover{color:FFFF00;
text-decoration:underline;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=Session("afa_chatbgcolor")%>" background="<%=Application("aqjh_chatimage")%>" bgproperties="fixed" leftMargin=1 marginwidth=0 marginheight=0 topMargin=1>
<div align=center><br><br>
<font color="#FFFFFF">【房间列表】</font>
<hr size="1" noshade color=FFFF00>
  <a href=javascript:history.go(0)><font class=p9>刷新</font></a><br>
  <br>
  <%for i=0 to chatroomnum
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	sj_chat_info=split(aqjh_roominfo(i),"|")%>
  <img src='img/room.gif'>
  <select name=chatroomselect>
    <option value='<%=i%>/<%=sj_chat_info(0)%>' selected><%=sj_chat_info(0)%></option>
    <option>聊天:<%=onlinenum%>人</option>
    <option>上限:<%=sj_chat_info(1)%>人</option>
    <option>限制:<%if cstr(sj_chat_info(2))=0 then%>没有<%else%>存在<%end if%></option>
    <option>动武:<%if cstr(sj_chat_info(5))=0 then%>允许<%else%>禁止<%end if%></option>
    <option>事件:<%if cstr(sj_chat_info(6))=0 then%>允许<%else%>禁止<%end if%></option>
    <option>保护:<%if cstr(sj_chat_info(7))=0 then%>允许<%else%>禁止<%end if%></option>
	<option>卡片:<%if cstr(sj_chat_info(8))=0 then%>允许<%else%>禁止<%end if%></option>
	<option>赌博:<%if cstr(sj_chat_info(9))=0 then%>允许<%else%>禁止<%end if%></option>
    <option>时限:前<%=sj_chat_info(10)%>分</option>
    <option>泡点:<%if cstr(sj_chat_info(11))=1 then%>允许<%else%>禁止<%end if%></option>
    <%if onlinenum>0 then%><option>----------</option>
    <%online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	useronlinename=replace(trim(Application("aqjh_useronlinename"&i)),"  ","</option><option>")%>
    <option><%=useronlinename%></option>")
    <%end if%>
  </select><br><br>
  <%next%>
</div>
</body>