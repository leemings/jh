<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Buffer =true
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
roomsn=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
chatroomname=Application("sjjh_chatroomname")%>
<head>
<META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<meta http-equiv=refresh content='300;url=selectchatroom.asp?roomsn=<%=roomsn%>'>
<title>选择区♀wWw.happyjh.com♀</title>
<style type='text/css'>
body{
CURSOR: url('40.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
font-size:9pt;}
input{font-size:9pt;background-color:FF9900;color:FFFFFF;border: 1 double}
a{font-size:9pt;color:FFFF00;text-decoration:none;}
a:hover{color:FFFF00;
text-decoration:underline;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#006699" bgproperties="fixed" leftMargin=10 marginwidth=0 marginheight=0 topMargin=10>
<div align="center"><font color="#ffffff" size="2"><strong>房 
        间 列 表</strong></font></div>
	
<hr width="134" size="1" noshade color=FFFF00>
<div align="center"><a href=javascript:history.go(0)><font class=p9>刷新</font></a><br>
  <br>
  <%for i=0 to chatroomnum
	online=split(trim(Application("sjjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	sj_chat_info=split(sjjh_roominfo(i),"|")%>
  <img src='img/room.gif'> 
  <select name=chatroomselect style='BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px'>
    <option value='<%=i%>/<%=sj_chat_info(0)%>' selected><%=sj_chat_info(0)%></option>
    <option>聊天:<%=onlinenum%>人</option>
    <option>上限:<%=sj_chat_info(1)%>人</option>
    <option>限制:
    <%if cstr(sj_chat_info(2))=0 then%>
    没有
    <%else%>
    存在
    <%end if%>
    </option>
    <option>动武:
    <%if cstr(sj_chat_info(5))=0 then%>
    允许
    <%else%>
    禁止
    <%end if%>
    </option>
    <option>事件:
    <%if cstr(sj_chat_info(6))=0 then%>
    允许
    <%else%>
    禁止
    <%end if%>
    </option>
    <option>保护:
    <%if cstr(sj_chat_info(7))=0 then%>
    允许
    <%else%>
    禁止
    <%end if%>
    </option>
    <option>卡片:
    <%if cstr(sj_chat_info(8))=0 then%>
    允许
    <%else%>
    禁止
    <%end if%>
    </option>
    <option>赌博:
    <%if cstr(sj_chat_info(9))=0 then%>
    允许
    <%else%>
    禁止
    <%end if%>
    </option>
    <%if onlinenum>0 then%>
    <option>----------</option>
    <%online=split(trim(Application("sjjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	useronlinename=replace(trim(Application("sjjh_useronlinename"&i)),"  ","</option><option>")%>
    <option><%=useronlinename%></option>")
    <%end if%>
  </select>
  <input  alt="<%=sj_chat_info(3)%>" type=image src="img/gn/queren.gif" width="30" height="19" border="0" align="absmiddle"  onclick=javascript:parent.d.location.replace("changeroom.asp?roomsn=<%=i%>&chatroomname=<%=sj_chat_info(0)%>");>
  <br>
  <%next%>
</div>
<table width=90% align="center" style=font-size:9pt>
  <tr>
    <td><font color=white>请选择您所要进入的房间.鼠标停在<font color="#00FF00">[确认]</font>上会显示进入的条件.</font></table>
</div>
</body>
