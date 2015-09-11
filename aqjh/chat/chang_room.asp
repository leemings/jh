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
chatinfo=split(aqjh_roominfo(roomsn),"|")
chatroomnum=ubound(aqjh_roominfo)-1
chatroomname=Application("aqjh_chatroomname")%>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<meta http-equiv=refresh content='300;url=chang_room.asp?roomsn=<%=roomsn%>'>
<title>选择区</title>
<style type='text/css'>
body{font-size:9pt;}
input{font-size:9pt;background-color:FF9900;color:FFFFFF;border: 1 double}
a{font-size:9pt;color:FFFF00;text-decoration:none;}
a:hover{color:FFFF00;
text-decoration:underline;}
</style>
</head>
<body bgcolor="#b7d4f1" bgproperties="fixed" leftMargin=0 marginwidth=0 marginheight=0 topMargin=0 oncontextmenu=window.event.returnValue=false><center>
<form name="form1">
<select name=chatroomselect onChange='javascript:changechatroom();' style='width: 100%; height: 100%'>
  <%for i=0 to chatroomnum
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	sj_chat_info=split(aqjh_roominfo(i),"|")%>
<option value='<%=i%>/<%=sj_chat_info(0)%>' <%if chatinfo(0)=sj_chat_info(0) then%>selected<%end if%> style=BACKGROUND-COLOR:#ffffff;font-size:9pt;color:red><%=sj_chat_info(0)%><%=onlinenum%>人聊天</option>
  <%next%>
</select></form>
</body>
<script language=javascript>
<!--
function changechatroom()
{
var chatroomoption=document.form1.chatroomselect.value;
var chatroomoptmp;
var chatroomsn;
var chatroomname;
chatroomtmp=chatroomoption.indexOf('/');
chatroomsn=chatroomoption.substring(0,chatroomtmp);
chatroomname=chatroomoption.substring(chatroomtmp+1);
document.form1.chatroomselect.disabled=false;
top.d.location.href="changeroom.asp?roomsn="+chatroomsn+"&chatroomname="+chatroomname;
}
-->
</script>