<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<script>
if(window.name!="aqjhhorse"){ var i=1;  while (i<=50)  {    window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ȥ�ɣ�����������50�Σ���");    i=i+1;  }top.location.href="../../exit.asp"} </script>
<head><title><%=Application("aqjh_chatroomname")%>��������ɻ��׵�����ɣ�</title><link rel=stylesheet href='css.css'></head>
<frameset rows="*,50">
<frame name=compfrm src="compete.asp" noresize  scrolling=no >
<frame name=betfrm src="chipin.asp"  scrolling=no marginheight=0 framespacing=0 marginwidth=0>
</frameset>