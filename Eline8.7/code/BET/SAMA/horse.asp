<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<script>
if(window.name!="sjjhhorse"){ var i=1;  while (i<=50)  {    window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ȥ�ɣ�����������50�Σ���");    i=i+1;  }top.location.href="../../exit.asp"} </script>
<head><title><%=Application("sjjh_chatroomname")%>�������</title><link rel=stylesheet href='css.css'></head>
<frameset rows="*,50">
<frame name=compfrm src="compete.asp" noresize  scrolling=no >
<frame name=betfrm src="chipin.asp"  scrolling=no marginheight=0 framespacing=0 marginwidth=0>
</frameset>