<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<HTML><HEAD><TITLE>E线江湖-单人台球</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT src="gamexver.js"></SCRIPT>

<SCRIPT language=JavaScript src="files_dir.js"></SCRIPT>

<SCRIPT language=javascript event=OnCmd(type,cmd) for=GX>
switch(type){
case "appinit":
	GX.Cmd("<sys I0d<KKBOBKGLHQIDPHNNHHOJNMMIOKLRJPDLMECDCNHNODCFDLJINNJOGIGSNJAMCLPHGEBNLFNHIQBPDEEOHEIMDDONPGGOEKIDJFPQIQLKCDEDCIMIFHJLMOLJPNKOEQDKEPOJMFOJMMPGNRCNMEDFMIOIBEBGDNGNBHDIDILFGIFQMKLDBLKHJSINOHEH>");
	break;
}
</SCRIPT>

<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#c0c0c0 leftMargin=0 topMargin=0 scroll=no>
<OBJECT id=GX 
codeBase=gamex.cab#Version=1,1,30910,860 
classid=CLSID:2D4851FD-0BFE-11D4-9260-9AF666D52059 width="100%" height="100%" 
VIEWASTEXT></OBJECT></BODY></HTML>
