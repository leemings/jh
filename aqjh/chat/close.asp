<%@ LANGUAGE=VBScript codepage ="936" %><html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>关闭窗口</title>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=FFFFFF>
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>
</body>
</html>