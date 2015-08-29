<%
	dim loopok,LinkUrl,midi,PlayerStats,BoardID
	MIDI       = Request("midi")			'Midi文件名
	BoardID	   = Request("BoardID")			'版面ID
	LoopOK	   = -1							'循环次数    -1 表示不断循环播放
	
	if not isnumeric(BoardID) or isnull(BoardID) then BoardID=0 
	PlayerStats= Request.Cookies("MidiPlayer")("Stats_"&BoardID)       '播放器状态
	LinkUrl    = Request.Cookies("MidiPlayer")("LinkUrl_"&BoardID)     'MIDI路径 
	
	
	
	if MIDI="-1" then   	'取消背景音乐  
		PlayerStats="stop"
		Response.Cookies("MidiPlayer")("Stats_"&BoardID)="stop"
		Response.Cookies("MidiPlayer").Expires=Date+30		'保留30天
	elseif MIDI<>"" and (not isnull(MIDI))  then    '点击某个MIDI时播放 
		LinkUrl="midi/"					'Midi文件存放路径目录，可以修改
		LinkUrl= LinkUrl & midi & ".mid"
		PlayerStats="play" 
		Response.Cookies("MidiPlayer")("Stats_"&BoardID)="play"
		Response.Cookies("MidiPlayer")("LinkUrl_"&BoardID)=LinkUrl
		Response.Cookies("MidiPlayer").Expires=Date+30
	end if
		
	if PlayerStats="play" then
		Response.write "<html><head><title>MIDI点歌系统</title></head><body>" 
		Response.Write "<bgsound src="""& LinkUrl & """ loop=""" & LoopOK & """>"
		Response.write "</body></html>"
	end if
%> 