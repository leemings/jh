<%
	dim loopok,LinkUrl,midi,PlayerStats,BoardID
	MIDI       = Request("midi")			'Midi�ļ���
	BoardID	   = Request("BoardID")			'����ID
	LoopOK	   = -1							'ѭ������    -1 ��ʾ����ѭ������
	
	if not isnumeric(BoardID) or isnull(BoardID) then BoardID=0 
	PlayerStats= Request.Cookies("MidiPlayer")("Stats_"&BoardID)       '������״̬
	LinkUrl    = Request.Cookies("MidiPlayer")("LinkUrl_"&BoardID)     'MIDI·�� 
	
	
	
	if MIDI="-1" then   	'ȡ����������  
		PlayerStats="stop"
		Response.Cookies("MidiPlayer")("Stats_"&BoardID)="stop"
		Response.Cookies("MidiPlayer").Expires=Date+30		'����30��
	elseif MIDI<>"" and (not isnull(MIDI))  then    '���ĳ��MIDIʱ���� 
		LinkUrl="midi/"					'Midi�ļ����·��Ŀ¼�������޸�
		LinkUrl= LinkUrl & midi & ".mid"
		PlayerStats="play" 
		Response.Cookies("MidiPlayer")("Stats_"&BoardID)="play"
		Response.Cookies("MidiPlayer")("LinkUrl_"&BoardID)=LinkUrl
		Response.Cookies("MidiPlayer").Expires=Date+30
	end if
		
	if PlayerStats="play" then
		Response.write "<html><head><title>MIDI���ϵͳ</title></head><body>" 
		Response.Write "<bgsound src="""& LinkUrl & """ loop=""" & LoopOK & """>"
		Response.write "</body></html>"
	end if
%> 