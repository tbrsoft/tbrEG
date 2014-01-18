private path

public sub execsql(sql)
	Set cn = CreateObject("adodb.connection")

'    	cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Martin\TBR\Cemi\BD\cemi.mdb;Persist Security Info=False"
    	cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Persist Security Info=False"
	cn.connectionstring = cs
    	cn.open
	set rs=cn.execute(sql)
	dim f
	for each f in rs.fields
		txt.text=txt.text+cstr(f.name)+" - "
	next
	txt.text=txt.text+vbcrlf+"------------------------------------------------------------------------"+vbcrlf

	while not rs.eof
		for each f in rs.fields
			txt.text=txt.text+cstr(f.value)+" - "
		next
		txt.text=txt.text+vbcrlf
		rs.movenext
	wend

end sub


public sub AbrirBaseDatos
	cd.dialogtitle="Seleccione la base de datos..."
	cd.showopen
	path = cd.filename	
end sub

public sub VerLogErrores()
	 Dim fso 'As FileSystemObject
    Dim f
        
    Dim s 'As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set f = fso.GetFile("d:\log.txt")
    
    Set ts = f.OpenAsTextStream(1)
    
    s = ts.ReadAll
    
    ts.Close
    txt=txt+ s
end sub

Public sub Print(cadena)
	txt.text=txt.text+cadena+vbcrlf+ ">" 
End sub

public sub SetearPathEnConfig()
	vbInt.SaveSettingA "TbrEmergencyGroup", "DBLayer", "PathDB", path	
end sub

public sub Conectar(ip)
	wsock.RemoteHost =ip
	wsock.connect
end sub