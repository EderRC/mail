Static Function PCorreo(cEmail, cAssunto, cCorpo, cAnexo, cAnexo2)
    Local cExecute := "/c ipm.note /m "
    Default cEmail := ""
    Default cAssunto := ""
    Default cCorpo := ""
    Default cAnexo := ""
	Default cAnexo2 := ""
     
    //Se tiver email, abre o outlook
    If !Empty(Alltrim(cEmail)) .And. ! Empty(cAssunto) .And. ! Empty(cCorpo)
        //Muda o -enter- e tira aspas
        cCorpo := StrTran(cCorpo, CRLF, " %0D%0A ")
        cCorpo := StrTran(cCorpo, '"', '')
 
        //Monta o comando
        cExecute += '"'
        cExecute += cEmail
        cExecute += '?subject=' + cAssunto
        cExecute += '&body=' + cCorpo
        cExecute += '"'
 
        //Se tiver anexo
        If ! Empty(cAnexo) .And. File(cAnexo)


	aAnexos:={cAnexo,cAnexo2}
		
		For nAtual := 1 To Len(aAnexos)
		cExecute += ' /a "' + (aAnexos[nAtual]) + '"'
		// cExecute+='/a"'+ (aAnexos[nAtual])'"'
		Next
	// cExecute += ' /a "' + aAnexos[1] + '"'
		// cExecute += '/a "C:\PR\nf_e_166321.pdf", /a "C:\PR\nf_e_166321.xml"' //agrega el primero omite el segundo
		// cExecute += '/a "C:\PR\nf_e_166321.pdf C:\PR\nf_e_166321.xml"'
		// cExecute += '/a "C:\PR\nf_e_166321.pdf" /a "C:\PR\nf_e_166321.xml"' //agrega el primero omite el segundo
		// cExecute += '/a "C:\PR\nf_e_166321.pdf","C:\PR\nf_e_166321.xml"'
		// cExecute += '/a "C:\PR\nf_e_166321.pdf,C:\PR\nf_e_166321.xml"'
		// cExecute += ' /a "' + cAnexo2 + '"'
            // cExecute += ' /a "' + cAnexo + '"'+'",'+cAnexo2+'"'
			// cExecute += ' /a "' + cAnexo + ',' + cAnexo2 + '"'
        EndIf

		// If ! Empty(cAnexo2) .And. File(cAnexo2)
        //     cExecute += ' /a "' + cAnexo2 + '"'
        // EndIf

	     //Abre a tela do outlook
        ShellExecute("OPEN", "outlook.exe", cExecute, "", 1)
    EndIf
Return
