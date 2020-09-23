<div align="center">

## Simple Search


</div>

### Description

Search in the site for any word
 
### More Info
 
The pages that containt the especific word


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jorge Cisneros](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jorge-cisneros.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jorge-cisneros-simple-search__4-6143/archive/master.zip)





### Source Code

```
<HTML>
<HEAD>
<% if Request("SearchText") <> "" then
		response.write "<TITLE>Buscando Resultados para"& Request("SearchText")& "</TITLE>"
	else
		response.write "<TITLE>Página de busqueda </TITLE>"
end if %>
</HEAD>
<BODY>
<form method="post" action="search.asp">
 <div align="center"></div>
 <div align="center">
  <table bgcolor="#0033CC" border="0" bordercolorlight="#00FFFF" bordercolordark="#000000">
   <tr>
    <td> <font color="#CCCCCC"> <font color="#FFFFFF">Buscar</font>
     <input type="text" name="SearchText" size="40">
     </font> </td>
    <td> <font color="#CCCCCC">
     <input type="submit" name="Submit2" value="Buscar">
     </font></td>
   </tr>
   <tr>
    <td height="32"> <font color="#CCCCCC"> <font color="#FFFFFF">Mostrar
     resultados</font>
     <select name="rLength" >
      <option value="200" SELECTED>Descripción Larga
      <option value="100">Descripción Corta
      <option value="0">Solo Url
     </select>
     <select name="rResults">
      <option value="10" SELECTED>10
      <option value="25">25
      <option value="50">50
     </select>
     </font> </td>
    <td height="32"> <font color="#CCCCCC">
     <input type="reset" name="Reset" value="Borrar">
     </font></td>
   </tr>
  </table>
 </div>
</form>
<p><% if Request("SearchText") <> "" then %> </p>
<p><B>Buscando Resultados para '<%=Request("SearchText")%>'</B><BR>
 <%
'
' Buscador Simple. Autor Jorge Cisneros jorgeci@hotmail.com
'
Const fsoForReading = 1
Dim objFile, objFolder, objSubFolder, objTextStream
Dim bolCase, bolFileFound
dim strDeTag, Ext, strFile, strContent, strRoot, strTag, strText, strTitle, strTitleL
Dim reqLength, reqNumber, count
strFile = ".asp .htm .html .js .txt .css"
strRoot = "/"
strText = Request("SearchText")
If Request("Case") = "on" Then bolCase = 0 Else bolCase = 1
If Request("rResults") = "10" Then reqNumber = 10
If Request("rResults") = "25" Then reqNumber = 25
If Request("rResults") = "50" Then reqNumber = 50
reqLength = Request("rLength")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(Server.MapPath(strRoot))
CurURL= "http://" & Request.serverVariables("SERVER_NAME")
CurPath = objFolder
schSubFol(objFolder)
For Each objSubFolder in objFolder.SubFolders
	schSubFol(objSubFolder)
Next
If Not bolFileFound then Response.Write "La busqueda no encontro nada.."
If bolFileFound then Response.Write "<B>Fin de la busqueda</B>"
Set objTextStream = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
Sub schSubFol(objFolder)
	For Each objFile in objFolder.Files
		If Count + 1 > reqNumber or strText = "" Then Exit Sub
		If Response.IsClientConnected Then
			'Abrir solo si es archivo conocido
			strext = right(objFile.Name,3)
			if instr(1,strFile,strext) > 0 then
 				Set objTextStream = objFSO.OpenTextFile(objFile.Path,fsoForReading)
				strContent = objTextStream.ReadAll
				If InStr(1,strContent,strtext) > 0 Then
				postitle = InStr(1, strContent, "<TITLE>",1)
				If postitle > 0 Then
					strTitle = Mid(strContent, postitle + 7, InStr(1, strContent, "</TITLE>", 1) - (postitle + 7))
				Else
					strTitle = "Sin Titulo"
				end if
				Count = Count + 1
				Response.Write "<DL><DT><B><I>"& Count &"</I></B> - <A HREF="&Obt_Url(objFile.path) & ">" & strTitle & "</A></DT><BR><DD>"
				strTitleL = InStr(1, strContent, "</TITLE>", 1) - InStr(1, strContent, "<TITLE>", 1) + 7
				strDeTag = ""
				bolTagFound = False
				Do While InStr(strContent, "<")
					bolTagFound = True
					strDeTag = strDeTag & " " & Left(strContent, InStr(strContent, "<") - 1)
					strContent = MID(strContent, InStr(strContent, ">") + 1)
				Loop
				strDeTag = strDeTag & strContent
				If Not bolTagFound Then strDeTag = strContent
					If reqLength = "0" Then
						Response.Write obt_url(objFile.Path) & "</DD></DL>"
					Else
						Response.Write Mid(strDeTag, strTitleL, reqLength) & "...<BR><I><FONT SIZE='2'>URL: " & obt_url(objFile.Path) & " - Ultima modificación: " & objFile.DateLastModified & " - " & FormatNumber(objFile.Size / 1024) & "Kbytes</FONT></I></DD></DL>"
					end if
					bolFileFound = True
				End If
				objTextStream.Close
			End If
		End If
	Next
End Sub
Function Obt_Url (nompath)
	obt_url = CurUrl +"/"+ right(nompath,len(nompath) - len(curpath)-1)
end function
%>
<% end if %>
</p>
</BODY></HTML>
```

