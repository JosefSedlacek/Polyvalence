Option Explicit

Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub Aktualizace_ze_SAP()

Dim response As VbMsgBoxResult
response = MsgBox("Jste přihlášení do SAP a máte otevřené SAP hlavní menu?", vbYesNo + vbQuestion, "Otevřený SAP na hlavním menu")

    If response = vbNo Then
        MsgBox "Nejprve se musíte přihlásit do SAP a být na hlavní obrazovce SAP (SAP menu)"
    Else
        Call Aktualizovat_ZPP_ZPHL
    End If

End Sub

Sub Aktualizovat_ZPP_ZPHL()

Worksheets("AKTUALIZACE").Unprotect Password:="123456"

Worksheets("AKTUALIZACE").Range("C12") = ""
Worksheets("AKTUALIZACE").Range("C14") = ""
Worksheets("AKTUALIZACE").Range("C16") = ""

'###########################
'Stav aktualizace - popisek:
Dim StavAktualizace As Range
Set StavAktualizace = Worksheets("AKTUALIZACE").Range("J8")

StavAktualizace = "Probíhá stahování ..."

'###########################################
'Napojení na scriptovací nástroj SAP GUI API
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

'##########################################
'Definovat rozmezí období pro SAP transakci
Dim Zacatek As Date
    Zacatek = Worksheets("AKTUALIZACE").Range("F9").Value
Dim Konec As Date
    Konec = Worksheets("AKTUALIZACE").Range("F10").Value

Dim StazenyZacatek As Range
Set StazenyZacatek = Worksheets("AKTUALIZACE").Range("A12")
StazenyZacatek = Zacatek

Dim StazenyKonec As Range
Set StazenyKonec = Worksheets("AKTUALIZACE").Range("A14")
StazenyKonec = Konec

'##############################################
'Scripting SAP - proběhne transakce ZPP_ZPHL
'Nastavení varianty SEDLAJOS a layoutu /POLYVAL

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "zpp_zphl"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "sedlajos"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press

'xxxxxxxxxxxxxxxxxxxxxxxxxx  Nastavení začátku a konce - datum:
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = Zacatek
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = Konec
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").selectContextMenuItem "&LOAD"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 11, "TEXT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "11"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "polyvalence.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


StavAktualizace = "Stažení proběhlo úspěšně." & vbCrLf _
& "NÁSLEDUJE  STAŽENÍ  DOCHÁZKY"
Worksheets("AKTUALIZACE").Range("C12") = "OK"
Worksheets("AKTUALIZACE").Protect Password:="123456"
End Sub

Sub Stazeni_dochazky()
  Worksheets("AKTUALIZACE").Unprotect Password:="123456"
  Dim OutlookApp As Object
  Dim OutlookMail As Object
      
  ' Vytvořte novou instanci aplikace Outlook
  Set OutlookApp = CreateObject("Outlook.Application")
  Set OutlookMail = OutlookApp.CreateItem(0)
      
      With OutlookMail
          .To = "josef.sedlacek@gumokov.cz"
          .Subject = "DOCHÁZKA POLYVALENCE"
          .Body = "Potřeba aktualizovat docházku pro všechny lisovny. Návod je " & vbCrLf _
          & "v aktualizačním excelu, který najdu na disku P"
          .Importance = 2
          .Send
      End With
      
      ' Uvolnění objektů
      Set OutlookMail = Nothing
      Set OutlookApp = Nothing
      MsgBox "Byl odeslán mail s upozorněním. Docházku stáhne: josef.sedlacek@gumokov.cz"
  
  Dim StavAktualizace As Range
  Set StavAktualizace = Worksheets("AKTUALIZACE").Range("J8")
  StavAktualizace = "Docházka se připravuje" & vbCrLf _
  & "NÁSLEDUJE  PROPSÁNÍ SEZNAMŮ"
  
  Worksheets("AKTUALIZACE").Range("C14") = "OK"
  Worksheets("AKTUALIZACE").Protect Password:="123456"
End Sub

Sub Propsat_seznamy()

  Worksheets("AKTUALIZACE").Unprotect Password:="123456"
  
  Dim StavAktualizace As Range
  Set StavAktualizace = Worksheets("AKTUALIZACE").Range("J8")
  
  StavAktualizace = "Probíhá propisování ..."
  Dim den As String
  Dim datum As Date
  den = Worksheets("AKTUALIZACE").Range("A5").Value
  datum = Worksheets("AKTUALIZACE").Range("A8").Value

  Application.ScreenUpdating = False
  
  '###################################################
  '#####################   Propsání seznamu pracovníků
  
  Dim ws As Worksheet
  Dim novySeznam As Workbook
  Dim cilovaCesta As String
  
  ' Nastaví cestu a název souboru
  cilovaCesta = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS\Pracovnici.xlsx"
  
  ' Nastaví list "Seznam pracovníků"
  Set ws = ThisWorkbook.Sheets("Seznam pracovníků")
  
  ' Vytvoří nový sešit
  Set novySeznam = Workbooks.Add
  
  ' Zkopíruje list do nového sešitu
  ws.Copy Before:=novySeznam.Sheets(1)
  
  ' Uloží nový sešit
  Application.DisplayAlerts = False ' Potlačí upozornění na přepsání souboru
  novySeznam.SaveAs Filename:=cilovaCesta, FileFormat:=xlOpenXMLWorkbook
  Application.DisplayAlerts = True ' Opět povolí upozornění
  
  ' Zavře nový sešit
  novySeznam.Close SaveChanges:=False
  
  
  '####################################################
  '#########################  Propsání seznamu odchylek
  
  Dim novySeznam2 As Workbook
  Dim cilovaCesta2 As String
  
  cilovaCesta2 = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS\Odchylky.xlsx"
  Set ws = ThisWorkbook.Sheets("Odchylky")
  Set novySeznam2 = Workbooks.Add
  ws.Copy Before:=novySeznam2.Sheets(1)
  Application.DisplayAlerts = False ' Potlačí upozornění na přepsání souboru
  novySeznam2.SaveAs Filename:=cilovaCesta2, FileFormat:=xlOpenXMLWorkbook
  Application.DisplayAlerts = True ' Opět povolí upozornění
  novySeznam2.Close SaveChanges:=False
  
  
  '####################################################
  '#########################  Propsání seznamu odchylek
  
  Dim novySeznam3 As Workbook
  Dim cilovaCesta3 As String
  
  cilovaCesta3 = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS\TvaryObtiznost.xlsx"
  Set ws = ThisWorkbook.Sheets("Tvary")
  Set novySeznam3 = Workbooks.Add
  ws.Copy Before:=novySeznam3.Sheets(1)
  Application.DisplayAlerts = False ' Potlačí upozornění na přepsání souboru
  novySeznam3.SaveAs Filename:=cilovaCesta3, FileFormat:=xlOpenXMLWorkbook
  Application.DisplayAlerts = True ' Opět povolí upozornění
  novySeznam3.Close SaveChanges:=False
  
  
  '######################################################
  '#########################  Propsání seznamu Kov ANO NE
  
  Dim novySeznam4 As Workbook
  Dim cilovaCesta4 As String
  
  cilovaCesta4 = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS\KovAnoNe.xlsx"
  Set ws = ThisWorkbook.Sheets("Kovy")
  Set novySeznam4 = Workbooks.Add
  ws.Copy Before:=novySeznam4.Sheets(1)
  Application.DisplayAlerts = False ' Potlačí upozornění na přepsání souboru
  novySeznam4.SaveAs Filename:=cilovaCesta4, FileFormat:=xlOpenXMLWorkbook
  Application.DisplayAlerts = True ' Opět povolí upozornění
  novySeznam4.Close SaveChanges:=False
  
  '################################
  'Vypsání informace o stavu makra:
  
  Dim Zacatek As Date
  Zacatek = Worksheets("AKTUALIZACE").Range("A12").Value
  Dim Konec As Date
  Konec = Worksheets("AKTUALIZACE").Range("A14").Value
  
  StavAktualizace = " ------------ AKTUALIZACE DOKONČENA ----------- " & vbCrLf _
  & "--------------------------------------------------" & vbCrLf _
  & "DATUM POSLEDNÍ AKTUALIZACE:   " & den & "   " & datum & vbCrLf _
  & "--------------------------------------------------" & vbCrLf _
  & "Načtené  období:" & vbCrLf _
  & "od:   " & Zacatek & vbCrLf _
  & "do:   " & Konec
  
  'odmrazit obrazovku
  Application.ScreenUpdating = True
  Worksheets("AKTUALIZACE").Range("C16") = "OK"
  Worksheets("AKTUALIZACE").Protect Password:="123456"

End Sub
