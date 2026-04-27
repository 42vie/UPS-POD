#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <EditConstants.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>
#include <File.au3>

Opt("MouseCoordMode", 1)
Opt("PixelCoordMode", 1)
HotKeySet("{ESC}", "_Quitter")

; ===================== CONFIG =====================
Global Const $SLOW_FACTOR       = 2.2
Global Const $WAIT_UPS_SEC      = 45
Global Const $WAIT_DOWNLOAD_SEC = 60
Global Const $WAIT_PDF_SEC      = 60
Global Const $WAIT_UPLOAD_SEC   = 90

Global Const $DOWNLOADS_DIR = @UserProfileDir & "\Downloads"
Global Const $TEMP_DIR      = @TempDir & "\UPS_POD_WORK"
DirCreate($TEMP_DIR)

; Dossier source POD (sans upload)
Global Const $POD_SOURCE = @ScriptDir & "\POD_SOURCE"
DirCreate($POD_SOURCE)

; Edge path
Global Const $EDGE_EXE_1 = @ProgramFilesDir & "\Microsoft\Edge\Application\msedge.exe"
Global Const $EDGE_EXE_2 = @ProgramFilesDir & " (x86)\Microsoft\Edge\Application\msedge.exe"

; EDOC
Global Const $EDOC_VIEWER_TITLE = "edoc Viewer CDG"
Global Const $EDOC_UPLOAD_TITLE = "Upload Documents CDG"
Global Const $EDOC_UPLOAD_BTN   = "TButton2" ; à ajuster si besoin
Global Const $EDOC_TYPE_TEXT    = "POD - Proof of Delivery"
Global Const $PRINTER_EDOC      = "edoc Upload"

; ===================== DATA =======================
; 0=NumJ 1=Tracking 2=OK POD 3=Livré date 4=Livré heure 5=ETA 6=PDF 7=EDOC
Global $gRows[0][8]

; ===================== GUI ========================
Global $hGUI = GUICreate("UPS POD Tool - Tracking / Download POD / Upload EDOC / Export", 1200, 720)

GUICtrlCreateLabel("3 boutons séparés : Tracking / Télécharger POD (sans upload) / Upload EDOC (print) + Export Excel", 15, 10, 1100, 18)

Global $btnLoad     = GUICtrlCreateButton("Charger Excel", 15, 40, 150, 32)
Global $btnCheckAll = GUICtrlCreateButton("Tout cocher", 175, 40, 120, 32)
Global $btnUncheck  = GUICtrlCreateButton("Tout décocher", 305, 40, 120, 32)

Global $btnTrack    = GUICtrlCreateButton("Tracking", 445, 40, 120, 32)
Global $btnDownload = GUICtrlCreateButton("Télécharger les POD", 575, 40, 200, 32)
Global $btnUpload   = GUICtrlCreateButton("Upload EDOC (PRINT)", 785, 40, 180, 32)
Global $btnExport   = GUICtrlCreateButton("Export Excel", 975, 40, 150, 32)

Global $lv = GUICtrlCreateListView("Traiter|NumJ|Tracking|OK POD ?|Livré (date)|Livré (heure)|ETA|PDF|EDOC", 15, 90, 1170, 520)
_GUICtrlListView_SetExtendedListViewStyle($lv, BitOR($LVS_EX_CHECKBOXES, $LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))

Global $log = GUICtrlCreateEdit("", 15, 620, 1170, 80, BitOR($ES_READONLY, $WS_VSCROLL, $ES_AUTOVSCROLL))

GUISetState(@SW_SHOW)
_Log("Outil prêt. Charge ton Excel.")

; ===================== LOOP ========================
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            Exit

        Case $btnLoad
            _LoadExcel()

        Case $btnCheckAll
            _SetAllChecks(True)

        Case $btnUncheck
            _SetAllChecks(False)

        Case $btnTrack
            _RunTracking()

        Case $btnDownload
            _RunDownloadPOD_Only()

        Case $btnUpload
            _RunUploadEDOC_Print_FromSource()

        Case $btnExport
            _ExportTableToXlsx()
    EndSwitch
WEnd
Func _LoadExcel()
    Local $f = FileOpenDialog("Excel NumJ / Tracking (Col A = NumJ / Col B = Tracking)", @ScriptDir, "Excel (*.xlsx;*.xls)")
    If @error Or $f = "" Then Return

    Local $oExcel = _Excel_Open(False)
    If @error Then Return MsgBox(16, "Erreur", "Impossible d'ouvrir Excel (COM).")

    Local $oBook = _Excel_BookOpen($oExcel, $f)
    If @error Then
        _Excel_Close($oExcel)
        Return MsgBox(16, "Erreur", "Impossible d'ouvrir le fichier Excel.")
    EndIf

    Local $aData = _Excel_RangeRead($oBook)
    _Excel_BookClose($oBook, False)
    _Excel_Close($oExcel)

    If Not IsArray($aData) Then Return MsgBox(16, "Erreur", "Lecture Excel impossible (tableau vide).")
    If UBound($aData, 2) < 2 Then Return MsgBox(16, "Erreur", "Il faut au moins 2 colonnes (NumJ / Tracking).")

    ReDim $gRows[0][8]
    _GUICtrlListView_DeleteAllItems($lv)

    For $i = 0 To UBound($aData) - 1
        Local $numJ = $aData[$i][0]
        Local $trk  = $aData[$i][1]

        If $numJ = "" Or $trk = "" Then ContinueLoop
        If StringInStr($trk, "Tracking") Then ContinueLoop ; skip header

        ReDim $gRows[UBound($gRows) + 1][8]
        Local $r = UBound($gRows) - 1

        $gRows[$r][0] = $numJ
        $gRows[$r][1] = $trk
        $gRows[$r][2] = "" ; OK POD
        $gRows[$r][3] = "" ; date
        $gRows[$r][4] = "" ; heure
        $gRows[$r][5] = "" ; eta
        $gRows[$r][6] = "" ; pdf
        $gRows[$r][7] = "" ; edoc

        GUICtrlCreateListViewItem(" |" & $numJ & "|" & $trk & "|||||", $lv)
    Next

    _Log("Excel chargé : " & UBound($gRows) & " ligne(s).")
EndFunc

Func _UpdateRow($i)
    _GUICtrlListView_SetItemText($lv, $i, $gRows[$i][2], 3) ; OK POD
    _GUICtrlListView_SetItemText($lv, $i, $gRows[$i][3], 4) ; date
    _GUICtrlListView_SetItemText($lv, $i, $gRows[$i][4], 5) ; heure
    _GUICtrlListView_SetItemText($lv, $i, $gRows[$i][5], 6) ; eta
    _GUICtrlListView_SetItemText($lv, $i, $gRows[$i][6], 7) ; pdf
    _GUICtrlListView_SetItemText($lv, $i, $gRows[$i][7], 8) ; edoc
EndFunc

Func _GetCheckedIndices()
    Local $res[0]
    Local $count = _GUICtrlListView_GetItemCount($lv)
    For $i = 0 To $count - 1
        If _GUICtrlListView_GetItemChecked($lv, $i) Then
            ReDim $res[UBound($res) + 1]
            $res[UBound($res) - 1] = $i
        EndIf
    Next
    Return $res
EndFunc

Func _SetAllChecks($b)
    Local $count = _GUICtrlListView_GetItemCount($lv)
    For $i = 0 To $count - 1
        _GUICtrlListView_SetItemChecked($lv, $i, $b)
    Next
EndFunc

Func _Log($s)
    Local $old = GUICtrlRead($log)
    GUICtrlSetData($log, $old & @CRLF & _NowStamp() & " - " & $s)
EndFunc

Func _NowStamp()
    Return StringFormat("%02d:%02d:%02d", @HOUR, @MIN, @SEC)
EndFunc

Func _Slow($ms)
    Local $target = Int($ms * $SLOW_FACTOR)
    Local $t = TimerInit()
    While TimerDiff($t) < $target
        Sleep(80)
    WEnd
EndFunc

Func _Quitter()
    Exit
EndFunc

Func _SafeName($s)
    Local $x = $s
    $x = StringReplace($x, "\", "_")
    $x = StringReplace($x, "/", "_")
    $x = StringReplace($x, ":", "_")
    $x = StringReplace($x, "*", "_")
    $x = StringReplace($x, "?", "_")
    $x = StringReplace($x, """", "_")
    $x = StringReplace($x, "<", "_")
    $x = StringReplace($x, ">", "_")
    $x = StringReplace($x, "|", "_")
    Return $x
EndFunc
Func _RunTracking()
    If UBound($gRows) = 0 Then Return MsgBox(48, "Info", "Charge d'abord un Excel.")
    Local $idx = _GetCheckedIndices()
    If UBound($idx) = 0 Then Return MsgBox(48, "Info", "Coche au moins une ligne.")

    _Log("Tracking: lecture via title (pas de JSON).")
    For $k = 0 To UBound($idx) - 1
        Local $i = $idx[$k]
        Local $trk = $gRows[$i][1]

        Local $a = _UPS_Tracking_Read($trk)
        If IsArray($a) Then
            $gRows[$i][2] = $a[0] ; OK POD
            $gRows[$i][3] = $a[1] ; date
            $gRows[$i][4] = $a[2] ; heure
            $gRows[$i][5] = $a[3] ; eta
        Else
            $gRows[$i][2] = "Non"
        EndIf

        _UpdateRow($i)
        Sleep(200)
    Next

    _Log("Tracking terminé.")
EndFunc

Func _UPS_Tracking_Read($tracking)
    Local $url = "https://www.ups.com/track?loc=fr_FR&tracknum=" & $tracking & "&requester=ST/trackdetails"
    ShellExecute($url)

    If Not _WaitBrowserActive($WAIT_UPS_SEC) Then Return SetError(1, 0, 0)
    Local $hBrowser = WinGetHandle("[ACTIVE]")

    Local $js = _BuildJsTrackingToTitle($tracking)
    _DevToolsRunJS_Edge($hBrowser, $js) ; exécute le JS dans DevTools

    ; attendre le marqueur dans le titre
    Local $t = TimerInit()
    While TimerDiff($t) < 8000
        Local $title = WinGetTitle($hBrowser)
        If StringInStr($title, "UPS_TRACK§") Then
            Local $a = _ParseTrackTitle($title)
            If IsArray($a) Then Return $a
        EndIf
        Sleep(200)
    WEnd

    ; fallback TXT
    Local $jsTxt = _BuildJsTrackingToTxt($tracking)
    _DevToolsRunJS_Edge($hBrowser, $jsTxt)
    Local $txtPath = _WaitNewestDownloadMatching("UPS_TRACK_" & $tracking & "_", ".txt", 12)
    If $txtPath <> "" Then
        Local $s = FileRead($txtPath)
        FileDelete($txtPath)
        Local $a2 = _ParseTrackTxt($s)
        If IsArray($a2) Then Return $a2
    EndIf

    Return SetError(2, 0, 0)
EndFunc

Func _BuildJsTrackingToTitle($tracking)
    Local $js = _
    "(function(){" & _
    "function clean(x){return (x||'').replace(/[\r\n]+/g,' ').replace(/§/g,'/').replace(/\s+/g,' ').trim();}" & _
    "function setTitle(ok,dd,dt,eta){document.title='UPS_TRACK§" & $tracking & "§'+clean(ok)+'§'+clean(dd)+'§'+clean(dt)+'§'+clean(eta)+'§END';}" & _
    "function norm(s){return (s||'').replace(/\s+/g,' ').trim().toLowerCase();}" & _
    "function waitFor(getter,cb){var tries=0;var t=setInterval(function(){" & _
    " tries++;try{var v=getter();if(v){clearInterval(t);cb(v);}}catch(e){}" & _
    " if(tries>120){clearInterval(t);cb(null);} },250);}" & _
    "function openPOD(){try{" & _
    " var btn=document.getElementById('stApp_btnProofOfDeliveryonDetails');" & _
    " if(btn){btn.click();return true;}" & _
    " var b=[...document.querySelectorAll('button,a,span')].find(x=>norm(x.textContent)==='preuve de livraison');" & _
    " if(b){b.click();return true;}" & _
    "}catch(e){} return false;}" & _
    "function byLabel(modal,lbl){" & _
    " var target=norm(lbl);" & _
    " var nodes=modal.querySelectorAll('span,div,dt,strong,p,li');" & _
    " for(var i=0;i<nodes.length;i++){" & _
    "  if(norm(nodes[i].textContent)===target){" & _
    "   var v='';" & _
    "   if(nodes[i].nextElementSibling) v=(nodes[i].nextElementSibling.innerText||'');" & _
    "   if(!v && nodes[i].parentElement && nodes[i].parentElement.nextElementSibling) v=(nodes[i].parentElement.nextElementSibling.innerText||'');" & _
    "   return (v||'').trim();" & _
    "  }" & _
    " }" & _
    " return '';" & _
    "}" & _
    "function extractDelivered(modal){" & _
    " var raw=byLabel(modal,'Livré le');" & _
    " var dd='', dt='';" & _
    " if(raw){" & _
    "  dd=raw;" & _
    "  var m=raw.match(/(\d{1,2}:\d{2})/);" & _
    "  if(m) dt='à '+m[1];" & _
    "  return {dd:dd,dt:dt};" & _
    " }" & _
    " var t=(modal.innerText||'');" & _
    " var m2=t.match(/Livr[ée]\s*le[^\n]*?(\d{1,2}:\d{2})/i);" & _
    " if(m2) dt='à '+m2[1];" & _
    " return {dd:'',dt:dt};" & _
    "}" & _
    "openPOD();" & _
    "waitFor(function(){return document.getElementById('stApp_podModal');}, function(modal){" & _
    " if(!modal){setTitle('Non','','','');return;}" & _
    " var d=extractDelivered(modal);" & _
    " var ok=(d.dt?'Oui':'Non');" & _
    " setTitle(ok,d.dd,d.dt,'');" & _
    "});" & _
    "})();"

    Return $js
EndFunc

Func _BuildJsTrackingToTxt($tracking)
    Local $js = _
        "(function(){" & _
        "function clean(x){return (x||'').replace(/[\r\n]+/g,' ').replace(/§/g,'/').replace(/\s+/g,' ').trim();}" & _
        "var body=(document.body?document.body.innerText:'')||'';" & _
        "var deliveredDate='';var deliveredTime='';var eta='';" & _
        "var m=body.match(/Livr[ée]\s+le\s*([^\n\r]+)/i);" & _
        "if(m&&m[1]){deliveredDate=m[1].trim();var m2=m[1].match(/à\s*(\d{1,2}:\d{2})/i);if(m2&&m2[1])deliveredTime='à '+m2[1];}" & _
        "var e1=body.match(/entre\s*(\d{1,2}:\d{2})\s*(?:-|et|à)\s*(\d{1,2}:\d{2})/i);" & _
        "if(e1)eta='entre '+e1[1]+' - '+e1[2];" & _
        "var ok=(deliveredTime?'Oui':'Non');" & _
        "var payload=clean(ok)+'§'+clean(deliveredDate)+'§'+clean(deliveredTime)+'§'+clean(eta);" & _
        "var blob=new Blob([payload],{type:'text/plain;charset=utf-8'});" & _
        "var a=document.createElement('a');a.href=URL.createObjectURL(blob);" & _
        "var ts=new Date().toISOString().slice(0,19).replace(/[:T]/g,'-');" & _
        "a.download='UPS_TRACK_" & $tracking & "_'+ts+'.txt';document.body.appendChild(a);a.click();a.remove();" & _
        "setTimeout(function(){URL.revokeObjectURL(a.href);},2000);" & _
        "})();"
    Return $js
EndFunc

Func _ParseTrackTitle($title)
    Local $pos = StringInStr($title, "UPS_TRACK§")
    If $pos = 0 Then Return SetError(1, 0, 0)

    Local $s = StringMid($title, $pos)

    ; coupe UNIQUEMENT à §END (évite le bug avec " - Livraison interne")
    Local $endPos = StringInStr($s, "§END")
    If $endPos > 0 Then
        $s = StringLeft($s, $endPos - 1)
    EndIf

    $s = StringRegExpReplace($s, "[\x00-\x1F]", " ")
    Local $a = StringSplit($s, "§", 2)
    If Not IsArray($a) Or UBound($a) < 6 Then Return SetError(2, 0, 0)
    If $a[0] <> "UPS_TRACK" Then Return SetError(3, 0, 0)

    Local $out[4]
    $out[0] = $a[2] ; Oui/Non
    $out[1] = $a[3] ; date
    $out[2] = $a[4] ; heure
    $out[3] = $a[5] ; ETA
    Return $out
EndFunc


Func _DevToolsRunJS_Edge($hBrowser, $sJS)
    If $hBrowser = 0 Then Return False

    WinActivate($hBrowser)
    WinWaitActive($hBrowser, "", 10)
    _Slow(900)

    SendKeepActive($hBrowser)

    ; Ouvre DevTools console
    Send("^+j")
    _Slow(2300)

    ; Nettoie la console input + colle
    ClipPut($sJS)
    _Slow(250)

    Send("^a")     ; important: évite un collage incomplet dans du texte existant
    _Slow(120)
    Send("^v")
    _Slow(250)

    ; Exécution forcée
    Send("^{ENTER}")
    _Slow(800)

    SendKeepActive("")
    Return True
EndFunc

Func _WaitBrowserActive($timeoutSec)
    Local $t = TimerInit()
    While TimerDiff($t) < ($timeoutSec * 1000)
        If WinActive("[CLASS:Chrome_WidgetWin_1]") Then Return True
        Sleep(250)
    WEnd
    Return False
EndFunc

Func _WaitNewestDownloadMatching($prefix, $ext, $timeoutSec)
    Local $t = TimerInit()
    Local $best = ""
    Local $bestTime = 0

    While TimerDiff($t) < ($timeoutSec * 1000)
        Local $h = FileFindFirstFile($DOWNLOADS_DIR & "\" & $prefix & "*" & $ext)
        If $h <> -1 Then
            While 1
                Local $f = FileFindNextFile($h)
                If @error Then ExitLoop
                If StringRight($f, 11) = ".crdownload" Then ContinueLoop

                Local $full = $DOWNLOADS_DIR & "\" & $f
                Local $mtime = FileGetTime($full, 0, 1)
                If $mtime > $bestTime Then
                    $bestTime = $mtime
                    $best = $full
                EndIf
            WEnd
            FileClose($h)
        EndIf

        If $best <> "" And FileGetSize($best) > 50 Then Return $best
        Sleep(250)
    WEnd
    Return ""
EndFunc
Func _RunDownloadPOD_Only()
    If UBound($gRows) = 0 Then Return MsgBox(48, "Info", "Charge d'abord un Excel.")
    Local $idx = _GetCheckedIndices()
    If UBound($idx) = 0 Then Return MsgBox(48, "Info", "Coche au moins une ligne.")

    Local $useFilter = False
    For $k = 0 To UBound($idx) - 1
        If $gRows[$idx[$k]][2] <> "" Then
            $useFilter = True
            ExitLoop
        EndIf
    Next

    _Log("Téléchargement POD (sans upload). Dossier source: " & $POD_SOURCE)

    For $k = 0 To UBound($idx) - 1
        Local $i = $idx[$k]
        Local $numJ = $gRows[$i][0]
        Local $trk  = $gRows[$i][1]

        If $useFilter And $gRows[$i][2] = "Non" Then
            $gRows[$i][6] = "IGNORÉ"
            _UpdateRow($i)
            ContinueLoop
        EndIf

        Local $outDir = $POD_SOURCE & "\" & _SafeName($numJ)
        DirCreate($outDir)

        Local $pdfPath = $outDir & "\POD_" & $trk & ".pdf"
        $gRows[$i][6] = "En cours..."
        _UpdateRow($i)

        If _UPS_GeneratePodPdf($trk, $pdfPath) Then
            $gRows[$i][6] = "PDF OK"
        Else
            $gRows[$i][6] = "KO PDF"
        EndIf
        _UpdateRow($i)
        Sleep(250)
    Next

    _Log("Téléchargement POD terminé (sans upload).")
EndFunc

Func _UPS_GeneratePodPdf($tracking, $pdfOutPath)
    ; UPS -> POD modal -> download HTML -> Edge headless -> PDF -> delete HTML
    Local $url = "https://www.ups.com/track?loc=fr_FR&tracknum=" & $tracking & "&requester=ST/trackdetails"
    ShellExecute($url)

    If Not _WaitBrowserActive($WAIT_UPS_SEC) Then Return False
    Local $hBrowser = WinGetHandle("[ACTIVE]")

    Local $js = _BuildJsExportPodHtml($tracking)
    If Not _DevToolsRunJS_Edge($hBrowser, $js) Then Return False

    Local $htmlPath = _WaitNewestDownloadMatching("POD_" & $tracking & "_", ".html", $WAIT_DOWNLOAD_SEC)
    If $htmlPath = "" Then Return False

    If FileExists($pdfOutPath) Then FileDelete($pdfOutPath)
    Local $ok = _EdgePrintToPdf($htmlPath, $pdfOutPath)

    If FileExists($htmlPath) Then FileDelete($htmlPath)
    Return $ok
EndFunc

Func _BuildJsExportPodHtml($tracking)
    Local $js = _
    "(function(){" & _
    "function waitFor(getter,cb){var tries=0;var i=setInterval(function(){" & _
    " tries++; try{var e=getter(); if(e){clearInterval(i); cb(e);} }catch(err){}" & _
    " if(tries>140){clearInterval(i); cb(null);} },250);} " & _

    "function norm(s){return (s||'').replace(/\s+/g,' ').trim().toLowerCase();}" & _
    "function esc(s){s=(s||''); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}" & _

    "function byLabel(modal,lbl){" & _
    " var target=norm(lbl);" & _
    " var nodes=modal.querySelectorAll('span,div,dt,strong,p,li');" & _
    " for(var i=0;i<nodes.length;i++){" & _
    "   if(norm(nodes[i].textContent)===target){" & _
    "     var v='';" & _
    "     if(nodes[i].nextElementSibling) v=(nodes[i].nextElementSibling.innerText||'');" & _
    "     if(!v && nodes[i].parentElement && nodes[i].parentElement.nextElementSibling) v=(nodes[i].parentElement.nextElementSibling.innerText||'');" & _
    "     return (v||'').trim();" & _
    "   }" & _
    " }" & _
    " return '';" & _
    "}" & _

    "function openPOD(){" & _
    " var btn=document.getElementById('stApp_btnProofOfDeliveryonDetails');" & _
    " if(btn){btn.click(); return true;}" & _
    " var b=[...document.querySelectorAll('button,a,span')].find(x=>norm(x.textContent)==='preuve de livraison');" & _
    " if(b){b.click(); return true;}" & _
    " return false;" & _
    "}" & _

    "openPOD();" & _
    "waitFor(function(){return document.getElementById('stApp_podModal');}, function(modal){" & _
    " if(!modal){console.log('POD modal not found'); return;}" & _

    " var trackingVal = byLabel(modal,'Numéro de suivi') || '" & $tracking & "';" & _
    " var service     = byLabel(modal,'Service');" & _
    " var weight      = byLabel(modal,'Poids');" & _
    " var billedOn    = byLabel(modal,'Envoyé / facturé le');" & _
    " var deliveredOn = byLabel(modal,'Livré le');" & _
    " var receivedBy  = byLabel(modal,'Réceptionné par');" & _
    " var deliveredTo = byLabel(modal,'Livré à');" & _
    " var deliveryLoc = byLabel(modal,'Adresse de livraison');" & _
    " var refs        = byLabel(modal,'Numéro(s) de référence');" & _

    " var signHtml='Signature non disponible';" & _
    " var signData='';" & _
    " var img=modal.querySelector('img[src^=""data:image""]');" & _
    " if(img && img.src) signData=img.src;" & _
    " if(!signData){" & _
    "   var canvas=modal.querySelector('canvas');" & _
    "   if(canvas){try{signData=canvas.toDataURL('image/png');}catch(e){signData='';}}" & _
    " }" & _
    " if(signData) signHtml='<img style=""max-width:320px;border:1px solid #999;padding:6px"" src=""'+signData+'"" />';" & _

    " var now=new Date();" & _
    " var ts=now.toISOString().slice(0,19).replace(/[:T]/g,'-');" & _
    " var gen=now.toLocaleString('fr-FR');" & _

    " var h=[];" & _
    " h.push('<!doctype html><html lang=""fr""><head><meta charset=""utf-8""><title>POD '+esc(trackingVal)+'</title>');" & _
    " h.push('<style>body{font-family:Arial;margin:28px;color:#000;}h1{font-size:20px;margin:0 0 18px;}');" & _
    " h.push('.box{border:1px solid #333;padding:18px;border-radius:10px;} .row{display:flex;gap:30px;margin:14px 0;}');" & _
    " h.push('.col{flex:1;} .k{font-weight:bold;font-size:13px;} .v{margin-top:4px;font-size:14px;white-space:pre-wrap;}');" & _
    " h.push('.small{font-size:11px;color:#444;margin-top:18px;line-height:1.4;}</style></head><body>');" & _
    " h.push('<h1>Preuve de livraison officielle UPS</h1><div class=""box"">');" & _
    " function row(aK,aV,bK,bV){h.push('<div class=""row""><div class=""col""><div class=""k"">'+aK+'</div><div class=""v"">'+aV+'</div></div>'+" & _
    " '<div class=""col""><div class=""k"">'+bK+'</div><div class=""v"">'+bV+'</div></div></div>');}" & _
    " row('Numéro de suivi', esc(trackingVal), 'Service', esc(service));" & _
    " row('Poids', esc(weight), 'Envoyé / facturé le', esc(billedOn));" & _
    " row('Livré le', esc(deliveredOn), 'Réceptionné par', esc(receivedBy));" & _
    " row('Livré à', esc(deliveredTo), 'Adresse de livraison', esc(deliveryLoc));" & _
    " row('Numéro(s) de référence', esc(refs), 'Signature', signHtml);" & _
    " h.push('<div class=""small"">Données officielles UPS. Généré automatiquement le '+esc(gen)+'.</div>');" & _
    " h.push('</div></body></html>');" & _
    " var html=h.join('');" & _
    " var blob=new Blob([html],{type:'text/html;charset=utf-8'});" & _
    " var a=document.createElement('a');a.href=URL.createObjectURL(blob);" & _
    " a.download='POD_" & $tracking & "_'+ts+'.html';document.body.appendChild(a);a.click();a.remove();" & _
    " setTimeout(function(){URL.revokeObjectURL(a.href);},2000);" & _
    "});" & _
    "})();"

    Return $js
EndFunc
Func _EdgePrintToPdf($htmlPath, $pdfPath)
    Local $edge = ""
    If FileExists($EDGE_EXE_1) Then
        $edge = $EDGE_EXE_1
    ElseIf FileExists($EDGE_EXE_2) Then
        $edge = $EDGE_EXE_2
    Else
        _Log("msedge.exe introuvable.")
        Return False
    EndIf

    Local $profile = $TEMP_DIR & "\EdgeProfile"
    DirCreate($profile)

    Local $fileUrl = _PathToFileUrlEncoded($htmlPath)
    Local $args = ' --headless=old --disable-gpu --user-data-dir="' & $profile & '"' & _
                  ' --virtual-time-budget=9000 --print-to-pdf="' & $pdfPath & '" "' & $fileUrl & '"'

    RunWait('"' & $edge & '"' & $args, "", @SW_HIDE)

    Local $t = TimerInit()
    While TimerDiff($t) < ($WAIT_PDF_SEC * 1000)
        If FileExists($pdfPath) And FileGetSize($pdfPath) > 2000 Then Return True
        Sleep(250)
    WEnd
    Return False
EndFunc

Func _PathToFileUrlEncoded($path)
    Local $p = StringReplace($path, "\", "/")
    $p = StringReplace($p, " ", "%20")
    Return "file:///" & $p
EndFunc
; =================================================================================
; 5) UPLOAD EDOC (PRINT) DEPUIS POD_SOURCE (NumJ par NumJ)
; =================================================================================

Func _RunUploadEDOC_Print_FromSource()
    _Log("Upload EDOC (PRINT) : scan POD_SOURCE -> NumJ par NumJ")

    ; Défaut imprimante = edoc Upload
    Local $oNet = ObjCreate("WScript.Network")
    $oNet.SetDefaultPrinter($PRINTER_EDOC)

    ; Liste des sous-dossiers NumJ
    Local $aFolders = _ListSubFolders($POD_SOURCE)
    If Not IsArray($aFolders) Then
        _Log("Aucun dossier NumJ trouvé dans " & $POD_SOURCE)
        Return
    EndIf

    For $f = 0 To UBound($aFolders) - 1
        Local $numJFolder = $aFolders[$f]
        Local $folderPath = $POD_SOURCE & "\" & $numJFolder

        ; Liste des PDFs dans le dossier NumJ
        Local $aPdfs = _ListPdfs($folderPath)
        If Not IsArray($aPdfs) Then ContinueLoop

        ; 1) Précharger dossier EDOC Viewer
        _EDocPreload_NumJ($numJFolder)

        ; 2) Print + validation pour chaque PDF
        For $p = 0 To UBound($aPdfs) - 1
            Local $pdfName = $aPdfs[$p]
            Local $pdfFull = $folderPath & "\" & $pdfName

            _Log("PRINT -> " & $pdfFull)
            ShellExecute($pdfFull, "", "", "print", @SW_HIDE)

            ; Scanner: attendre que Upload Documents CDG apparaisse et soit prêt
            If Not _ScanWaitUploadReady($WAIT_UPLOAD_SEC) Then
                _MovePdfToKO($folderPath, $pdfName)
                _MarkEdocStatusByPdfName($pdfName, "Upload KO")
                ContinueLoop
            EndIf

            ; Validation EDOC après print (fenêtre se ferme si OK)
            Local $ok = _ValidateUploadWindow_AfterPrint_WaitClose(35)

            If $ok Then
                _MarkEdocStatusByPdfName($pdfName, "Upload OK")
                _SafeDeleteWithRetry($pdfFull, 6, 600)
            Else
                _MovePdfToKO($folderPath, $pdfName)
                _MarkEdocStatusByPdfName($pdfName, "Upload KO")
            EndIf

            Sleep(800)
        Next
    Next

    _Log("Upload EDOC (PRINT) terminé.")
EndFunc

Func _EDocPreload_NumJ($numJ)
    If Not WinExists($EDOC_VIEWER_TITLE) Then
        _Log("EDOC Viewer introuvable : " & $EDOC_VIEWER_TITLE)
        Return
    EndIf

    WinActivate($EDOC_VIEWER_TITLE)
    WinWaitActive($EDOC_VIEWER_TITLE, "", 10)

    ; Boucliers anti-crash comme ton style
    Sleep(500)
    ControlSetText($EDOC_VIEWER_TITLE, "", "Edit1", "")
    Sleep(400)
    ControlSetText($EDOC_VIEWER_TITLE, "", "Edit1", $numJ)
    Sleep(400)
    ControlSend($EDOC_VIEWER_TITLE, "", "Edit1", "{ENTER}")

    ; EDOC est lent à charger le dossier
    Sleep(2500)
EndFunc

Func _ValidateUploadWindow_AfterPrint_WaitClose($timeoutSec)
    WinActivate($EDOC_UPLOAD_TITLE)
    WinWaitActive($EDOC_UPLOAD_TITLE, "", 10)

    ; Bouclier: EDOC génère son image/preview après print
    Sleep(1800)

    ; Type doc (TAB + texte)
    Send("{TAB}")
    Sleep(350)
    Send($EDOC_TYPE_TEXT)
    Sleep(450)

    ; Upload
    ControlClick($EDOC_UPLOAD_TITLE, "", $EDOC_UPLOAD_BTN)

    ; Si OK => la fenêtre se ferme (tu as confirmé que c'est le cas)
    If WinWaitClose($EDOC_UPLOAD_TITLE, "", $timeoutSec) Then
        Sleep(400)
        Return True
    EndIf

    Return False
EndFunc

Func _ScanWaitUploadReady($timeoutSec)
    ; Scanner pratique: WinWait + contrôles présents + stabilité UI (PixelChecksum)
    Local $t = TimerInit()

    While TimerDiff($t) < ($timeoutSec * 1000)
        If WinExists($EDOC_UPLOAD_TITLE) Then
            Local $state = WinGetState($EDOC_UPLOAD_TITLE)
            If BitAND($state, 2) Then ; visible
                WinActivate($EDOC_UPLOAD_TITLE)
                WinWaitActive($EDOC_UPLOAD_TITLE, "", 5)

                ; Vérifie que le bouton Upload est présent
                Local $hUpload = ControlGetHandle($EDOC_UPLOAD_TITLE, "", $EDOC_UPLOAD_BTN)
                If $hUpload <> "" Then
                    ; Stabilité UI (évite d'écrire trop tôt)
                    Local $pos = WinGetPos($EDOC_UPLOAD_TITLE)
                    If IsArray($pos) Then
                        Local $chk1 = PixelChecksum($pos[0] + 10, $pos[1] + 10, $pos[0] + 180, $pos[1] + 70)
                        Sleep(250)
                        Local $chk2 = PixelChecksum($pos[0] + 10, $pos[1] + 10, $pos[0] + 180, $pos[1] + 70)
                        If $chk1 = $chk2 Then
                            Sleep(800)
                            Return True
                        EndIf
                    Else
                        Sleep(800)
                        Return True
                    EndIf
                EndIf
            EndIf
        EndIf

        Sleep(250)
    WEnd

    Return False
EndFunc
; =================================================================================
; 6) HELPERS (SCAN SOURCE / KO / DELETE SAFE / MAJ EDOC) + EXPORT EXCEL
; =================================================================================

Func _ListSubFolders($base)
    Local $h = FileFindFirstFile($base & "\*")
    If $h = -1 Then Return SetError(1,0,0)

    Local $arr[0]
    While 1
        Local $name = FileFindNextFile($h)
        If @error Then ExitLoop
        If $name = "." Or $name = ".." Then ContinueLoop

        If StringInStr(FileGetAttrib($base & "\" & $name), "D") Then
            ReDim $arr[UBound($arr) + 1]
            $arr[UBound($arr) - 1] = $name
        EndIf
    WEnd
    FileClose($h)

    If UBound($arr) = 0 Then Return SetError(2,0,0)
    Return $arr
EndFunc

Func _ListPdfs($folder)
    Local $h = FileFindFirstFile($folder & "\*.pdf")
    If $h = -1 Then Return SetError(1,0,0)

    Local $arr[0]
    While 1
        Local $name = FileFindNextFile($h)
        If @error Then ExitLoop
        ReDim $arr[UBound($arr) + 1]
        $arr[UBound($arr) - 1] = $name
    WEnd
    FileClose($h)

    If UBound($arr) = 0 Then Return SetError(2,0,0)
    Return $arr
EndFunc

Func _MovePdfToKO($folderPath, $pdfName)
    Local $koDir = $folderPath & "\KO"
    DirCreate($koDir)

    Local $src = $folderPath & "\" & $pdfName
    Local $dst = $koDir & "\" & $pdfName

    If FileExists($src) Then
        FileMove($src, $dst, 9)
    EndIf
EndFunc

Func _SafeDeleteWithRetry($path, $tries, $sleepMs)
    For $i = 1 To $tries
        If Not FileExists($path) Then Return True
        FileDelete($path)
        Sleep($sleepMs)
    Next
    Return (Not FileExists($path))
EndFunc

Func _MarkEdocStatusByPdfName($pdfName, $status)
    ; pdfName attendu : POD_<tracking>.pdf
    Local $trk = ""
    If StringLeft($pdfName, 4) = "POD_" Then
        $trk = StringTrimRight(StringTrimLeft($pdfName, 4), 4) ; retire POD_ et .pdf
    EndIf
    If $trk = "" Then Return

    For $i = 0 To UBound($gRows) - 1
        If $gRows[$i][1] = $trk Then
            $gRows[$i][7] = $status
            _UpdateRow($i)
            ExitLoop
        EndIf
    Next
EndFunc

Func _ExportTableToXlsx()
    If UBound($gRows) = 0 Then
        MsgBox(48, "Info", "Rien à exporter (charge un Excel d'abord).")
        Return
    EndIf

    Local $stamp = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & @MIN & @SEC
    Local $outPath = @ScriptDir & "\EXPORT_UPS_POD_" & $stamp & ".xlsx"

    Local $rCount = UBound($gRows) + 1
    Local $data[$rCount][9]

    ; Headers
    $data[0][0] = "NumJ"
    $data[0][1] = "Tracking"
    $data[0][2] = "OK POD ?"
    $data[0][3] = "Livré (date)"
    $data[0][4] = "Livré (heure)"
    $data[0][5] = "ETA"
    $data[0][6] = "PDF"
    $data[0][7] = "EDOC"
    $data[0][8] = "Traiter (coché)"

    For $i = 0 To UBound($gRows) - 1
        $data[$i+1][0] = $gRows[$i][0]
        $data[$i+1][1] = $gRows[$i][1]
        $data[$i+1][2] = $gRows[$i][2]
        $data[$i+1][3] = $gRows[$i][3]
        $data[$i+1][4] = $gRows[$i][4]
        $data[$i+1][5] = $gRows[$i][5]
        $data[$i+1][6] = $gRows[$i][6]
        $data[$i+1][7] = $gRows[$i][7]

        Local $checked = _GUICtrlListView_GetItemChecked($lv, $i)
        If $checked Then
            $data[$i+1][8] = "Oui"
        Else
            $data[$i+1][8] = "Non"
        EndIf
    Next

    Local $oExcel = _Excel_Open(False)
    If @error Then
        MsgBox(16, "Erreur", "Impossible d'ouvrir Excel (COM).")
        Return
    EndIf

    Local $oBook = _Excel_BookNew($oExcel)
    If @error Then
        _Excel_Close($oExcel)
        MsgBox(16, "Erreur", "Impossible de créer un classeur Excel.")
        Return
    EndIf

    _Excel_RangeWrite($oBook, Default, $data, "A1")

    ; 51 = xlOpenXMLWorkbook (.xlsx)
    Local Const $xlOpenXMLWorkbook = 51
    _Excel_BookSaveAs($oBook, $outPath, $xlOpenXMLWorkbook, True)

    _Excel_BookClose($oBook, False)
    _Excel_Close($oExcel)

    MsgBox(64, "Export OK", "Export Excel créé :" & @CRLF & $outPath)
EndFunc
