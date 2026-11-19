; @altrun.title          OCR - screenshot to clipboard
; @altrun.keyword        ocr
; @altrun.mode           silent
; @altrun.author         sfcboy
; @altrun.description    Select a screen area and recognize its text with online OCR (better accuracy); result goes to the clipboard. Falls back to free offline Windows OCR if no API key or the online call fails.
; @altrun.description_zh 框选屏幕区域, 用在线 OCR 识别文字 (准确率更高), 结果复制到剪贴板; 未配置密钥或在线失败 (断网等) 时自动回退到免费离线的 Windows OCR
; @altrun.network        visual.volcengineapi.com
; @altrun.requires       Optional, for online OCR: user env vars VOLC_AK / VOLC_SK (use an IAM sub-user allowed only for OCR; screenshots are uploaded to Volcengine, so avoid confidential drawings). Without them: offline Windows OCR (Windows 10 1809+ with an OCR language pack).
; @altrun.requires_zh    在线 OCR 可选: 设置用户环境变量 VOLC_AK / VOLC_SK (用只授权 OCR 的 IAM 子用户; 截图会上传到火山引擎, 涉密图纸请勿使用). 未配置则用离线 Windows OCR (需 Windows 10 1809+ 和 OCR 语言包).
;
; HOW TO USE / 用法
;   Recommended: Preferences > Hotkeys > Add..., action "script:OCR.ahk", e.g. Ctrl+Alt+O.
;   Press the hotkey, drag over the text on screen, release: the text is copied to your clipboard and a notification shows how many characters were found. Press Esc to cancel.
;   The hotkey starts region selection immediately - it does not press Ctrl+C on the current window (safe in CAD, where Ctrl+C is "copy object").
;   Or type "ocr" in ALTRun and press Enter.
;   推荐: 偏好设置 > 热键 > 添加, 动作选 "script:OCR.ahk" (例如 Ctrl+Alt+O). 按热键后直接拖拽框选屏幕上的文字, 松开即复制到剪贴板并通知字符数; Esc 取消. 按热键会直接开始框选, 不会读取当前窗口的选中内容 (在 CAD 里不会触发复制对象命令). 也可以在 ALTRun 里输入 "ocr" 回车.
;
; LANGUAGE / 语言
;   Set ALTRUN_OCR_LANG to force a language tag for the offline engine, e.g. zh-Hans-CN, en-US, ja (empty = Windows user-profile languages; the online engine ignores it).
;   可用 ALTRUN_OCR_LANG 指定离线引擎的识别语言, 留空则用系统语言 (在线引擎忽略此设置).
; ENGINE / 引擎
;   Default engine is online (Volcengine OCRNormal, visual.volcengineapi.com) for much better accuracy.
;   Setup: set user env vars VOLC_AK and VOLC_SK. No ALTRun restart is needed: the script reads them from the process environment first, then from HKCU\Environment in the registry.
;   Use a dedicated IAM sub-user allowed only for OCR calls; do not use the main account's keys. Screenshots are uploaded to Volcengine - do not use this on confidential drawings.
;   No Python / pip / volcengine SDK required: the Volcengine V4 signature and the HTTP call are done in PowerShell (built into Windows).
;   If VOLC_AK / VOLC_SK are unset, or the online call fails (no network, bad key,
;   quota exhausted), it silently falls back to the offline Windows OCR engine.
;   Force offline with ALTRUN_OCR_ENGINE=windows.
;   默认用在线引擎 (火山引擎 OCRNormal, visual.volcengineapi.com), 准确率更高.
;   配置: 设置用户环境变量 VOLC_AK / VOLC_SK. 不用重启 ALTRun - 先从进程环境读, 读不到再从注册表 HKCU\Environment 读.
;   建议用只授权 OCR 的 IAM 子用户, 不要用主账号密钥; 截图会上传到火山引擎, 涉密图纸请勿使用.
;   不需要 Python / pip / volcengine SDK - V4 签名和 HTTP 调用都在 Windows 自带的 PowerShell 里完成.
;   未设置密钥或在线识别失败 (断网/密钥错/额度用尽) 时自动回退到离线 Windows OCR. 设 ALTRUN_OCR_ENGINE=windows 可强制离线.
;
#Requires AutoHotkey v2.0
#NoTrayIcon

;-------------------------------------------------------------------------------
; Config
;-------------------------------------------------------------------------------
; Read credentials from the process environment first, then fall back to the
; persisted user registry value. Windows does NOT refresh an already-running
; process' environment block after SetEnvironmentVariable(..., "User"), so
; EnvGet can be empty even though the key is correctly stored in the registry.
ak := EnvGet("VOLC_AK")
sk := EnvGet("VOLC_SK")
if (ak = "") {
    try {
        ak := RegRead("HKEY_CURRENT_USER\Environment", "VOLC_AK")
    } catch {
        ak := ""
    }
}
if (sk = "") {
    try {
        sk := RegRead("HKEY_CURRENT_USER\Environment", "VOLC_SK")
    } catch {
        sk := ""
    }
}
; Inject into this process so the child PowerShell inherits them.
if (ak != "")
    EnvSet("VOLC_AK", ak)
if (sk != "")
    EnvSet("VOLC_SK", sk)

engine := EnvGet("ALTRUN_OCR_ENGINE")
if (engine = "") {
    ; Same registry fallback as the keys, so switching to offline needs no ALTRun restart.
    try {
        engine := RegRead("HKEY_CURRENT_USER\Environment", "ALTRUN_OCR_ENGINE")
    } catch {
        engine := ""
    }
}
if (engine = "")
    engine := "online"                               ; online OCR is far more accurate than Windows OCR
fallbackReason := ""                                ; empty = no fallback; otherwise why offline OCR was used
if (engine = "online" && (ak = "" || sk = "")) {
    engine := "windows"                             ; no credentials -> silently use offline OCR
    fallbackReason := "no API key"
}
lang := EnvGet("ALTRUN_OCR_LANG")

; Temporary files. Every name includes this process' PID, so two concurrent
; runs never collide (the first to exit cannot delete the other's engine).
; All are removed on exit (including on error / Esc), leaving nothing in %TEMP%.
pid := DllCall("GetCurrentProcessId")                ; AHK v2 has no built-in PID variable; ask Windows directly
enginePs := A_Temp "\altrun-ocr-" pid "-Engine.ps1"  ; offline Windows OCR engine (generated)
enginePsOnline := A_Temp "\altrun-ocr-" pid "-Online.ps1"  ; online Volcengine OCR engine (generated)
tmpPng := A_Temp "\altrun-ocr-" pid ".png"
tmpTxt := A_Temp "\altrun-ocr-" pid ".txt"
OnExit(CleanupTemp)

; No @altrun.selection / argument, so ALTRun never passes a file to this script.
; Always start by selecting a screen region.
if !SelectRegion(&x, &y, &w, &h)                     ; Esc or a too-small selection cancels silently
    ExitApp
if !CaptureRegion(x, y, w, h, tmpPng)
    Fail("Could not capture the screen")

text := Ocr(tmpPng, enginePs, enginePsOnline, lang)

if (text = "") {
    FileAppend("OCR: no text found", "*", "UTF-8-RAW")
    ExitApp
}

A_Clipboard := text                                  ; the result is ready to paste anywhere
; Keep the notification pure ASCII. ALTRun's ChildProcess decodes stdout as UTF-8
; only when the whole buffer is valid UTF-8; other bytes can force an OEM/GBK
; decode and garble Chinese. The full recognized text is on the clipboard already.
notice := "OCR copied " StrLen(text) " chars - result on clipboard"
if (fallbackReason != "")
    notice .= " (offline - " fallbackReason ")"
FileAppend(notice, "*", "UTF-8-RAW")                 ; silent mode shows the last line as a notification

;-------------------------------------------------------------------------------
; Remove our own temporary files on every exit path (success, error, Esc).
;-------------------------------------------------------------------------------
CleanupTemp(*) {
    global tmpPng, tmpTxt, enginePs, enginePsOnline
    try FileDelete(tmpPng)
    try FileDelete(tmpTxt)
    try FileDelete(enginePs)
    try FileDelete(enginePsOnline)
}

;-------------------------------------------------------------------------------
; Select a screen region with a translucent overlay; returns false when cancelled
; -DPIScale keeps the blue box aligned with the captured area at non-100% scaling.
;-------------------------------------------------------------------------------
SelectRegion(&x, &y, &w, &h) {
    ; Virtual-screen bounds, so all monitors are covered
    vx := DllCall("GetSystemMetrics", "Int", 76)     ; SM_XVIRTUALSCREEN
    vy := DllCall("GetSystemMetrics", "Int", 77)     ; SM_YVIRTUALSCREEN
    vw := DllCall("GetSystemMetrics", "Int", 78)     ; SM_CXVIRTUALSCREEN
    vh := DllCall("GetSystemMetrics", "Int", 79)     ; SM_CYVIRTUALSCREEN

    bg := Gui("-Caption +ToolWindow +AlwaysOnTop -DPIScale", "ALTRun OCR Selection")
    bg.BackColor := "000000"
    bg.Show("x" vx " y" vy " w" vw " h" vh)
    WinSetTransparent(90, bg.Hwnd)

    sel := Gui("-Caption +ToolWindow +AlwaysOnTop -DPIScale", "ALTRun OCR Region")
    sel.BackColor := "00AAFF"
    sel.Show("x" vx " y" vy " w1 h1 NoActivate")
    WinSetTransparent(70, sel.Hwnd)

    CoordMode("Mouse", "Screen")
    dragging := false
    sx := 0, sy := 0, mx := 0, my := 0
    while true {
        if GetKeyState("Esc", "P") {
            bg.Destroy(), sel.Destroy()
            return false
        }
        MouseGetPos(&mx, &my)
        if GetKeyState("LButton", "P") {
            if !dragging {
                dragging := true, sx := mx, sy := my
            }
            rx := Min(sx, mx), ry := Min(sy, my)
            rw := Abs(mx - sx), rh := Abs(my - sy)
            if (rw > 0 && rh > 0)
                sel.Show("x" rx " y" ry " w" rw " h" rh " NoActivate")
        } else if dragging {
            break
        }
        Sleep(10)
    }
    bg.Destroy(), sel.Destroy()
    x := Min(sx, mx), y := Min(sy, my)
    w := Abs(mx - sx), h := Abs(my - sy)
    return (w >= 5 && h >= 5)                        ; ignore accidental clicks
}

;-------------------------------------------------------------------------------
; Copy a screen rectangle to a PNG file (BitBlt + GDI+, as in Tests\Screenshots)
;-------------------------------------------------------------------------------
CaptureRegion(x, y, w, h, file) {
    hdcScreen := DllCall("GetDC", "Ptr", 0, "Ptr")
    hdcMem := DllCall("CreateCompatibleDC", "Ptr", hdcScreen, "Ptr")
    hbm := DllCall("CreateCompatibleBitmap", "Ptr", hdcScreen, "Int", w, "Int", h, "Ptr")
    old := DllCall("SelectObject", "Ptr", hdcMem, "Ptr", hbm, "Ptr")
    DllCall("BitBlt", "Ptr", hdcMem, "Int", 0, "Int", 0, "Int", w, "Int", h
        , "Ptr", hdcScreen, "Int", x, "Int", y, "UInt", 0x40CC0020)
    DllCall("SelectObject", "Ptr", hdcMem, "Ptr", old)
    DllCall("DeleteDC", "Ptr", hdcMem)
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)

    DllCall("LoadLibrary", "Str", "gdiplus")
    input := Buffer(24, 0)
    NumPut("UInt", 1, input)
    token := 0
    DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", input, "Ptr", 0)
    bitmap := 0
    DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hbm, "Ptr", 0, "Ptr*", &bitmap)
    clsid := Buffer(16)
    DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)
    try FileDelete(file)
    status := DllCall("gdiplus\GdipSaveImageToFile", "Ptr", bitmap, "WStr", file, "Ptr", clsid, "Ptr", 0)
    DllCall("gdiplus\GdipDisposeImage", "Ptr", bitmap)
    DllCall("gdiplus\GdiplusShutdown", "Ptr", token)
    DllCall("DeleteObject", "Ptr", hbm)
    return status = 0
}

;-------------------------------------------------------------------------------
; Run OCR and return its text. If the online engine fails (network down, bad key,
; quota exhausted, ...), automatically retry with the offline Windows engine.
; Returns a blank string only when both engines found no text.
;-------------------------------------------------------------------------------
Ocr(imagePath, enginePs, enginePsOnline, lang) {
    global engine, fallbackReason
    switch engine {
        case "windows":
            return RunWindowsOcr(imagePath, enginePs, lang)
        case "online":
            try {
                text := RunOnlineOcr(imagePath, enginePsOnline)
                return text
            } catch {
                ; Any online failure -> fall back to free offline Windows OCR.
                fallbackReason := "online unavailable"
            }
            return RunWindowsOcr(imagePath, enginePs, lang)
        default:
            Fail("Unknown ALTRUN_OCR_ENGINE: " engine)
    }
}

; Run the online engine. Throws on any failure so Ocr() can catch and fall back.
RunOnlineOcr(imagePath, enginePsOnline) {
    global tmpTxt
    WriteOnlineEngine(enginePsOnline)               ; self-contained: generate the signed PowerShell engine
    try FileDelete(tmpTxt)
    cmd := 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' enginePsOnline '" -Image "' imagePath '" -Out "' tmpTxt '"'
    RunWait(cmd, , "Hide")
    if !FileExist(tmpTxt)
        throw Error("Online OCR returned nothing")
    result := FileRead(tmpTxt, "UTF-8")
    if SubStr(result, 1, 7) = "ERROR: "
        throw Error(SubStr(result, 8))
    return result
}

; Run the offline Windows OCR engine. Kept separate so it can be reused as a fallback.
RunWindowsOcr(imagePath, enginePs, lang) {
    global tmpTxt
    WriteEngine(enginePs)                           ; self-contained: generate the PowerShell engine
    try FileDelete(tmpTxt)
    cmd := 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' enginePs '" -Image "' imagePath '" -Out "' tmpTxt '"'
    if (lang != "")
        cmd .= ' -Lang "' lang '"'
    RunWait(cmd, , "Hide")
    if !FileExist(tmpTxt)
        Fail("OCR engine returned nothing (is the OCR language pack installed?)")
    result := FileRead(tmpTxt, "UTF-8")
    if SubStr(result, 1, 7) = "ERROR: "
        Fail(SubStr(result, 8))
    return result
}

; Write the embedded Windows OCR (WinRT) engine to a temporary .ps1 file
WriteEngine(path) {
    src := "
(
param([string]$Image, [string]$Out, [string]$Lang)
$ErrorActionPreference = 'Stop'
# .NET error messages follow the OS UI culture (Chinese here). Force English so
# errors written to stdout survive ALTRun's GBK fallback without mojibake.
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'
try {
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $tick = [char]96
  $asyncName = 'IAsyncOperation' + $tick + '1'
  $asTask = @([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq $asyncName
  })[0]
  function Await($op, [type]$resultType) {
    $m = $asTask.MakeGenericMethod($resultType)
    $t = $m.Invoke($null, @($op))
    $t.Wait(-1) | Out-Null
    $t.Result
  }
  [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime] | Out-Null
  [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
  [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime] | Out-Null
  [Windows.Globalization.Language, Windows.Globalization, ContentType=WindowsRuntime] | Out-Null

  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Image)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $soft = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

  $ocr = $null
  if ($Lang) {
    try {
      $language = [Windows.Globalization.Language]::new($Lang)
      $ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
    } catch { }
  }
  if ($null -eq $ocr) {
    $ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
  }
  if ($null -eq $ocr) {
    throw 'No OCR language pack is installed for the current user languages.'
  }
  $result = Await ($ocr.RecognizeAsync($soft)) ([Windows.Media.Ocr.OcrResult])
  $text = ($result.Lines | ForEach-Object { $_.Text }) -join [Environment]::NewLine
  $cjk = '[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]'
  $text = [regex]::Replace($text, '(?<=' + $cjk + ')\s+(?=' + $cjk + ')', '')
  [System.IO.File]::WriteAllText($Out, $text, [System.Text.UTF8Encoding]::new($false))
} catch {
  [System.IO.File]::WriteAllText($Out, 'ERROR: ' + $_.Exception.Message, [System.Text.UTF8Encoding]::new($false))
  exit 1
}
)"
    f := FileOpen(path, "w", "UTF-8-RAW")             ; the embedded engine is pure ASCII
    f.Write(src)
    f.Close()
}

;-------------------------------------------------------------------------------
; Write the embedded online OCR (Volcengine OCRNormal) engine to a temp .ps1 file.
; It signs the request with Volcengine's V4 HMAC-SHA256 scheme and calls the API
; directly, so no Python / SDK is needed. Pure PowerShell, pure ASCII source.
;-------------------------------------------------------------------------------
WriteOnlineEngine(path) {
    src := "
(
param([string]$Image, [string]$Out)
$ErrorActionPreference = 'Stop'
# English .NET error messages avoid mojibake when ALTRun decodes stdout as GBK.
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'
$lf = [string][char]10
function Get-Hex([byte[]]$b) { ($b | ForEach-Object { $_.ToString('x2') }) -join '' }
function HmacBytes([byte[]]$key, [string]$data) {
  $h = [System.Security.Cryptography.HMACSHA256]::new()
  $h.Key = $key
  $h.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($data))
}
function Sha256Hex([string]$data) {
  $s = [System.Security.Cryptography.SHA256]::Create()
  Get-Hex $s.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($data))
}
try {
  $ak = $env:VOLC_AK
  $sk = $env:VOLC_SK
  if (-not $ak -or -not $sk) { throw 'VOLC_AK / VOLC_SK not set' }
  $service = 'cv'
  $region = 'cn-north-1'
  $hostName = 'visual.volcengineapi.com'
  $bytes = [System.IO.File]::ReadAllBytes($Image)
  $b64 = [Convert]::ToBase64String($bytes)
  $body = 'image_base64=' + [System.Net.WebUtility]::UrlEncode($b64)
  $bodyHash = Sha256Hex $body
  $utc = (Get-Date).ToUniversalTime()
  $date = $utc.ToString('yyyyMMdd')
  $xdate = $date + 'T' + $utc.ToString('HHmmss') + 'Z'
  $scope = $date + '/' + $region + '/' + $service + '/request'
  $signedHeaders = 'content-type;host;x-content-sha256;x-date'
  $signedStr = 'content-type:application/x-www-form-urlencoded' + $lf +
    'host:' + $hostName + $lf +
    'x-content-sha256:' + $bodyHash + $lf +
    'x-date:' + $xdate + $lf
  $canonical = 'POST' + $lf + '/' + $lf +
    'Action=OCRNormal&Version=2020-08-26' + $lf +
    $signedStr + $lf + $signedHeaders + $lf + $bodyHash
  $hashedCanon = Sha256Hex $canonical
  $signingStr = 'HMAC-SHA256' + $lf + $xdate + $lf + $scope + $lf + $hashedCanon
  $kDate = HmacBytes ([System.Text.Encoding]::UTF8.GetBytes($sk)) $date
  $kRegion = HmacBytes $kDate $region
  $kService = HmacBytes $kRegion $service
  $kSigning = HmacBytes $kService 'request'
  $signature = Get-Hex (HmacBytes $kSigning $signingStr)
  $auth = 'HMAC-SHA256 Credential=' + $ak + '/' + $scope + ', SignedHeaders=' + $signedHeaders + ', Signature=' + $signature
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $url = 'https://' + $hostName + '/?Action=OCRNormal&Version=2020-08-26'
  # HttpWebRequest instead of Invoke-RestMethod: on HTTP errors the latter
  # hides the response body, while this one exposes ResponseMetadata.Error.
  $req = [Net.HttpWebRequest]::Create($url)
  $req.Method = 'POST'
  $req.ContentType = 'application/x-www-form-urlencoded'
  $req.Timeout = 60000
  $req.Headers.Add('X-Date', $xdate)
  $req.Headers.Add('X-Content-Sha256', $bodyHash)
  $req.Headers.Add('Authorization', $auth)
  $bodyBytes = [Text.Encoding]::UTF8.GetBytes($body)
  $req.ContentLength = $bodyBytes.Length
  $rs = $req.GetRequestStream()
  $rs.Write($bodyBytes, 0, $bodyBytes.Length)
  $rs.Close()
  $respText = ''
  try {
    $resp = $req.GetResponse()
    $sr = [IO.StreamReader]::new($resp.GetResponseStream())
    $respText = $sr.ReadToEnd()
    $sr.Close()
    $resp.Close()
  } catch [Net.WebException] {
    $er = $_.Exception.Response
    if ($null -eq $er) { throw }
    $sr = [IO.StreamReader]::new($er.GetResponseStream())
    $respText = $sr.ReadToEnd()
    $sr.Close()
  }
  $respObj = $respText | ConvertFrom-Json
  $err = $respObj.ResponseMetadata.Error
  if ($null -ne $err) {
    throw ('code=' + $err.Code + ' message=' + $err.Message)
  }
  if ($null -ne $respObj.code -and $respObj.code -ne 10000) {
    throw ('code=' + $respObj.code + ' message=' + $respObj.message)
  }
  $lines = @($respObj.data.line_texts)
  $text = $lines -join $lf
  [System.IO.File]::WriteAllText($Out, $text, [System.Text.UTF8Encoding]::new($false))
} catch {
  [System.IO.File]::WriteAllText($Out, 'ERROR: ' + $_.Exception.Message, [System.Text.UTF8Encoding]::new($false))
  exit 1
}
)"
    f := FileOpen(path, "w", "UTF-8-RAW")             ; the embedded engine is pure ASCII
    f.Write(src)
    f.Close()
}

;-------------------------------------------------------------------------------
Fail(message) {
    FileAppend("OCR error: " message "`n", "**", "UTF-8-RAW")   ; first stderr line is shown on failure
    ExitApp(1)
}
