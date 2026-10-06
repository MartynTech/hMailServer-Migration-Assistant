//go:build windows

package main

import (
    "bufio"
    "bytes"
    "context"
    "encoding/json"
    "errors"
    "fmt"
    "html"
    "io"
    "net"
    "os"
    "os/exec"
    "path/filepath"
    "regexp"
    "strconv"
    "strings"
    "sync"
    "syscall"
    "time"
    "unsafe"
)

const (
    appName    = "hMailServer Migration Assistant"
    appVersion = "1.0"

    WM_CREATE  = 0x0001
    WM_DESTROY = 0x0002
    WM_SIZE    = 0x0005
    WM_COMMAND = 0x0111
    WM_CLOSE   = 0x0010
    WM_SETFONT = 0x0030
    WM_APP     = 0x8000

    WS_OVERLAPPEDWINDOW = 0x00CF0000
    WS_VISIBLE          = 0x10000000
    WS_CHILD            = 0x40000000
    WS_TABSTOP          = 0x00010000
    WS_BORDER           = 0x00800000
    WS_VSCROLL          = 0x00200000

    ES_AUTOHSCROLL = 0x0080
    ES_PASSWORD    = 0x0020
    ES_MULTILINE   = 0x0004
    ES_AUTOVSCROLL = 0x0040
    ES_READONLY    = 0x0800

    BS_PUSHBUTTON    = 0x00000000
    BS_AUTOCHECKBOX  = 0x00000003
    BS_GROUPBOX      = 0x00000007
    CBS_DROPDOWNLIST = 0x0003

    BM_GETCHECK = 0x00F0
    BM_SETCHECK = 0x00F1
    BST_CHECKED = 1

    CB_ADDSTRING    = 0x0143
    CB_RESETCONTENT = 0x014B
    CB_GETCURSEL    = 0x0147
    CB_SETCURSEL    = 0x014E

    PBM_SETRANGE32 = 0x0406
    PBM_SETPOS     = 0x0402

    MF_STRING    = 0x0000
    MF_POPUP     = 0x0010
    MF_SEPARATOR = 0x0800

    SW_SHOWNORMAL = 1
    SW_SHOW       = 5

    MB_OK              = 0x00000000
    MB_ICONINFORMATION = 0x00000040
    MB_ICONERROR       = 0x00000010
    MB_ICONWARNING     = 0x00000030
    MB_YESNO           = 0x00000004
    IDYES              = 6

    IDC_ARROW = 32512

    GWLP_USERDATA = -21

    ID_DISCOVER   = 1001
    ID_VALIDATE   = 1002
    ID_BENCHMARK  = 1003
    ID_INITIAL    = 1004
    ID_CUTOVER    = 1005
    ID_VERIFY     = 1006
    ID_ROLLBACK   = 1007
    ID_OPENLOGS   = 1008
    ID_REIMPORT   = 1009
    ID_DNS        = 1010
    ID_USE_SRC    = 1011
    ID_USE_DST    = 1012
    ID_BROWSE_SI  = 1020
    ID_BROWSE_SD  = 1021
    ID_BROWSE_SDA = 1022
    ID_BROWSE_SE  = 1023
    ID_BROWSE_DI  = 1030
    ID_BROWSE_DD  = 1031
    ID_BROWSE_DDA = 1032
    ID_BROWSE_DE  = 1033

    ID_MENU_EXIT  = 2001
    ID_MENU_LOGS  = 2002
    ID_MENU_EULA  = 2003
    ID_MENU_ABOUT = 2004

    APP_LOG      = WM_APP + 1
    APP_STATUS   = WM_APP + 2
    APP_DONE     = WM_APP + 3
    APP_DISCOVER = WM_APP + 4
    APP_PROGRESS = WM_APP + 5
)

var (
    user32   = syscall.NewLazyDLL("user32.dll")
    kernel32 = syscall.NewLazyDLL("kernel32.dll")
    shell32  = syscall.NewLazyDLL("shell32.dll")
    gdi32    = syscall.NewLazyDLL("gdi32.dll")
    comctl32 = syscall.NewLazyDLL("comctl32.dll")
    comdlg32 = syscall.NewLazyDLL("comdlg32.dll")
    mpr      = syscall.NewLazyDLL("mpr.dll")

    pRegisterClassExW = user32.NewProc("RegisterClassExW")
    pCreateWindowExW  = user32.NewProc("CreateWindowExW")
    pDefWindowProcW   = user32.NewProc("DefWindowProcW")
    pShowWindow       = user32.NewProc("ShowWindow")
    pUpdateWindow     = user32.NewProc("UpdateWindow")
    pGetMessageW      = user32.NewProc("GetMessageW")
    pTranslateMessage = user32.NewProc("TranslateMessage")
    pDispatchMessageW = user32.NewProc("DispatchMessageW")
    pPostQuitMessage  = user32.NewProc("PostQuitMessage")
    pSendMessageW     = user32.NewProc("SendMessageW")
    pPostMessageW     = user32.NewProc("PostMessageW")
    pSetWindowTextW   = user32.NewProc("SetWindowTextW")
    pGetWindowTextW   = user32.NewProc("GetWindowTextW")
    pGetWindowTextLen = user32.NewProc("GetWindowTextLengthW")
    pMessageBoxW      = user32.NewProc("MessageBoxW")
    pEnableWindow     = user32.NewProc("EnableWindow")
    pLoadCursorW      = user32.NewProc("LoadCursorW")
    pSetMenu          = user32.NewProc("SetMenu")
    pCreateMenu       = user32.NewProc("CreateMenu")
    pCreatePopupMenu  = user32.NewProc("CreatePopupMenu")
    pAppendMenuW      = user32.NewProc("AppendMenuW")
    pGetClientRect    = user32.NewProc("GetClientRect")
    pMoveWindow       = user32.NewProc("MoveWindow")

    pGetModuleHandleW = kernel32.NewProc("GetModuleHandleW")
    pCreateFontW      = gdi32.NewProc("CreateFontW")
    pShellExecuteW    = shell32.NewProc("ShellExecuteW")
    pIsUserAnAdmin    = shell32.NewProc("IsUserAnAdmin")
    pInitCommon       = comctl32.NewProc("InitCommonControls")

    pWNetAddConnection2W = mpr.NewProc("WNetAddConnection2W")
    pWNetCancelConnection2W = mpr.NewProc("WNetCancelConnection2W")
)

type WNDCLASSEX struct {
    CbSize        uint32
    Style         uint32
    LpfnWndProc   uintptr
    CbClsExtra    int32
    CbWndExtra    int32
    HInstance     syscall.Handle
    HIcon         syscall.Handle
    HCursor       syscall.Handle
    HbrBackground syscall.Handle
    LpszMenuName  *uint16
    LpszClassName *uint16
    HIconSm       syscall.Handle
}

type MSG struct {
    Hwnd    syscall.Handle
    Message uint32
    WParam  uintptr
    LParam  uintptr
    Time    uint32
    Pt      struct{ X, Y int32 }
}

type RECT struct{ Left, Top, Right, Bottom int32 }

type NETRESOURCE struct {
    Scope       uint32
    Type        uint32
    DisplayType uint32
    Usage       uint32
    LocalName   *uint16
    RemoteName  *uint16
    Comment     *uint16
    Provider    *uint16
}

type Candidate struct {
    IniWindows      string
    IniAccess       string
    DatabaseWindows string
    DatabaseAccess  string
    DataWindows     string
    DataAccess      string
    EventsWindows   string
    EventsAccess    string
    ProgramFolder   string
    DatabaseType    string
    DatabaseSize    int64
    Version         string
    SevenZip        string
    Active          bool
}

type DiscoveryResult struct {
    Source      []Candidate
    Destination []Candidate
    Err         error
}

type StoreInfo struct {
    Files int64
    Bytes int64
}

type Snapshot struct {
    Domains           int               `json:"Domains"`
    Accounts          int               `json:"Accounts"`
    Aliases           int               `json:"Aliases"`
    DistributionLists int               `json:"DistributionLists"`
    Messages          int               `json:"Messages"`
    IMAPFolders       int               `json:"IMAPFolders"`
    Routes            int               `json:"Routes"`
    IPRanges          int               `json:"IPRanges"`
    Rules             int               `json:"Rules"`
    HostName          string            `json:"HostName"`
    DomainsList       []string          `json:"DomainsList"`
    Mailboxes         map[string]int    `json:"Mailboxes"`
    Certificates      []ExternalFile    `json:"Certificates"`
    DKIMKeys          []ExternalFile    `json:"DKIMKeys"`
}

type ExternalFile struct {
    Kind string `json:"Kind"`
    Name string `json:"Name"`
    Path string `json:"Path"`
}

type PreflightResult struct {
    SourceStore       StoreInfo
    DestinationStore  StoreInfo
    SourceVersion     string
    DestinationVersion string
    SourceDBSize      int64
    DestinationDBSize int64
    DestinationFree   uint64
    SourceSevenZip    string
    DestinationSevenZip string
    ServiceAccount    string
    SourceSnapshot    *Snapshot
    ExternalFiles     []ExternalFile
    Warnings          []string
}

type ProgressInfo struct {
    Percent int
    Text    string
}

type App struct {
    hwnd syscall.Handle
    font syscall.Handle

    srcHost, srcShare, srcUser, srcPass, publicIP syscall.Handle
    stopSource, allowMismatch, ask7zip syscall.Handle
    srcCombo, dstCombo syscall.Handle
    srcIni, srcDB, srcData, srcEvents syscall.Handle
    dstIni, dstDB, dstData, dstEvents syscall.Handle
    srcAdminPass, dstAdminPass, archiveStage syscall.Handle
    transferMethod, threads syscall.Handle
    progress syscall.Handle
    status, speed, log syscall.Handle

    sourceCandidates []Candidate
    destCandidates   []Candidate
    sourceSelected   *Candidate
    destSelected     *Candidate
    preflight        *PreflightResult
    lastBackup       string
    lastReport       string
    logPath          string

    busy bool
    mu sync.Mutex
}

var app *App

func u16(s string) *uint16 { p, _ := syscall.UTF16PtrFromString(s); return p }

func loword(v uintptr) uint16 { return uint16(v & 0xffff) }

func createFont(size int) syscall.Handle {
    r, _, _ := pCreateFontW.Call(uintptr(int32(-size)), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, uintptr(unsafe.Pointer(u16("Segoe UI"))))
    return syscall.Handle(r)
}

func setFont(h, f syscall.Handle) { pSendMessageW.Call(uintptr(h), WM_SETFONT, uintptr(f), 1) }

func createControl(class, text string, style uint32, x,y,w,h int, parent syscall.Handle, id int) syscall.Handle {
    r, _, _ := pCreateWindowExW.Call(0, uintptr(unsafe.Pointer(u16(class))), uintptr(unsafe.Pointer(u16(text))), uintptr(style|WS_CHILD|WS_VISIBLE), uintptr(x), uintptr(y), uintptr(w), uintptr(h), uintptr(parent), uintptr(id), 0, 0)
    hwnd := syscall.Handle(r)
    if app != nil && app.font != 0 { setFont(hwnd, app.font) }
    return hwnd
}

func addLabel(text string, x,y,w,h int, parent syscall.Handle) syscall.Handle {
    return createControl("STATIC", text, 0, x,y,w,h,parent,0)
}
func addEdit(text string, x,y,w,h int, parent syscall.Handle, password bool) syscall.Handle {
    st := uint32(WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL)
    if password { st |= ES_PASSWORD }
    return createControl("EDIT", text, st, x,y,w,h,parent,0)
}
func addButton(text string, x,y,w,h,id int,parent syscall.Handle) syscall.Handle {
    return createControl("BUTTON", text, WS_TABSTOP|BS_PUSHBUTTON, x,y,w,h,parent,id)
}
func addCheck(text string, x,y,w,h int, id int,parent syscall.Handle, checked bool) syscall.Handle {
    hnd := createControl("BUTTON", text, WS_TABSTOP|BS_AUTOCHECKBOX, x,y,w,h,parent,id)
    if checked { pSendMessageW.Call(uintptr(hnd),BM_SETCHECK,BST_CHECKED,0) }
    return hnd
}
func addGroup(text string,x,y,w,h int,parent syscall.Handle) syscall.Handle {
    return createControl("BUTTON", text, BS_GROUPBOX, x,y,w,h,parent,0)
}
func addCombo(x,y,w,h,id int,parent syscall.Handle) syscall.Handle {
    return createControl("COMBOBOX", "", WS_TABSTOP|CBS_DROPDOWNLIST, x,y,w,h,parent,id)
}
func comboAdd(h syscall.Handle, s string) { pSendMessageW.Call(uintptr(h),CB_ADDSTRING,0,uintptr(unsafe.Pointer(u16(s)))) }
func comboReset(h syscall.Handle) { pSendMessageW.Call(uintptr(h),CB_RESETCONTENT,0,0) }
func comboSelect(h syscall.Handle, idx int) { pSendMessageW.Call(uintptr(h),CB_SETCURSEL,uintptr(idx),0) }
func comboSelected(h syscall.Handle) int { r,_,_:=pSendMessageW.Call(uintptr(h),CB_GETCURSEL,0,0); return int(int32(r)) }

func getText(h syscall.Handle) string {
    n,_,_:=pGetWindowTextLen.Call(uintptr(h)); if n==0 { return "" }
    b:=make([]uint16,n+1); pGetWindowTextW.Call(uintptr(h),uintptr(unsafe.Pointer(&b[0])),n+1)
    return syscall.UTF16ToString(b)
}
func setText(h syscall.Handle,s string){ pSetWindowTextW.Call(uintptr(h),uintptr(unsafe.Pointer(u16(s)))) }
func checked(h syscall.Handle) bool { r,_,_:=pSendMessageW.Call(uintptr(h),BM_GETCHECK,0,0); return r==BST_CHECKED }

func messageBox(text,title string,flags uintptr) int { r,_,_:=pMessageBoxW.Call(uintptr(app.hwnd),uintptr(unsafe.Pointer(u16(text))),uintptr(unsafe.Pointer(u16(title))),flags); return int(r) }

func postString(msg uint32,s string){ p:=new(string); *p=s; pPostMessageW.Call(uintptr(app.hwnd),uintptr(msg),0,uintptr(unsafe.Pointer(p))) }
func postProgress(percent int,text string){ pi:=&ProgressInfo{Percent:percent,Text:text}; pPostMessageW.Call(uintptr(app.hwnd),APP_PROGRESS,0,uintptr(unsafe.Pointer(pi))) }
func logf(level,format string,args ...any){ line:=fmt.Sprintf("[%s] [%s] %s",time.Now().Format("2006-01-02 15:04:05"),strings.ToUpper(level),fmt.Sprintf(format,args...)); appendLogFile(line); postString(APP_LOG,line) }

func ensureLog() string {
    if app.logPath!="" { return app.logPath }
    root:=`C:\ProgramData\hMailServer\MigrationAssistant\Logs`; _=os.MkdirAll(root,0755)
    app.logPath=filepath.Join(root,"Migration-"+time.Now().Format("20060102-150405")+".log")
    return app.logPath
}
func appendLogFile(line string){ path:=ensureLog(); f,err:=os.OpenFile(path,os.O_CREATE|os.O_APPEND|os.O_WRONLY,0644); if err==nil { fmt.Fprintln(f,line); _=f.Close() } }

func setBusy(v bool,status string){ app.mu.Lock(); app.busy=v; app.mu.Unlock(); postString(APP_STATUS,status) }

func formatBytes(n int64) string {
    if n<0 { return "--" }
    f:=float64(n)
    switch { case f>=1<<40:return fmt.Sprintf("%.2f TB",f/(1<<40)); case f>=1<<30:return fmt.Sprintf("%.2f GB",f/(1<<30)); case f>=1<<20:return fmt.Sprintf("%.2f MB",f/(1<<20)); case f>=1<<10:return fmt.Sprintf("%.2f KB",f/(1<<10)); default:return fmt.Sprintf("%d B",n) }
}
func formatDuration(d time.Duration) string { if d<0{return "--"}; d=d.Round(time.Second); h:=int(d.Hours()); m:=int(d.Minutes())%60; s:=int(d.Seconds())%60; if h>0{return fmt.Sprintf("%dh %dm %ds",h,m,s)}; if m>0{return fmt.Sprintf("%dm %ds",m,s)}; return fmt.Sprintf("%ds",s) }

func isAdmin() bool { r,_,_:=pIsUserAnAdmin.Call(); return r!=0 }
func elevateIfNeeded() bool {
    if isAdmin(){return false}
    exe,_:=os.Executable(); verb:=u16("runas"); file:=u16(exe)
    r,_,_:=pShellExecuteW.Call(0,uintptr(unsafe.Pointer(verb)),uintptr(unsafe.Pointer(file)),0,0,SW_SHOWNORMAL)
    return r>32
}

func connectSource() error {
    host:=strings.TrimSpace(getText(app.srcHost)); share:=strings.TrimSpace(getText(app.srcShare)); user:=strings.TrimSpace(getText(app.srcUser)); pass:=getText(app.srcPass)
    if host=="" { return errors.New("enter the source host / IP address") }
    if share=="" { share="C$"; setText(app.srcShare,share) }
    remote:=`\\`+host+`\`+share
    nr:=NETRESOURCE{Type:1,RemoteName:u16(remote)}
    var up,pp *uint16
    if user!="" { up=u16(user) }
    if pass!="" { pp=u16(pass) }
    r,_,_:=pWNetAddConnection2W.Call(uintptr(unsafe.Pointer(&nr)),uintptr(unsafe.Pointer(pp)),uintptr(unsafe.Pointer(up)),0)
    if r!=0 && r!=1219 && r!=85 { return fmt.Errorf("unable to connect to %s (Windows error %d)",remote,r) }
    return nil
}

func disconnectSource(){ host:=strings.TrimSpace(getText(app.srcHost)); share:=strings.TrimSpace(getText(app.srcShare)); if host==""||share==""{return}; remote:=`\\`+host+`\`+share; pWNetCancelConnection2W.Call(uintptr(unsafe.Pointer(u16(remote))),0,0) }

func parseIni(path string) (map[string]map[string]string,error) {
    b,err:=os.ReadFile(path); if err!=nil{return nil,err}
    res:=map[string]map[string]string{}; sec:=""
    sc:=bufio.NewScanner(bytes.NewReader(b))
    for sc.Scan(){ line:=strings.TrimSpace(strings.TrimPrefix(sc.Text(),"\ufeff")); if line==""||strings.HasPrefix(line,";")||strings.HasPrefix(line,"#"){continue}; if strings.HasPrefix(line,"[")&&strings.HasSuffix(line,"]"){sec=strings.TrimSpace(line[1:len(line)-1]); if _,ok:=res[sec];!ok{res[sec]=map[string]string{}}; continue}; if i:=strings.Index(line,"=");i>=0&&sec!=""{res[sec][strings.TrimSpace(line[:i])]=strings.TrimSpace(line[i+1:])} }
    return res,sc.Err()
}
func iniVal(m map[string]map[string]string,sec,key string) string { for s,v:=range m{if strings.EqualFold(s,sec){for k,x:=range v{if strings.EqualFold(k,key){return x}}}}; return "" }

func windowsToUNC(host, share, p string) string {
    p=strings.TrimSpace(p); if p==""{return ""}
    if strings.HasPrefix(p,`\\`){return p}
    if len(p)>=3&&p[1]==':'&&(p[2]=='\\'||p[2]=='/'){
        drive:=strings.ToUpper(string(p[0])); rel:=strings.TrimLeft(strings.ReplaceAll(p[2:],"/",`\`),`\`)
        // Administrative shares are the safest translation for hMailServer paths on arbitrary drives.
        return `\\`+host+`\`+drive+`$\`+rel
    }
    return `\\`+host+`\`+strings.Trim(share,`\`)+`\`+strings.TrimLeft(p,`\`)
}

func candidateFromIni(iniAccess,iniWindows string,remote bool) (Candidate,error) {
    m,err:=parseIni(iniAccess); if err!=nil{return Candidate{},err}
    dbFolder:=iniVal(m,"Directories","DatabaseFolder"); dataFolder:=iniVal(m,"Directories","DataFolder"); eventFolder:=iniVal(m,"Directories","EventFolder"); program:=iniVal(m,"Directories","ProgramFolder"); dbName:=iniVal(m,"Database","Database"); if dbName==""{dbName="hMailServer"}; if !strings.HasSuffix(strings.ToLower(dbName),".sdf"){dbName += ".sdf"}
    c:=Candidate{IniWindows:iniWindows,IniAccess:iniAccess,DatabaseWindows:filepath.Join(dbFolder,dbName),DataWindows:dataFolder,EventsWindows:eventFolder,ProgramFolder:program,DatabaseType:iniVal(m,"Database","Type")}
    if remote { host:=strings.TrimSpace(getText(app.srcHost)); share:=strings.TrimSpace(getText(app.srcShare)); c.DatabaseAccess=windowsToUNC(host,share,c.DatabaseWindows); c.DataAccess=windowsToUNC(host,share,c.DataWindows); c.EventsAccess=windowsToUNC(host,share,c.EventsWindows) } else { c.DatabaseAccess=c.DatabaseWindows; c.DataAccess=c.DataWindows; c.EventsAccess=c.EventsWindows }
    if fi,e:=os.Stat(c.DatabaseAccess);e==nil{c.DatabaseSize=fi.Size()}
    c.Version=getHMailVersion(program,remote)
    c.SevenZip=find7zip(remote)
    return c,nil
}

func discoverCandidates(remote bool) ([]Candidate,error) {
    var paths []struct{access,win string}
    if remote {
        if err:=connectSource();err!=nil{return nil,err}
        host:=strings.TrimSpace(getText(app.srcHost)); share:=strings.TrimSpace(getText(app.srcShare))
        wins:=[]string{`C:\ProgramData\hMailServer\hMailServer.ini`,`C:\Program Files\hMailServer\Bin\hMailServer.ini`,`C:\Program Files (x86)\hMailServer\Bin\hMailServer.ini`}
        for _,w:=range wins{paths=append(paths,struct{access,win string}{windowsToUNC(host,share,w),w})}
    } else {
        wins:=[]string{`C:\ProgramData\hMailServer\hMailServer.ini`,`C:\Program Files\hMailServer\Bin\hMailServer.ini`,`C:\Program Files (x86)\hMailServer\Bin\hMailServer.ini`}
        for _,w:=range wins{paths=append(paths,struct{access,win string}{w,w})}
    }
    seen:=map[string]bool{}; var out []Candidate
    for _,p:=range paths{ if seen[strings.ToLower(p.access)]{continue}; seen[strings.ToLower(p.access)]=true; if _,err:=os.Stat(p.access);err!=nil{continue}; c,err:=candidateFromIni(p.access,p.win,remote);if err==nil{out=append(out,c)} }
    if len(out)==0{return nil,errors.New("no hMailServer.ini installation candidates were found")}
    return out,nil
}

func find7zip(remote bool) string {
    candidates:=[]string{`C:\Program Files\7-Zip\7z.exe`,`C:\Program Files (x86)\7-Zip\7z.exe`}
    if remote { host:=strings.TrimSpace(getText(app.srcHost)); share:=strings.TrimSpace(getText(app.srcShare)); for _,w:=range candidates{ if _,e:=os.Stat(windowsToUNC(host,share,w));e==nil{return w} } } else { for _,p:=range candidates{if _,e:=os.Stat(p);e==nil{return p}} }
    return ""
}

func getHMailVersion(program string, remote bool) string {
    if program==""{return ""}; exe:=filepath.Join(program,"Bin","hMailServer.exe"); if remote{exe=windowsToUNC(strings.TrimSpace(getText(app.srcHost)),strings.TrimSpace(getText(app.srcShare)),exe)}
    ps:=fmt.Sprintf(`$v=(Get-Item -LiteralPath '%s' -ErrorAction Stop).VersionInfo.FileVersion; Write-Output $v`,strings.ReplaceAll(exe,"'","''"))
    out,err:=powershell(ps,nil); if err!=nil{return ""}; return strings.TrimSpace(out)
}

func powershell(script string, env map[string]string) (string,error) {
    cmd:=exec.Command("powershell.exe","-NoProfile","-NonInteractive","-ExecutionPolicy","Bypass","-Command",script)
    cmd.Env=os.Environ(); for k,v:=range env{cmd.Env=append(cmd.Env,k+"="+v)}
    b,err:=cmd.CombinedOutput(); if err!=nil{return string(b),fmt.Errorf("%w: %s",err,strings.TrimSpace(string(b)))}; return string(b),nil
}

func serviceRunning(remote bool) bool {
    args:=[]string{"query","hMailServer"}; if remote{host:=strings.TrimSpace(getText(app.srcHost)); args=append([]string{`\\`+host},args...)}
    b,_:=exec.Command("sc.exe",args...).CombinedOutput(); return strings.Contains(strings.ToUpper(string(b)),"RUNNING")
}
func stopService(remote bool) error { args:=[]string{"stop","hMailServer"}; if remote{args=append([]string{`\\`+strings.TrimSpace(getText(app.srcHost))},args...)}; b,err:=exec.Command("sc.exe",args...).CombinedOutput(); if err!=nil&&!strings.Contains(strings.ToLower(string(b)),"service has not been started"){return fmt.Errorf("stop hMailServer: %s",strings.TrimSpace(string(b)))}; deadline:=time.Now().Add(45*time.Second); for time.Now().Before(deadline){if !serviceRunning(remote){return nil};time.Sleep(time.Second)}; return errors.New("timed out waiting for hMailServer to stop") }
func startService(remote bool) error { args:=[]string{"start","hMailServer"}; if remote{args=append([]string{`\\`+strings.TrimSpace(getText(app.srcHost))},args...)}; b,err:=exec.Command("sc.exe",args...).CombinedOutput(); if err!=nil&&!strings.Contains(strings.ToLower(string(b)),"already been started"){return fmt.Errorf("start hMailServer: %s",strings.TrimSpace(string(b)))}; return nil }

func getServiceAccount() string { b,_:=exec.Command("sc.exe","qc","hMailServer").CombinedOutput(); for _,ln:=range strings.Split(string(b),"\n"){if strings.Contains(ln,"SERVICE_START_NAME"){if i:=strings.Index(ln,":");i>=0{return strings.TrimSpace(ln[i+1:])}}};return "" }

func getDiskFree(path string) uint64 {
    // PowerShell is used here because it handles arbitrary local volumes cleanly and is not in the hot data-copy path.
    vol:=filepath.VolumeName(path); if vol==""{return 0}; ps:=fmt.Sprintf(`$d=Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='%s'"; if($d){$d.FreeSpace}`,strings.ReplaceAll(vol,"'","''")); out,_:=powershell(ps,nil); n,_:=strconv.ParseUint(strings.TrimSpace(out),10,64); return n
}
func writeTest(dir string) error { if dir==""{return errors.New("destination Data folder is blank")}; if err:=os.MkdirAll(dir,0755);err!=nil{return err}; p:=filepath.Join(dir,fmt.Sprintf(".hma-write-test-%d.tmp",time.Now().UnixNano())); if err:=os.WriteFile(p,[]byte("test"),0600);err!=nil{return err}; return os.Remove(p) }

var reFilesLine=regexp.MustCompile(`(?mi)^\s*Files\s*:\s*([0-9.,]+)`) 
var reBytesLine=regexp.MustCompile(`(?mi)^\s*Bytes\s*:\s*([0-9.,]+)`)

func robocopyEnumerate(data string) (StoreInfo,error) {
    tmp,err:=os.MkdirTemp("","hma-enum-"); if err!=nil{return StoreInfo{},err}; defer os.RemoveAll(tmp)
    args:=[]string{data,tmp,"*.eml","/L","/E","/BYTES","/NJH","/NDL","/NFL","/NP","/R:0","/W:0","/XJ"}
    ctx,cancel:=context.WithTimeout(context.Background(),20*time.Minute); defer cancel()
    cmd:=exec.CommandContext(ctx,"robocopy.exe",args...); b,err:=cmd.CombinedOutput(); if ctx.Err()!=nil{return StoreInfo{},errors.New("Robocopy /L mail enumeration timed out")}
    // Robocopy codes 0-7 are success states.
    if err!=nil { if ee,ok:=err.(*exec.ExitError);!ok||ee.ExitCode()>7{return StoreInfo{},fmt.Errorf("Robocopy /L failed: %s",strings.TrimSpace(string(b)))} }
    parseNum:=func(s string)(int64,error){s=strings.ReplaceAll(strings.ReplaceAll(strings.TrimSpace(s),",",""),".",""); return strconv.ParseInt(s,10,64)}
    fm:=reFilesLine.FindStringSubmatch(string(b)); bm:=reBytesLine.FindStringSubmatch(string(b)); if len(fm)<2||len(bm)<2{return StoreInfo{},errors.New("unable to parse Robocopy mail-store summary")}
    f,e1:=parseNum(fm[1]); by,e2:=parseNum(bm[1]); if e1!=nil||e2!=nil{return StoreInfo{},errors.New("invalid Robocopy mail-store summary")}; return StoreInfo{Files:f,Bytes:by},nil
}

func buildPreflight() (*PreflightResult,error) {
    src,dst,err:=currentCandidates(); if err!=nil{return nil,err}
    if !strings.EqualFold(src.DatabaseType,"MSSQLCE") { return nil,fmt.Errorf("source database type is %q, expected MSSQLCE",src.DatabaseType) }
    if !strings.EqualFold(dst.DatabaseType,"MSSQLCE") { return nil,fmt.Errorf("destination database type is %q, expected MSSQLCE",dst.DatabaseType) }
    for label,path:=range map[string]string{"Source INI":src.IniAccess,"Source database":src.DatabaseAccess,"Source Data":src.DataAccess,"Destination INI":dst.IniAccess,"Destination database":dst.DatabaseAccess,"Destination Data":dst.DataAccess}{if _,e:=os.Stat(path);e!=nil{return nil,fmt.Errorf("%s failed: %v",label,e)}}
    if err:=writeTest(dst.DataAccess);err!=nil{return nil,fmt.Errorf("destination Data folder write test failed: %v",err)}
    r:=&PreflightResult{SourceVersion:src.Version,DestinationVersion:dst.Version,SourceDBSize:src.DatabaseSize,DestinationDBSize:dst.DatabaseSize,DestinationFree:getDiskFree(dst.DataWindows),SourceSevenZip:src.SevenZip,DestinationSevenZip:dst.SevenZip,ServiceAccount:getServiceAccount()}
    postString(APP_STATUS,"Enumerating source mail store with Robocopy /L...")
    s,e:=robocopyEnumerate(src.DataAccess); if e!=nil{return nil,fmt.Errorf("source mail discovery failed: %v",e)}; r.SourceStore=s
    postString(APP_STATUS,"Enumerating destination mail store with Robocopy /L...")
    d,e:=robocopyEnumerate(dst.DataAccess); if e!=nil{return nil,fmt.Errorf("destination mail discovery failed: %v",e)}; r.DestinationStore=d
    if src.Version!=""&&dst.Version!=""&&src.Version!=dst.Version&&!checked(app.allowMismatch){return nil,fmt.Errorf("hMailServer version mismatch (source %q, destination %q)",src.Version,dst.Version)}
    if r.DestinationFree>0&&uint64(r.SourceStore.Bytes)>r.DestinationFree{r.Warnings=append(r.Warnings,"Destination free space is lower than the physical source mail-store size.")}
    if pw:=getText(app.srcAdminPass);pw!=""{if snap,e:=runCOMSnapshot(true,pw);e==nil{r.SourceSnapshot=&snap;r.ExternalFiles=append(r.ExternalFiles,snap.Certificates...);r.ExternalFiles=append(r.ExternalFiles,snap.DKIMKeys...)}else{r.Warnings=append(r.Warnings,"Source COM comparison unavailable: "+e.Error())}}
    return r,nil
}

func currentCandidates() (Candidate,Candidate,error) {
    if app.sourceSelected==nil||app.destSelected==nil{return Candidate{},Candidate{},errors.New("discover and select both source and destination installations first")}
    s:=*app.sourceSelected; d:=*app.destSelected
    s.IniAccess=strings.TrimSpace(getText(app.srcIni)); s.DatabaseAccess=strings.TrimSpace(getText(app.srcDB)); s.DataAccess=strings.TrimSpace(getText(app.srcData)); s.EventsAccess=strings.TrimSpace(getText(app.srcEvents))
    d.IniAccess=strings.TrimSpace(getText(app.dstIni)); d.DatabaseAccess=strings.TrimSpace(getText(app.dstDB)); d.DataAccess=strings.TrimSpace(getText(app.dstData)); d.EventsAccess=strings.TrimSpace(getText(app.dstEvents))
    if s.IniAccess==""||s.DatabaseAccess==""||s.DataAccess==""||d.IniAccess==""||d.DatabaseAccess==""||d.DataAccess==""{return s,d,errors.New("INI, database and Data paths are required for both source and destination")}
    return s,d,nil
}

func runCOMSnapshot(remote bool,password string)(Snapshot,error){
    script:=`$ErrorActionPreference='Stop'
$app=New-Object -ComObject hMailServer.Application
if(-not $app.Authenticate('Administrator',$env:HMA_HMAILPASS)){ throw 'hMailServer Administrator authentication failed' }
$r=[ordered]@{Domains=0;Accounts=0;Aliases=0;DistributionLists=0;Messages=0;IMAPFolders=0;Routes=0;IPRanges=0;Rules=0;Mailboxes=@{};Certificates=@();DKIMKeys=@();HostName='';DomainsList=@()}
try{$r.HostName=[string]$app.Settings.HostName}catch{}
try{$r.Routes=[int]$app.Settings.Routes.Count}catch{}
try{$r.IPRanges=[int]$app.Settings.SecurityRanges.Count}catch{}
try{$r.Rules=[int]$app.Rules.Count}catch{}
try{ $cs=$app.Settings.SSLCertificates; for($i=0;$i -lt $cs.Count;$i++){ $c=$cs.Item($i); $r.Certificates += [pscustomobject]@{Kind='TLS certificate';Name=[string]$c.Name;Path=[string]$c.CertificateFile}; $r.Certificates += [pscustomobject]@{Kind='TLS private key';Name=[string]$c.Name;Path=[string]$c.PrivateKeyFile} } }catch{}
$d=$app.Domains;$r.Domains=[int]$d.Count
for($i=0;$i -lt $d.Count;$i++){ $dom=$d.Item($i); $r.DomainsList += [string]$dom.Name; try{if($dom.DKIMSignEnabled -and $dom.DKIMPrivateKeyFile){$r.DKIMKeys += [pscustomobject]@{Kind='DKIM private key';Name=[string]$dom.Name;Path=[string]$dom.DKIMPrivateKeyFile}}}catch{}; try{$r.Aliases += [int]$dom.Aliases.Count}catch{}; try{$r.DistributionLists += [int]$dom.DistributionLists.Count}catch{}; $ac=$dom.Accounts; $r.Accounts += [int]$ac.Count; for($j=0;$j -lt $ac.Count;$j++){ $a=$ac.Item($j); $mc=0; try{$mc=[int]$a.Messages.Count}catch{}; $r.Messages += $mc; $r.Mailboxes[[string]$a.Address]=$mc; try{$r.IMAPFolders += [int]$a.IMAPFolders.Count}catch{} } }
$r | ConvertTo-Json -Compress -Depth 8`
    env:=map[string]string{"HMA_HMAILPASS":password}
    if !remote { out,err:=powershell(script,env); if err!=nil{return Snapshot{},err}; var s Snapshot; err=json.Unmarshal([]byte(strings.TrimSpace(out)),&s); return s,err }
    host:=strings.TrimSpace(getText(app.srcHost)); user:=strings.TrimSpace(getText(app.srcUser)); pass:=getText(app.srcPass)
    wrapper:=`$ErrorActionPreference='Stop';$sb={param($hma) $env:HMA_HMAILPASS=$hma;`+script+`};`+remoteInvokePrefix(user)+`Invoke-Command -ComputerName $env:HMA_HOST `+remoteCredentialArg(user)+` -ScriptBlock $sb -ArgumentList $env:HMA_HMAILPASS`
    env["HMA_HOST"]=host; env["HMA_USER"]=user; env["HMA_PASS"]=pass
    out,err:=powershell(wrapper,env); if err!=nil{return Snapshot{},err}; var s Snapshot; err=json.Unmarshal([]byte(strings.TrimSpace(out)),&s); return s,err
}
func remoteInvokePrefix(user string)string{if user==""{return ""};return `$s=ConvertTo-SecureString $env:HMA_PASS -AsPlainText -Force;$c=New-Object PSCredential($env:HMA_USER,$s);`}
func remoteCredentialArg(user string)string{if user==""{return ""};return `-Credential $c`}

func parseRobocopyNeeded(source,dest string)(StoreInfo,error){tmp,err:=os.MkdirTemp("","hma-plan-");if err!=nil{return StoreInfo{},err};defer os.RemoveAll(tmp);_ = tmp; args:=[]string{source,dest,"/L","/E","/BYTES","/NJH","/NDL","/NFL","/NP","/R:0","/W:0","/XJ"};b,e:=exec.Command("robocopy.exe",args...).CombinedOutput();if e!=nil{if ee,ok:=e.(*exec.ExitError);!ok||ee.ExitCode()>7{return StoreInfo{},fmt.Errorf("Robocopy plan failed: %s",string(b))}};fm:=reFilesLine.FindStringSubmatch(string(b));bm:=reBytesLine.FindStringSubmatch(string(b));if len(fm)<2||len(bm)<2{return StoreInfo{},errors.New("unable to parse Robocopy plan")};clean:=func(s string)int64{n,_:=strconv.ParseInt(strings.NewReplacer(",","",".","").Replace(strings.TrimSpace(s)),10,64);return n};return StoreInfo{Files:clean(fm[1]),Bytes:clean(bm[1])},nil}

func robocopyData(source,dest,label string,threads int) error {
    if threads<1{threads=16}; plan,_:=parseRobocopyNeeded(source,dest); logf("INFO","%s - planned transfer: %d files, %s",label,plan.Files,formatBytes(plan.Bytes)); postProgress(0,label+" - starting")
    args:=[]string{source,dest,"/E","/COPY:DAT","/DCOPY:DAT","/Z","/R:2","/W:2",fmt.Sprintf("/MT:%d",threads),"/XJ","/BYTES","/NP"}
    cmd:=exec.Command("robocopy.exe",args...); stdout,err:=cmd.StdoutPipe();if err!=nil{return err};cmd.Stderr=cmd.Stdout; if err:=cmd.Start();err!=nil{return err}
    start:=time.Now(); sc:=bufio.NewScanner(stdout); for sc.Scan(){line:=sc.Text(); if strings.TrimSpace(line)!=""{logf("ROBOCOPY","%s",line)}; elapsed:=time.Since(start); if plan.Bytes>0{postProgress(0,fmt.Sprintf("%s - %s planned | elapsed %s",label,formatBytes(plan.Bytes),formatDuration(elapsed)))} }
    err=cmd.Wait(); code:=0;if err!=nil{if ee,ok:=err.(*exec.ExitError);ok{code=ee.ExitCode()}else{return err}};if code>7{return fmt.Errorf("Robocopy exit code %d",code)};postProgress(100,label+" completed");return nil
}

func initialSync() error {
    src,dst,err:=currentCandidates();if err!=nil{return err}; threads,_:=strconv.Atoi(strings.TrimSpace(getText(app.threads)));if threads<1{threads=32}
    method:=comboSelected(app.transferMethod); if method<0{method=0}
    useArchive:=method==2
    if method==0 && checked(app.ask7zip) && app.preflight!=nil && app.preflight.SourceStore.Files>=50000 && src.SevenZip!="" && dst.SevenZip!="" {
        if messageBox(fmt.Sprintf("The source contains %d .eml files.\n\nWould you like the assistant to create a fast-compression 7-Zip archive on the SOURCE host, transfer the single archive, then extract it locally?\n\nA final incremental Robocopy will still run afterwards to catch messages created during archiving.",app.preflight.SourceStore.Files),"7-Zip initial seed",MB_YESNO|MB_ICONINFORMATION)==IDYES{useArchive=true}
    }
    if useArchive {
        if src.SevenZip==""||dst.SevenZip==""{return errors.New("7-Zip archive mode requires 7-Zip installed on both source and destination")}
        if err:=remoteWinRMTest();err!=nil{return fmt.Errorf("7-Zip source-side archive requires PowerShell remoting: %v",err)}
        archive,err:=createRemoteArchive(src);if err!=nil{return err}; local:=filepath.Join(os.TempDir(),filepath.Base(archive)); if err:=copySingleWithProgress(windowsToUNC(strings.TrimSpace(getText(app.srcHost)),strings.TrimSpace(getText(app.srcShare)),archive),local,"7-Zip archive transfer");err!=nil{return err}; if err:=extractArchive(dst.SevenZip,local,dst.DataAccess);err!=nil{return err}; _=os.Remove(local); logf("INFO","Archive extracted. Running incremental mail sync to catch changes...")
    }
    logf("INFO","Robocopy source: %s",src.DataAccess);logf("INFO","Robocopy destination: %s",dst.DataAccess)
    return robocopyData(src.DataAccess,dst.DataAccess,"Initial mail sync",threads)
}

func remoteWinRMTest() error { host:=strings.TrimSpace(getText(app.srcHost));user:=strings.TrimSpace(getText(app.srcUser));pass:=getText(app.srcPass);script:=`$ErrorActionPreference='Stop';`+remoteInvokePrefix(user)+`Invoke-Command -ComputerName $env:HMA_HOST `+remoteCredentialArg(user)+` -ScriptBlock { 'HMA_OK' }`;out,err:=powershell(script,map[string]string{"HMA_HOST":host,"HMA_USER":user,"HMA_PASS":pass});if err!=nil{return err};if !strings.Contains(out,"HMA_OK"){return errors.New("WinRM test did not return HMA_OK")};return nil}

func createRemoteArchive(src Candidate)(string,error){ host:=strings.TrimSpace(getText(app.srcHost));user:=strings.TrimSpace(getText(app.srcUser));pass:=getText(app.srcPass); stage:=strings.TrimSpace(getText(app.archiveStage));if stage==""{stage=filepath.Join(filepath.VolumeName(src.DataWindows),`\Migration-Data`)}; archive:=filepath.Join(stage,"hmail-data-"+time.Now().Format("20060102-150405")+".7z");script:=`$ErrorActionPreference='Stop';`+remoteInvokePrefix(user)+`$sb={param($seven,$data,$archive) New-Item -ItemType Directory -Force -Path (Split-Path -Parent $archive)|Out-Null; Push-Location $data; try{ & $seven a -t7z $archive '.\*' -mx=1 -mmt=on -y -bso1 -bsp0; if($LASTEXITCODE -gt 1){throw "7-Zip failed with exit code $LASTEXITCODE"} } finally {Pop-Location}}; Invoke-Command -ComputerName $env:HMA_HOST `+remoteCredentialArg(user)+` -ScriptBlock $sb -ArgumentList $env:HMA_7Z,$env:HMA_DATA,$env:HMA_ARCHIVE`;_,err:=powershell(script,map[string]string{"HMA_HOST":host,"HMA_USER":user,"HMA_PASS":pass,"HMA_7Z":src.SevenZip,"HMA_DATA":src.DataWindows,"HMA_ARCHIVE":archive});return archive,err}

func copySingleWithProgress(src,dst,label string)error{fi,err:=os.Stat(src);if err!=nil{return err};in,err:=os.Open(src);if err!=nil{return err};defer in.Close();out,err:=os.Create(dst);if err!=nil{return err};defer out.Close();buf:=make([]byte,4<<20);var done int64;start:=time.Now();for{n,e:=in.Read(buf);if n>0{if _,w:=out.Write(buf[:n]);w!=nil{return w};done+=int64(n);pct:=0;if fi.Size()>0{pct=int(done*100/fi.Size())};rate:=float64(done)/time.Since(start).Seconds();eta:=time.Duration(0);if rate>0{eta=time.Duration(float64(fi.Size()-done)/rate)*time.Second};postProgress(pct,fmt.Sprintf("%s - %s of %s | %.2f MB/s | ETA %s",label,formatBytes(done),formatBytes(fi.Size()),rate/(1<<20),formatDuration(eta)))};if e==io.EOF{break};if e!=nil{return e}};return nil}
func extractArchive(seven,archive,dest string)error{_ = os.MkdirAll(dest,0755);cmd:=exec.Command(seven,"x",archive,"-o"+dest,"-y");b,err:=cmd.CombinedOutput();if err!=nil{return fmt.Errorf("7-Zip extraction failed: %v: %s",err,string(b))};return nil}

func replaceIniDatabaseSection(destIni,srcIni string) error {
    srcMap,err:=parseIni(srcIni);if err!=nil{return err};db:=map[string]string{};for s,v:=range srcMap{if strings.EqualFold(s,"Database"){for k,x:=range v{db[k]=x}}};if len(db)==0{return errors.New("[Database] section not found in source INI")}
    b,err:=os.ReadFile(destIni);if err!=nil{return err};lines:=strings.Split(strings.ReplaceAll(string(b),"\r\n","\n"),"\n");var out []string;skip:=false;replaced:=false
    emit:=func(){out=append(out,"[Database]");keys:=[]string{"Type","Username","Password","PasswordEncryption","Port","Server","Database","Internal"};used:=map[string]bool{};for _,k:=range keys{for dk,dv:=range db{if strings.EqualFold(dk,k){out=append(out,dk+"="+dv);used[dk]=true}}};for k,v:=range db{if !used[k]{out=append(out,k+"="+v)}}}
    for _,ln:=range lines{trim:=strings.TrimSpace(ln);if strings.HasPrefix(trim,"[")&&strings.HasSuffix(trim,"]"){name:=strings.TrimSpace(trim[1:len(trim)-1]);if strings.EqualFold(name,"Database"){if !replaced{emit();replaced=true};skip=true;continue};if skip{skip=false}};if !skip{out=append(out,ln)}};if !replaced{out=append(out,"", "[Database]");for k,v:=range db{out=append(out,k+"="+v)}};return os.WriteFile(destIni,[]byte(strings.Join(out,"\r\n")),0644)
}

func backupDestination(dst Candidate)(string,error){root:=filepath.Join(`C:\ProgramData\hMailServer\MigrationAssistant\Backups`,time.Now().Format("20060102-150405"));if err:=os.MkdirAll(root,0755);err!=nil{return "",err};if err:=copyFile(dst.IniAccess,filepath.Join(root,"hMailServer.ini"));err!=nil{return "",err};if err:=copyFile(dst.DatabaseAccess,filepath.Join(root,"hMailServer.sdf"));err!=nil{return "",err};manifest:=map[string]string{"DestinationIni":dst.IniAccess,"DestinationDatabase":dst.DatabaseAccess,"Created":time.Now().Format(time.RFC3339)};b,_:=json.MarshalIndent(manifest,"","  ");_ = os.WriteFile(filepath.Join(root,"manifest.json"),b,0644);app.lastBackup=root;return root,nil}
func copyFile(src,dst string)error{in,err:=os.Open(src);if err!=nil{return err};defer in.Close();if err:=os.MkdirAll(filepath.Dir(dst),0755);err!=nil{return err};out,err:=os.Create(dst);if err!=nil{return err};_,err=io.Copy(out,in);cerr:=out.Close();if err!=nil{return err};return cerr}

func preserveDestinationOnly(srcData,dstData,backup string)(int,error){root:=filepath.Join(backup,"DestinationOnlyMail");count:=0;err:=filepath.Walk(dstData,func(path string,info os.FileInfo,e error)error{if e!=nil{return nil};if info.IsDir()||!strings.EqualFold(filepath.Ext(path),".eml"){return nil};rel,e:=filepath.Rel(dstData,path);if e!=nil{return nil};if _,e=os.Stat(filepath.Join(srcData,rel));os.IsNotExist(e){to:=filepath.Join(root,rel);if e=copyFile(path,to);e==nil{count++}};return nil});return count,err}

func copyExternalFiles(files []ExternalFile) { if len(files)==0{return}; host:=strings.TrimSpace(getText(app.srcHost));share:=strings.TrimSpace(getText(app.srcShare));for _,f:=range files{if f.Path==""{continue};src:=windowsToUNC(host,share,f.Path);if _,e:=os.Stat(src);e!=nil{logf("WARN","External file unavailable: %s - %s",f.Kind,f.Path);continue};if _,e:=os.Stat(f.Path);e==nil{continue};if e:=copyFile(src,f.Path);e!=nil{logf("WARN","External file copy failed (%s): %v",f.Path,e)}else{logf("INFO","External file copied: %s",f.Path)}} }

func finalCutover() error {
    src,dst,err:=currentCandidates();if err!=nil{return err};if messageBox("Final Cutover will stop hMailServer on the source and destination, perform the final incremental sync, back up the destination configuration/database, replace the destination database, preserve destination-only messages, migrate external files where possible, and then start the destination service.\n\nContinue?","Confirm final cutover",MB_YESNO|MB_ICONWARNING)!=IDYES{return errors.New("cutover cancelled")}
    if checked(app.stopSource){if err:=stopService(true);err!=nil{return err}}else if serviceRunning(true){return errors.New("source hMailServer is still running; stop it before cutover or enable automatic stop")}
    if err:=stopService(false);err!=nil{return err}
    backup,err:=backupDestination(dst);if err!=nil{return err};logf("INFO","Destination backup: %s",backup)
    preserved,_:=preserveDestinationOnly(src.DataAccess,dst.DataAccess,backup);if preserved>0{logf("WARN","Detected and preserved %d destination-only .eml files. They remain on disk and may need Data Directory Synchronizer re-import after database switch.",preserved)}
    threads,_:=strconv.Atoi(strings.TrimSpace(getText(app.threads)));if threads<1{threads=32};if err:=robocopyData(src.DataAccess,dst.DataAccess,"Final mail sync",threads);err!=nil{return err}
    if err:=copyFile(src.DatabaseAccess,dst.DatabaseAccess);err!=nil{return fmt.Errorf("database copy failed: %v",err)}
    if err:=replaceIniDatabaseSection(dst.IniAccess,src.IniAccess);err!=nil{return fmt.Errorf("INI database section migration failed: %v",err)}
    if src.EventsAccess!=""&&dst.EventsAccess!=""{_ = robocopyData(src.EventsAccess,dst.EventsAccess,"Events sync",8)}
    if app.preflight!=nil{copyExternalFiles(app.preflight.ExternalFiles)}
    if err:=startService(false);err!=nil{return err}; logf("INFO","Cutover complete. %d destination-only messages need re-import review.",preserved);return nil
}

func verifyMigration() error {
    _,dst,err:=currentCandidates();if err!=nil{return err};if !serviceRunning(false){return errors.New("destination hMailServer service is not running")};store,e:=robocopyEnumerate(dst.DataAccess);if e!=nil{return e};logf("INFO","Destination mail store: %d .eml files, %s",store.Files,formatBytes(store.Bytes));if app.preflight!=nil{logf("INFO","Source pre-flight mail store: %d .eml files, %s",app.preflight.SourceStore.Files,formatBytes(app.preflight.SourceStore.Bytes));if store.Files<app.preflight.SourceStore.Files{logf("WARN","Destination physical .eml count is lower than the source pre-flight count")}}
    if pw:=getText(app.dstAdminPass);pw!=""{dstSnap,e:=runCOMSnapshot(false,pw);if e!=nil{logf("WARN","Destination COM verification failed: %v",e)}else if app.preflight!=nil&&app.preflight.SourceSnapshot!=nil{src:=app.preflight.SourceSnapshot;if src.Domains!=dstSnap.Domains||src.Accounts!=dstSnap.Accounts||src.Messages!=dstSnap.Messages{logf("WARN","COM counts differ from source snapshot: source domains/accounts/messages %d/%d/%d, destination %d/%d/%d",src.Domains,src.Accounts,src.Messages,dstSnap.Domains,dstSnap.Accounts,dstSnap.Messages)}else{logf("INFO","COM source/destination counts match")}}}
    report,err:=generateReport(store);if err==nil{app.lastReport=report;logf("INFO","HTML migration report: %s",report)};return nil
}

func dnsChecks() error {
    var domains []string;if app.preflight!=nil&&app.preflight.SourceSnapshot!=nil{domains=app.preflight.SourceSnapshot.DomainsList};if len(domains)==0{return errors.New("no domains available; enter an hMailServer Administrator password and validate first")};for _,d:=range domains{mx,e:=net.LookupMX(d);if e!=nil{logf("WARN","MX %s -> unresolved (%v)",d,e)}else{var xs []string;for _,m:=range mx{xs=append(xs,m.Host)};logf("INFO","MX %s -> %s",d,strings.Join(xs,", "))};txt,_:=net.LookupTXT(d);var spf []string;for _,t:=range txt{if strings.HasPrefix(strings.ToLower(t),"v=spf1"){spf=append(spf,t)}};logf("INFO","SPF %s -> %s",d,strings.Join(spf," | "))};if ip:=strings.TrimSpace(getText(app.publicIP));ip!=""{names,e:=net.LookupAddr(ip);if e!=nil{logf("WARN","PTR %s -> unresolved",ip)}else{logf("INFO","PTR %s -> %s",ip,strings.Join(names,", "))}};for _,p:=range []int{25,465,587,110,995,143,993}{c,e:=net.DialTimeout("tcp",fmt.Sprintf("127.0.0.1:%d",p),700*time.Millisecond);if e==nil{_ = c.Close();logf("INFO","Local listener check: TCP %d OPEN",p)}};return nil
}

func generateReport(dstStore StoreInfo)(string,error){root:=`C:\ProgramData\hMailServer\MigrationAssistant\Reports`;if err:=os.MkdirAll(root,0755);err!=nil{return "",err};p:=filepath.Join(root,"Migration-Report-"+time.Now().Format("20060102-150405")+".html");srcText:="Not available";if app.preflight!=nil{srcText=fmt.Sprintf("%d .eml / %s",app.preflight.SourceStore.Files,formatBytes(app.preflight.SourceStore.Bytes))};body:=fmt.Sprintf(`<!doctype html><html><head><meta charset="utf-8"><title>hMailServer Migration Report</title><style>body{font-family:Segoe UI,Arial;margin:40px;color:#1f2937}h1{color:#17324d}.card{border:1px solid #d1d5db;border-radius:8px;padding:18px;margin:16px 0}pre{white-space:pre-wrap;background:#f7f8fa;padding:14px;border-radius:6px}table{border-collapse:collapse}td,th{border:1px solid #ddd;padding:8px;text-align:left}</style></head><body><h1>hMailServer Migration Report</h1><p>Generated %s by Migration Assistant v%s</p><div class="card"><h2>Mail store</h2><table><tr><th>Source</th><td>%s</td></tr><tr><th>Destination</th><td>%d .eml / %s</td></tr><tr><th>Backup</th><td>%s</td></tr></table></div><div class="card"><h2>Activity log</h2><pre>%s</pre></div></body></html>`,time.Now().Format(time.RFC1123),appVersion,html.EscapeString(srcText),dstStore.Files,html.EscapeString(formatBytes(dstStore.Bytes)),html.EscapeString(app.lastBackup),html.EscapeString(readLog()));return p,os.WriteFile(p,[]byte(body),0644)}
func readLog()string{b,_:=os.ReadFile(ensureLog());return string(b)}

func rollback() error { if app.lastBackup==""{return errors.New("no migration backup has been created in this session")};if messageBox("Restore the most recent pre-migration hMailServer.ini and database backup? Mail files will not be deleted.","Confirm rollback",MB_YESNO|MB_ICONWARNING)!=IDYES{return errors.New("rollback cancelled")};_,dst,err:=currentCandidates();if err!=nil{return err};_ = stopService(false);if err:=copyFile(filepath.Join(app.lastBackup,"hMailServer.ini"),dst.IniAccess);err!=nil{return err};if err:=copyFile(filepath.Join(app.lastBackup,"hMailServer.sdf"),dst.DatabaseAccess);err!=nil{return err};return startService(false) }

func launchReimport() error { _,dst,err:=currentCandidates();if err!=nil{return err};base:=filepath.Dir(filepath.Dir(dst.IniWindows));candidates:=[]string{filepath.Join(base,"Addons","DataDirectorySynchronizer","DataDirectorySynchronizer.exe"),filepath.Join(dst.ProgramFolder,"Addons","DataDirectorySynchronizer","DataDirectorySynchronizer.exe")};for _,p:=range candidates{if _,e:=os.Stat(p);e==nil{return exec.Command(p).Start()}};return errors.New("Data Directory Synchronizer was not found under the selected hMailServer installation") }

func benchmarkConnection() error { src,_,err:=currentCandidates();if err!=nil{return err};tmp,err:=os.CreateTemp(src.DataAccess,"hma-bench-*.bin");if err!=nil{return err};name:=tmp.Name();defer os.Remove(name);block:=make([]byte,1<<20);for i:=0;i<32;i++{if _,e:=tmp.Write(block);e!=nil{tmp.Close();return e}};tmp.Close();dst:=filepath.Join(os.TempDir(),filepath.Base(name));start:=time.Now();if err:=copyFile(name,dst);err!=nil{return err};d:=time.Since(start);fi,_:=os.Stat(dst);_ = os.Remove(dst);if fi==nil{return errors.New("benchmark copy failed")};rate:=float64(fi.Size())/d.Seconds();logf("INFO","Connection benchmark: %.2f MB/s",rate/(1<<20));if app.preflight!=nil&&rate>0{eta:=time.Duration(float64(app.preflight.SourceStore.Bytes)/rate)*time.Second;logf("INFO","Estimated full-store transfer time: %s",formatDuration(eta))};return nil }

func discovery() {setBusy(true,"Discovering hMailServer installations...");go func(){src,e1:=discoverCandidates(true);dst,e2:=discoverCandidates(false);var e error;if e1!=nil{e=e1}else if e2!=nil{e=e2};r:=&DiscoveryResult{Source:src,Destination:dst,Err:e};pPostMessageW.Call(uintptr(app.hwnd),APP_DISCOVER,0,uintptr(unsafe.Pointer(r)))}()}
func validateAsync(){setBusy(true,"Validating migration readiness...");go func(){r,e:=buildPreflight();if e!=nil{logf("ERROR","Pre-flight failed: %v",e);postString(APP_STATUS,"Pre-flight failed")}else{app.preflight=r;logf("INFO","Mail store\r\n  Source: %d .eml files, %s total data\r\n  Destination: %d .eml files, %s total data",r.SourceStore.Files,formatBytes(r.SourceStore.Bytes),r.DestinationStore.Files,formatBytes(r.DestinationStore.Bytes));for _,w:=range r.Warnings{logf("WARN","%s",w)};postString(APP_STATUS,"READY FOR MIGRATION")};pPostMessageW.Call(uintptr(app.hwnd),APP_DONE,0,0)}()}
func benchmarkAsync(){setBusy(true,"Benchmarking connection...");go func(){if e:=benchmarkConnection();e!=nil{logf("ERROR","Benchmark failed: %v",e)};pPostMessageW.Call(uintptr(app.hwnd),APP_DONE,0,0)}()}
func initialSyncAsync(){setBusy(true,"Initial mail sync running...");go func(){if e:=initialSync();e!=nil{logf("ERROR","Initial mail sync failed: %v",e)}else{logf("INFO","Initial mail sync completed")};pPostMessageW.Call(uintptr(app.hwnd),APP_DONE,0,0)}()}
func cutoverAsync(){setBusy(true,"Final cutover running...");go func(){if e:=finalCutover();e!=nil{logf("ERROR","Final cutover failed: %v",e)}else{logf("INFO","Final cutover completed")};pPostMessageW.Call(uintptr(app.hwnd),APP_DONE,0,0)}()}
func verifyAsync(){setBusy(true,"Verifying migration...");go func(){if e:=verifyMigration();e!=nil{logf("ERROR","Verification failed: %v",e)}else{logf("INFO","Verification completed")};pPostMessageW.Call(uintptr(app.hwnd),APP_DONE,0,0)}()}
func dnsAsync(){setBusy(true,"Running DNS / mail-flow checks...");go func(){if e:=dnsChecks();e!=nil{logf("ERROR","DNS checks failed: %v",e)};pPostMessageW.Call(uintptr(app.hwnd),APP_DONE,0,0)}()}

func populateCandidateFields(source bool,idx int){var c Candidate;if source{if idx<0||idx>=len(app.sourceCandidates){return};c=app.sourceCandidates[idx];app.sourceSelected=&app.sourceCandidates[idx];setText(app.srcIni,c.IniAccess);setText(app.srcDB,c.DatabaseAccess);setText(app.srcData,c.DataAccess);setText(app.srcEvents,c.EventsAccess)}else{if idx<0||idx>=len(app.destCandidates){return};c=app.destCandidates[idx];app.destSelected=&app.destCandidates[idx];setText(app.dstIni,c.IniAccess);setText(app.dstDB,c.DatabaseAccess);setText(app.dstData,c.DataAccess);setText(app.dstEvents,c.EventsAccess)}}

const eulaText = `END USER LICENCE AGREEMENT

This End User Licence Agreement ("Agreement") governs use of hMailServer Migration Assistant ("Software"). By using the Software, you agree to the terms below.

1. LICENCE
You are granted a non-exclusive, non-transferable licence to use the Software for lawful hMailServer migration, verification, backup and related administrative purposes.

2. AUTHORISED USE
You may use the Software only on systems, accounts and data that you own or are authorised to administer. You are responsible for obtaining any permissions required to access source and destination hosts, mail data, credentials, certificates and network resources.

3. BACKUPS AND MIGRATION RESPONSIBILITY
Mail-system migration can result in interruption, configuration changes or data loss if performed incorrectly. You are responsible for maintaining verified backups of the source environment and for validating the destination before decommissioning the source. The Software includes safeguards and rollback features, but these do not replace an independent backup.

4. THIRD-PARTY SOFTWARE
The Software may call or interoperate with Microsoft Windows components, hMailServer, Robocopy, PowerShell, 7-Zip, Tailscale or other third-party products. Those products remain subject to their own licences and terms. No third-party component is licensed to you under this Agreement.

5. SECURITY AND CREDENTIALS
You are responsible for protecting administrative credentials, backup files, mail data, private keys and migration reports.

6. DATA PROTECTION
You are responsible for complying with applicable privacy, data-protection, retention, employment and communications laws when transferring or processing email data.

7. NO WARRANTY
The Software is provided "as is" without warranties of any kind, to the maximum extent permitted by applicable law.

8. LIMITATION OF LIABILITY
To the maximum extent permitted by applicable law, the author is not liable for indirect, incidental, special or consequential loss arising from use of the Software.

Copyright © 2026 Martyn Beech. All rights reserved.`

const aboutText = `hMailServer Migration Assistant
Version 1.0

Author: Martyn Beech

Enterprise migration utility for moving hMailServer Microsoft SQL Server Compact Edition installations between Windows hosts.

Capabilities
• Source and destination installation discovery
• MSSQL CE database and mail-data migration
• Robocopy and optional 7-Zip transfer staging
• Live migration logging and transfer status
• Certificate and DKIM discovery
• Destination-only mail preservation
• Post-cutover verification and reporting
• Configuration/database rollback

Independent administration utility. hMailServer and other referenced product names remain the property of their respective owners. No affiliation or endorsement is implied.

Licence: Help > End User Licence Agreement`

func showTextWindow(title,text string){messageBox(text,title,MB_OK|MB_ICONINFORMATION)}

func setupMenu(hwnd syscall.Handle){m,_,_:=pCreateMenu.Call();file,_,_:=pCreatePopupMenu.Call();help,_,_:=pCreatePopupMenu.Call();pAppendMenuW.Call(file,MF_STRING,ID_MENU_LOGS,uintptr(unsafe.Pointer(u16("Open logs and reports"))));pAppendMenuW.Call(file,MF_SEPARATOR,0,0);pAppendMenuW.Call(file,MF_STRING,ID_MENU_EXIT,uintptr(unsafe.Pointer(u16("Exit"))));pAppendMenuW.Call(help,MF_STRING,ID_MENU_EULA,uintptr(unsafe.Pointer(u16("End User Licence Agreement"))));pAppendMenuW.Call(help,MF_STRING,ID_MENU_ABOUT,uintptr(unsafe.Pointer(u16("About"))));pAppendMenuW.Call(m,MF_POPUP,file,uintptr(unsafe.Pointer(u16("File"))));pAppendMenuW.Call(m,MF_POPUP,help,uintptr(unsafe.Pointer(u16("Help"))));pSetMenu.Call(uintptr(hwnd),m)}

func createUI(hwnd syscall.Handle){app=&App{hwnd:hwnd};app.font=createFont(16);setupMenu(hwnd)
    addLabel("hMailServer Migration Assistant",20,12,520,28,hwnd); addLabel("Version 1.0",1180,16,100,22,hwnd)
    addGroup("1  Source connection and discovery",12,48,1338,116,hwnd)
    labels:=[]struct{t string;x int}{ {"Source host / IP address",28},{"Admin share",316},{"Windows user",550},{"Password",786},{"New public IP (optional)",1022} }
    for _,l:=range labels{addLabel(l.t,l.x,70,210,20,hwnd)}
    app.srcHost=addEdit("",28,92,270,26,hwnd,false);app.srcShare=addEdit("C$",316,92,216,26,hwnd,false);app.srcUser=addEdit("",550,92,218,26,hwnd,false);app.srcPass=addEdit("",786,92,218,26,hwnd,true);app.publicIP=addEdit("",1022,92,216,26,hwnd,false)
    addButton("1  Discover installations",28,126,210,28,ID_DISCOVER,hwnd);app.stopSource=addCheck("Stop source hMailServer automatically during cutover",260,129,360,24,0,hwnd,true);app.allowMismatch=addCheck("Allow hMailServer version mismatch (advanced)",640,129,360,24,0,hwnd,false)
    addGroup("Source installation",12,176,660,292,hwnd);addGroup("Destination installation",690,176,660,292,hwnd)
    addLabel("Detected installation",28,200,180,20,hwnd);app.srcCombo=addCombo(28,222,500,200,0,hwnd);addButton("Use selected",540,222,112,27,ID_USE_SRC,hwnd)
    addLabel("Detected installation",706,200,180,20,hwnd);app.dstCombo=addCombo(706,222,500,200,0,hwnd);addButton("Use selected",1218,222,112,27,ID_USE_DST,hwnd)
    y:=260; names:=[]string{"INI","Database","Data","Events"}; srcPtrs:=[]*syscall.Handle{&app.srcIni,&app.srcDB,&app.srcData,&app.srcEvents};dstPtrs:=[]*syscall.Handle{&app.dstIni,&app.dstDB,&app.dstData,&app.dstEvents};srcIDs:=[]int{ID_BROWSE_SI,ID_BROWSE_SD,ID_BROWSE_SDA,ID_BROWSE_SE};dstIDs:=[]int{ID_BROWSE_DI,ID_BROWSE_DD,ID_BROWSE_DDA,ID_BROWSE_DE};for i,n:=range names{addLabel(n,28,y+5,70,20,hwnd);*srcPtrs[i]=addEdit("",102,y,450,26,hwnd,false);addButton("Browse",562,y,90,26,srcIDs[i],hwnd);addLabel(n,706,y+5,70,20,hwnd);*dstPtrs[i]=addEdit("",782,y,450,26,hwnd,false);addButton("Browse",1240,y,90,26,dstIDs[i],hwnd);y+=36}
    addLabel("hMailServer Administrator password",28,408,290,20,hwnd);app.srcAdminPass=addEdit("",336,404,316,26,hwnd,true);addLabel("Optional — enables source COM comparison",28,434,320,20,hwnd);addLabel("Archive staging folder on source",28,456,280,20,hwnd);app.archiveStage=addEdit("",336,452,316,26,hwnd,false)
    addLabel("hMailServer Administrator password",706,408,290,20,hwnd);app.dstAdminPass=addEdit("",1014,404,316,26,hwnd,true);addLabel("Optional — enables post-cutover COM verification",706,434,420,20,hwnd)
    addGroup("2  Readiness and transfer method",12,480,1338,110,hwnd);addButton("2  Validate / pre-flight",28,510,180,30,ID_VALIDATE,hwnd);addButton("Benchmark connection",220,510,170,30,ID_BENCHMARK,hwnd);addLabel("Transfer method",410,495,160,20,hwnd);app.transferMethod=addCombo(410,518,240,150,0,hwnd);comboAdd(app.transferMethod,"Automatic");comboAdd(app.transferMethod,"Standard Robocopy");comboAdd(app.transferMethod,"7-Zip initial seed");comboSelect(app.transferMethod,0);addLabel("Robocopy threads",670,495,150,20,hwnd);app.threads=addEdit("32",670,518,90,26,hwnd,false);app.ask7zip=addCheck("Ask whether to 7-Zip mail store when many .eml files are detected",790,516,500,26,0,hwnd,true)
    addGroup("3  Migration workflow",12,602,1338,104,hwnd);addButton("3  Initial mail sync",28,632,180,32,ID_INITIAL,hwnd);addButton("4  Final cutover",220,632,170,32,ID_CUTOVER,hwnd);addButton("5  Verify migration",402,632,180,32,ID_VERIFY,hwnd);addButton("DNS / mail-flow",594,632,160,32,ID_DNS,hwnd);addButton("Re-import mail",766,632,150,32,ID_REIMPORT,hwnd);addButton("Roll back config / DB",928,632,180,32,ID_ROLLBACK,hwnd);addButton("Open logs",1120,632,140,32,ID_OPENLOGS,hwnd)
    app.progress=createControl("msctls_progress32","",0,28,672,850,20,hwnd,0);pSendMessageW.Call(uintptr(app.progress),PBM_SETRANGE32,0,100);app.speed=addLabel("Transfer rate: --",900,670,250,22,hwnd);app.status=addLabel("Ready",1160,670,160,22,hwnd)
    addGroup("Activity",12,718,1338,200,hwnd);app.log=createControl("EDIT","",WS_BORDER|WS_VSCROLL|ES_MULTILINE|ES_AUTOVSCROLL|ES_READONLY,28,746,1304,154,hwnd,0);setFont(app.log,app.font);ensureLog();logf("INFO","%s v%s started",appName,appVersion)
}

func appendLogUI(line string){old:=getText(app.log);if len(old)>60000{old=old[len(old)-45000:]};setText(app.log,old+line+"\r\n")}

func openLogs(){root:=`C:\ProgramData\hMailServer\MigrationAssistant`;_ = exec.Command("explorer.exe",root).Start()}

func windowProc(hwnd syscall.Handle,msg uint32,wParam,lParam uintptr) uintptr {
    switch msg {
    case WM_CREATE: createUI(hwnd);return 0
    case WM_COMMAND:
        id:=int(loword(wParam));if app!=nil&&app.busy&&id!=ID_MENU_EXIT{return 0}
        switch id {
        case ID_DISCOVER:discovery()
        case ID_VALIDATE:validateAsync()
        case ID_BENCHMARK:benchmarkAsync()
        case ID_INITIAL:initialSyncAsync()
        case ID_CUTOVER:cutoverAsync()
        case ID_VERIFY:verifyAsync()
        case ID_DNS:dnsAsync()
        case ID_ROLLBACK:go func(){if e:=rollback();e!=nil{logf("ERROR","Rollback failed: %v",e)}else{logf("INFO","Rollback completed")}}()
        case ID_REIMPORT:if e:=launchReimport();e!=nil{messageBox(e.Error(),"Re-import mail",MB_OK|MB_ICONERROR)}
        case ID_OPENLOGS,ID_MENU_LOGS:openLogs()
        case ID_USE_SRC:populateCandidateFields(true,comboSelected(app.srcCombo))
        case ID_USE_DST:populateCandidateFields(false,comboSelected(app.dstCombo))
        case ID_MENU_EULA:showTextWindow("End User Licence Agreement",eulaText)
        case ID_MENU_ABOUT:showTextWindow("About",aboutText)
        case ID_MENU_EXIT:pPostQuitMessage.Call(0)
        }
        return 0
    case APP_LOG: p:=(*string)(unsafe.Pointer(lParam));appendLogUI(*p);return 0
    case APP_STATUS:p:=(*string)(unsafe.Pointer(lParam));setText(app.status,*p);return 0
    case APP_PROGRESS:p:=(*ProgressInfo)(unsafe.Pointer(lParam));pSendMessageW.Call(uintptr(app.progress),PBM_SETPOS,uintptr(p.Percent),0);setText(app.speed,p.Text);return 0
    case APP_DONE:app.mu.Lock();app.busy=false;app.mu.Unlock();if getText(app.status)!="READY FOR MIGRATION"{setText(app.status,"Ready")};return 0
    case APP_DISCOVER:r:=(*DiscoveryResult)(unsafe.Pointer(lParam));app.busy=false;if r.Err!=nil{logf("ERROR","Discovery error: %v",r.Err);setText(app.status,"Discovery error");return 0};app.sourceCandidates=r.Source;app.destCandidates=r.Destination;comboReset(app.srcCombo);comboReset(app.dstCombo);for _,c:=range r.Source{comboAdd(app.srcCombo,fmt.Sprintf("%s | DB %s | %s",c.IniWindows,formatBytes(c.DatabaseSize),c.Version))};for _,c:=range r.Destination{comboAdd(app.dstCombo,fmt.Sprintf("%s | DB %s | %s",c.IniWindows,formatBytes(c.DatabaseSize),c.Version))};if len(r.Source)==1{comboSelect(app.srcCombo,0);populateCandidateFields(true,0)};if len(r.Destination)==1{comboSelect(app.dstCombo,0);populateCandidateFields(false,0)};logf("INFO","Discovery completed: %d source candidate(s), %d destination candidate(s)",len(r.Source),len(r.Destination));setText(app.status,"Ready");return 0
    case WM_CLOSE:pPostQuitMessage.Call(0);return 0
    case WM_DESTROY:disconnectSource();pPostQuitMessage.Call(0);return 0
    }
    r,_,_:=pDefWindowProcW.Call(uintptr(hwnd),uintptr(msg),wParam,lParam);return r
}

func main(){
    if elevateIfNeeded(){return}
    pInitCommon.Call()
    hInst,_,_:=pGetModuleHandleW.Call(0);cursor,_,_:=pLoadCursorW.Call(0,IDC_ARROW);class:=u16("HMailMigrationAssistantWindow");wc:=WNDCLASSEX{CbSize:uint32(unsafe.Sizeof(WNDCLASSEX{})),LpfnWndProc:syscall.NewCallback(windowProc),HInstance:syscall.Handle(hInst),HCursor:syscall.Handle(cursor),HbrBackground:syscall.Handle(6),LpszClassName:class};pRegisterClassExW.Call(uintptr(unsafe.Pointer(&wc)))
    hwnd,_,_:=pCreateWindowExW.Call(0,uintptr(unsafe.Pointer(class)),uintptr(unsafe.Pointer(u16(appName+"  •  v"+appVersion))),WS_OVERLAPPEDWINDOW|WS_VISIBLE,50,30,1380,980,0,0,hInst,0);if hwnd==0{return};pShowWindow.Call(hwnd,SW_SHOW);pUpdateWindow.Call(hwnd)
    var msg MSG;for{r,_,_:=pGetMessageW.Call(uintptr(unsafe.Pointer(&msg)),0,0,0);if int32(r)<=0{break};pTranslateMessage.Call(uintptr(unsafe.Pointer(&msg)));pDispatchMessageW.Call(uintptr(unsafe.Pointer(&msg)))}
}
