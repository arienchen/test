@if (true == false) @end /*
@echo off 
cscript //nologo //e:javascript "%~dpnx0" %*
goto :EOF */
/**
 * CHECKPC.CMD
 * 個人 PC 自行檢核
 * 
 * LAST MODIFIED: V1.0 20180327 陳清祥
 * COPYRIGHT    : 台北富邦銀行/稽核室
 *
 * 程式說明
 * --------------------------------------------------------
 * 
 *
 * 修改記錄
 * -------------------------------------------------------
 * 20180327 V1.0 陳清祥
 *     新增
 *
 */

// CONSTANTS 
var WSH_FAILED   = 2;
var WSH_FINISHED = 1;
var WSH_RUNNING  = 0;
var Err = {ExitCode:null,StdOut:null,StdErr:null};

// FUNCTIONS

/**
 * show_error(e)
 * 顯示 RUNTIME 錯誤
 */
function show_error(e) {
    if (!e) {
        return;
    }

    var msg = "執行錯誤： " + e.number + "\n"
            + "錯誤訊息： " + e.message + "\n"
            ;

    var shell = WScript.CreateObject("WScript.Shell");

    shell.Popup(msg);

    shell = null;                     
}

/**
 * 將 Date 轉換為 YYYYMMDD
 */
Date.prototype.ymd = function(sep) {
    var     y = this.getFullYear();
    var     m = this.getMonth() + 1;
    var     d = this.getDate();
    
    if (sep === undefined) {
        sep = "";
    }

    s = (y < 10)   ? "000" : 
        (y < 100)  ? "00"  :
        (y < 1000) ? "0"   : ""
      ;

    return s + y.toString()
           + sep 
           + ((m < 10) ? "0" : "")
           + m.toString()
           + sep 
           + ((d < 10) ? "0" : "")
           + d.toString()
           ;     
} 

/**
 * 將 Date 轉換為 HHMMSS
 */
Date.prototype.hms = function(sep) {
    var     h = this.getHours();
    var     m = this.getMinutes();
    var     s = this.getSeconds();
    
    if (sep === undefined) {
        sep = "";
    }

    return ((h < 10) ? "0" : "") 
           + h.toString()
           + sep 
           + ((m < 10) ? "0" : "")
           + m.toString()
           + sep 
           + ((s < 10) ? "0" : "")
           + s.toString()
           ;     
} 

/**
 * 將 Date 轉換為 YYYY-MM-DD hh:mm:ss
 */
Date.prototype.toString = function() {
    return this.ymd("-") + " " + this.hms(":");
} 

/**
 * 將 Date 轉換為 YYYYMMDDhhmmss
 */
Date.prototype.timestamp = function() {
    return this.ymd() + this.hms();
} 

/**
 * 將 Unix Epoch Time 轉換為 Date 
 * Epoch = 1970/1/1 開始的秒數
 */
Date.epoch = function(sec) {
    var   dt = new Date();

    dt.setTime(sec * 1000);
    return dt;
}

String.prototype.toDate = function(fmt, sep) {
    if (fmt.toUpperCase() == "MDY") {
        if (sep === undefined) {
            sep = "/";
        }

        
        var s = this.split(sep);

        if (s.length != 3) {
            return null;
        }

        var d = new Date(s[2],s[0],s[1]);

        return isNaN(d) ? Null : d;
    }

    return null;
} 

/**
 * 讀取 Registry 值
 */
function reg_get(name) {
    var shell = WScript.CreateObject("WScript.Shell");

    return shell.RegRead(name);
}

/**
 * cmd_str(cmd) 
 * 執行命令列，並回傳執行結果為字串
 *
 * Parameters:
 * cmd       Commands to run in shell
 *
 * Return
 * String    StdOut 
 * 
 */
function cmd_str(cmd) {
    var shell = WScript.CreateObject("WScript.Shell");
    var out   = null;

    try {
        var exec  = shell.Exec(cmd);
        out   = exec.StdOut;
    } catch(e) {
        Err.StdErr = e.message;
        Err.ExitCode = e.number;
    }

    return out;
}

/**
 * cmd_code(cmd) 
 * 執行命令列，並回傳錯誤碼
 *
 * Parameters:
 * cmd       Commands to run in shell
 *
 * Return
 * String    StdOut 
 * 
 */
function cmd_code(cmd) {
    var shell = WScript.CreateObject("WScript.Shell");
    
    try {
        var exec  = shell.Exec(cmd);

        Err.StdOut = exec.StdOut.ReadAll();
        Err.ExitCode = exec.ExitCode;
        Err.StdErr = exec.StdErr.ReadAll();

        return exec.ExitCode;
    } catch(e) {
        Err.StdErr = e.message;
        Err.ExitCode = e.number;
        return e.number;
    }
}

/**
 * WMI 
 */
var WMI_RETURN_IMMEDIATELY = 0x10;
var WMI_FORWARD_ONLY       = 0x20;

var WMI = new Object();

WMI.LogonType = function(val) {
    if (!val) {
        return null;  
    }   

    switch(val) {
    case 0 : return "System";
    case 2 : return "Interactive";
    case 3 : return "Network";
    case 4 : return "Batch";
    case 5 : return "Service";
    case 6 : return "Proxy";
    case 7 : return "Unlock";
    case 8 : return "NetworkClearText";
    case 9 : return "NewCredentials";
    case 10: return "RemoteInteractive";
    case 11: return "CachedInteractive";
    case 12: return "cachedRemoteInteractive";
    case 13: return "CachedUnlock";
    default: return "###";
    }
}

/**
 * run_wmi() 
 * 執行 WMIC 查詢，並回傳為 LIST OBJECT
 *  
 */
function run_wmi(qry) {
    var wmi  = GetObject("winmgmts:\\\\.\\root\\CIMV2");
    var data = wmi.ExecQuery(qry, "WQL", WMI_RETURN_IMMEDIATELY | WMI_FORWARD_ONLY);
    return new Enumerator(data); 
}

/**
 * 將 WMI 日期，轉換為 JScript Date 
 * YYYYMMDDhhmmss.SSSSSS -> Date
 */
function wmi_date(dt) {
    
    if (!dt || dt.length <= 21) {
        // INVALID DATA 
        return null;
    }
    
    var     d = null;

    try {
        d = new Date(dt.substring(0,4),
                     dt.substring(4,6) - 1,
                     dt.substring(6,8),
                     dt.substring(8,10),
                     dt.substring(10,12),
                     dt.substring(12,21)
                    );
    } catch(e) {
        show_error(e);        
    }

    return d;
}

/**
 * wmi_nic()
 * 取得網卡 IPADDR 等資訊
 * 同
 * wmic nicconfig where(IPE=TRUE) LIST FULL
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Description                 
 * +IPAddress                    
 * MACAddress                   
 * +DefaultIPGateway Gateway     
 * DHCPEnabled      
 * DHCPServer
 * DNSDomain
 * +DNSDomainSuffixSearchOrder 
 * DNSHostName
 * +DNSServerSearchOrder  
 * IPSubnet
 * ServiceName
 * 
 */
function wmi_nic() {
    var nic = run_wmi("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=TRUE");

    if (nic.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !nic.atEnd() ; nic.moveNext()) {
        var  item = nic.item();
        var  obj  = new Object();

        obj.Description = item.Description;
        obj.MACAddress  = item.MACAddress;
        obj.DHCPEnabled = item.DHCPEnabled;
        obj.DHCPServer  = item.DHCPServer;
        obj.DNSDomain   = item.DNSDomain;
        obj.DNSHostName = item.DNSHostName;
        obj.ServiceName = item.ServiceName;
        obj.IPSubnet    = item.IPSubnet;
        obj.IPAddress   = item.IPAddress.toArray();
        
        obj.DefaultIPGateway = item.DefaultIPGateway.toArray();
        obj.DNSDomainSuffixSearchOrder = item.DNSDomainSuffixSearchOrder.toArray();
        obj.DNSServerSearchOrder = item.DNSServerSearchOrder.toArray();

        list.push(obj);
        
    }

    return list;
}


/**
 * wmi_os()
 * 同
 * wmic os 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Caption
 * CSDVersion
 * CSName 
 * FreePhysicalMemory
 * FreeSpaceInPagingFiles
 * FreeVirtualMemory
 * InstallDate
 * LastBootUpTime
 * LocalDateTime
 * NumberOfUsers
 * OSArchitecture
 * RegisteredUser
 * SerialNumber
 * ServicePackMajorVersion
 * ServicePackMinorVersion
 * SystemDevice
 * SystemDirectory
 * SystemDrive
 * Version
 * WindowsDirectory
 * 
 */
function wmi_os() {
    var data = run_wmi("SELECT * FROM Win32_OperatingSystem ");

    if (data.atEnd()) {
        return null;
    }

    // 只有一筆資料
    var  item = data.item();
    var  obj  = new Object();

    obj.Caption                 = item.Caption;
    obj.CSDVersion              = item.CSDVersioin;
    obj.CSName                  = item.CSName;
    obj.InstallDate             = wmi_date(item.InstallDate);
    obj.LastBootUpTime          = wmi_date(item.LastBootUpTime);
    obj.LocalDateTime           = wmi_date(item.LocalDateTime);
    obj.NumberOfUsers           = item.NumberOfUsers;
    obj.OSArchitecture          = item.OSArchitecture;
    obj.RegisteredUser          = item.RegisteredUser;
    obj.SerialNumber            = item.SerialNumber;
    obj.ServicePackMajorVersion = item.ServicePackMajorVersion;
    obj.ServicePackMinorVersion = item.ServicePackMinorVersion;
    obj.SystemDevice            = item.SystemDevice;
    obj.SystemDirectory         = item.SystemDirectory;
    obj.SystemDrive             = item.SystemDrive;
    obj.Version                 = item.Version;
    obj.WindowsDirectory        = item.WindowsDirectory;
    obj.FreePhysicalMemory      = item.FreePhysicalMemory;
    obj.FreeSpaceInPagingFiles  = item.FreeSpaceInPagingFiles;
    obj.FreeVirtualMemory       = item.FreeVirtualMemory;
    
    return obj;
}

/**
 * wmi_csproduct()
 * 同
 * wmic csproduct 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * IdentifyingNumber
 * Name 
 * UUID
 * Vendor
 * Version
 * 
 */
function wmi_csproduct() {
    var data = run_wmi("SELECT * FROM Win32_ComputerSystemProduct");

    if (data.atEnd()) {
        return null;
    }

    // 只有一筆資料
    var  item = data.item();
    var  obj  = new Object();

    obj.IdentifingNumber        = item.IdentifyingNumber;
    obj.Name                    = item.Name;
    obj.UUID                    = item.UUID;
    obj.Vendor                  = item.Vendor;
    obj.Version                 = item.Version;
    
    return obj;
}

/**
 * wmi_logon()
 * 同
 * wmic logon
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * AuthenticationPackage
 * LogonId
 * LogonType
 * StartTime
 * Status
 * 
 */
function wmi_logon() {
    var data = run_wmi("SELECT * FROM Win32_LogonSession");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.AuthenticationPackage = item.AuthenticationPackage;
        obj.LogonId               = item.LogonId;
        obj.LogonType             = item.LogonType;
        obj.StartTime             = wmi_date(item.StartTime);
        obj.Status                = item.Status;
        
        list.push(obj);
        
    }

    return list;
}

/**
 * wmi_netlogin()
 * 同
 * wmic netlogin WHERE (UserType IS NOT NULL)
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * FullName
 * Name
 * UserType
 * 
 */
function wmi_netlogin() {
    var data = run_wmi("SELECT * FROM Win32_NetworkLoginProfile WHERE UserType IS NOT NULL");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.FullName              = item.FullName;
        obj.Name                  = item.Name;
        obj.UserType              = item.UserType;
        
        list.push(obj);
        
    }

    return list;
}

/**
 * wmi_user_account() 本機使用者帳號
 * 同
 * wmic useraccount 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Description
 * Disabled
 * FullName 
 * LocalAccount
 * Lockout 
 * Name 
 * PasswrodChangeable 
 * PasswordExpires 
 * PasswordRequired
 * Status  
 * 
 */
function wmi_user_account() {
    var data = run_wmi("SELECT * FROM Win32_UserAccount ");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.Description           = item.Description;
        obj.Disabled              = item.Disabled;
        obj.LocalAccount          = item.LocalAccount 
        obj.Lockout               = item.Lockout 
        obj.Name                  = item.Name 
        obj.PasswordChangeable    = item.PasswordChangeable
        obj.PasswordExpires       = item.PasswordExpires 
        obj.PasswordRequired      = item.PasswordRequired 
        obj.Status                = item.Status    

        list.push(obj);
        
    }

    return list;
}

/**
 * wmi_group() 本機群組
 * 同
 * wmic group 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Description
 * Domain 
 * Name 
 * Status 
 * +Users
 * 
 */
function wmi_group() {
    var data = run_wmi("SELECT * FROM Win32_Group ");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.Description           = item.Description;
        obj.Domain                = item.Domain;
        obj.Name                  = item.Name; 
        obj.Status                = item.Status; 
        obj.Users                 = net_local_group(item.Name);
        list.push(obj);
        
    }

    return list;
}

/**
 * wmi_group_user() 
 * 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Description
 * Disabled
 * FullName 
 * LocalAccount
 * Lockout 
 * Name 
 * PasswrodChangeable 
 * PasswordExpires 
 * PasswordRequired
 * Status  
 * 
 */
function wmi_group_user() {
    var sql = "SELECT PartComponent.Name  FROM Win32_GroupUser "
            + "WHERE "
            + "groupcomponent=\"win32_group.name=\\\"administrators\\\",domain=\\\"NB05-A-000837\\\"\""
            ;
              
    var data = run_wmi(sql);

    WSH.Echo(sql);

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        WSH.Echo("-> " + item.Name);
        list.push(obj);
        
    }

    return list;
}

/**
 * wmi_ntdomain() 
 * 同
 * wmic ntdomain 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * DnsForestName 
 * DomainControllerAddress
 * DomainControllerName
 * DomainName 
 * Status 
 * 
 */
function wmi_ntdomain() {
    var data = run_wmi("SELECT * FROM Win32_NTDomain WHERE DomainName IS NOT NULL ");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.DnsForestName           = item.DnsForestName;
        obj.DomainControllerAddress = item.DomainControllerAddress;
        obj.DomainControllerName    = item.DomainControllerName;
        obj.DomainName              = item.DomainName; 
        obj.Status                  = item.Status 
        
        list.push(obj);
        
    }

    return list;
}


/**
 * wmi_share() 本機共用分享 
 * 同
 * wmic share 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Description 
 * Name
 * Path
 * Status 
 * Type 
 * 
 */
function wmi_share() {
    var data = run_wmi("SELECT * FROM Win32_Share ");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.Description             = item.Description;
        obj.Name                    = item.Name;
        obj.Path                    = item.Path;
        obj.Status                  = item.Status;
        obj.Type                    = item.Type;

        list.push(obj);
        
    }

    return list;
}

/**
 * wmi_desktop() 本機螢幕保護 
 * 同
 * wmic desktop where(Name="%USERDOMAIN%\\%USERNAME%") 
 *    
 * Fields (+ is Array)
 * --------------------------------------------
 * Name
 * ScreenSaverActive
 * ScreenSaverExecutable
 * ScreenSaverSecure
 * SecreenSaverTimeout 
 * 
 * 用 registry 比較簡單
 * HKEY_CURRENT_USER\Control Panel\Desktop 
 * --------------------------------------------
 * ScreenSaveActive      -> ScreenSaverActive 
 * ScreenSaverIsSecure   -> ScreenSaverSecure 
 * ScreenSaveTimeout     -> ScreenSaverTimeout 
 * SCRNSAVE.EXE          -> ScreenSaverExecutable 
 *
 */
function wmi_desktop() {
    var  path = "HKCU\\Control Panel\\Desktop\\";
    var  obj  = new Object();
    var  val  = "";

    val = reg_get(path + "ScreenSaveActive");
    obj.ScreenSaverActive           = (val == 1) ? true : false;

    val = reg_get(path + "ScreenSaverIsSecure");
    obj.ScreenSaverSecure           = (val == 1) ? true : false;

    val = reg_get(path + "ScreenSaveTimeout");
    obj.ScreenSaverTimeout          = parseInt(val, 10);
    
    val = reg_get(path + "SCRNSAVE.EXE");
    obj.ScreenSaverExecutable       = val;
    
    return obj;
}


/**
 * wmi_service() 開機服務 
 * 同
 * wmic service where (StartMode="Auto")
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Caption
 * DesktopInteract
 * DisplayName
 * Name
 * PathName
 * ServiceType
 * StartMode
 * Started 
 * StartName 
 * State 
 * Status
 * 
 */
function wmi_service() {
    var data = run_wmi("SELECT * FROM Win32_Service WHERE (StartMode=\"Auto\")");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.Caption                 = item.Caption;
        obj.DesktopInteract         = item.DesktopInteract;
        obj.DisplayName             = item.DisplayName;
        obj.Name                    = item.Name;
        obj.PathName                = item.PathName;
        obj.ServiceType             = item.ServiceType;
        obj.Started                 = item.Started;
        obj.StartMode               = item.StartMode;
        obj.StartName               = item.StartName;
        obj.State                   = item.State;
        obj.Status                  = item.Status;

        list.push(obj);
        
    }

    return list;
}


/**
 * wmi_product() 安裝軟體 
 * 同
 * wmic product 
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Caption
 * Description 
 * InstallDate
 * InstallSource
 * Name
 * Vendor
 * Version
 * 
 */
function wmi_product() {
    var data = run_wmi("SELECT * FROM Win32_Product");

    if (data.atEnd()) {
        return null;
    }

    var list = new Array();

    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();
        var  obj  = new Object();

        obj.Caption                 = item.Caption;
        obj.Description             = item.Description;
        obj.InstallDate             = item.InstallDate;
        obj.InstallSource           = item.InstallSource; 
        obj.InstallState            = item.InstallState; 
        obj.Name                    = item.Name;
        obj.Vendor                  = item.Vendor;
        obj.Version                 = item.Version;
        
        list.push(obj);
        
    }

    return list;
}

/**
 * wmi_quickfix() 系統更新 
 * 同
 * wmic qfe 
 * 使用 ADOR 作為排序
 * 
 * Fields (+ is Array)
 * --------------------------------------------
 * Caption                    Information URL
 * Description                Update | Security Update 
 * HotFixID                   KBnnnnnnn
 * InstalledBy
 * InstalledOn                M/D/YYYY  
 * 
 */
function wmi_quickfix(cnt) {
    var data = run_wmi("SELECT * FROM Win32_QuickFixEngineering ");
    var dao  = WScript.CreateObject("ADOR.Recordset");

    if (data.atEnd()) {
        return null;
    }

    dao.Fields.Append ("HotFixID"   , 200, 50);    // 字串 adVarChar = 200 
    dao.Fields.Append ("Caption"    , 200, 200);
    dao.Fields.Append ("Description", 200, 200);
    dao.Fields.Append ("InstalledBy", 200, 200);
    dao.Fields.Append ("InstalledOn", 7);   // 日期 adDate = 7 

    dao.Open();
    
    for( ; !data.atEnd() ; data.moveNext()) {
        var  item = data.item();

        dao.AddNew();
        
        dao("HotFixID")             = item.HotFixID;
        dao("Caption")              = item.Caption;
        dao("Description")          = item.Description;
        dao("InstalledBy")          = item.InstalledBy;
        dao("InstalledOn")          = item.InstalledOn.toDate("MDY");
        
        dao.Update();
    }

    var list = new Array();

    dao.Sort = "InstalledOn DESC";

    dao.MoveFirst();

    for(var i = 0 ; !dao.EOF && i < cnt; i++) {
        var obj = new Object();

        obj.HotFixID        = dao("HotFixID").Value;
        obj.Caption         = dao("Caption").Value;
        obj.Description     = dao("Description").Value;
        obj.InstalledBy     = dao("InstalledBy").Value;
        obj.InstalledOn     = new Date(dao("InstalledOn").Value);

        list[i] = obj;
    
        dao.MoveNext();
    }
    
    dao.Close();

    return list;
}

/**
 * net localgroup $group 
 *
 * Stdout 的前 6 行，及最後 2 行
 * 的內容為 garbage 
 * 
 */
function net_local_group(group) {
    var out = cmd_str("net localgroup \"" + group + "\"");

    if (!out) {
        // ERROR 
        return null;
    }

    var data = new Array();

    for(var i = 0; !out.AtEndOfStream; i++) {
        var s = out.ReadLine();
        if (i >=6) {
            data.push(s);  
        }  
    }

    data.pop();
    data.pop();

    return data;

}

/**
 *
 */
function net_listen() {
    var     out = cmd_str("cmd.exe /c netstat -na -p tcp | findstr LISTEN");

    if (!out) {
        // ERROR 
        return null;
    }

    var data = new Array();

    for(var i = 0; !out.AtEndOfStream; i++) {
        var s    = out.ReadLine();
        //   TCP    0.0.0.0:135            0.0.0.0:0              LISTENING
        var svc  = s.substring(9,32).split(":");
        var obj  = new Object();

        obj.Protocol = s.substring(2,5);
        obj.IP       = svc[0];
        obj.Port     = parseInt(svc[1], 10);

        data.push(obj);  
          
    }

    return data;
}


/**
 * offscan_schedule() OfficeScan 預約掃描及結果
 * 
 * HKLM\Software\TrendMicro\PC-cillinNTCorp\CurrentVersion
 *     \Prescheduled Scan Configuration
 * --------------------------------------------
 * Enable                -> 1 = Enable 
 * Frequency             -> 2 = Weekly 
 *                          3 = Daily 
 *                          4 = Hourly
 *                          5 = Once 
 * Hour                  -> 0 ~ 12 
 * DayOfWeek             -> 1 ~ 7 
 * AmPm                  -> 1 = AM, 2 = PM 
 * ScanStartTime         -> Unix/Epoch 
 * NextScanTime          -> Unix/Epoch 
 * LastScanStartTime     -> Unix/Epoch 
 * LastScanStatus        -> 0 = Completed
 *                          1 = Interrupted 
 *                          2 = Stopped Unexpectedly
 * LastScanFinishTime    -> Unix/Epoch 
 * 
 * ScheduleScanStatus 
 * LaunchScheduleScanWhenBoot 
 * ScanNetwork
 * ExcludedFolder 
 * ExcludedFile 
 * ExcludedExt 
 * ScanAllFiles 
 * VirusFundAction 
 * 
 */
function offscan_schedule() {
    var  path = "HKLM\\Software\\TrendMicro\\PC-cillinNTCorp"
              + "\\CurrentVersion\\Prescheduled Scan Configuration\\"
              ;
    var  obj  = new Object();
    var  val  = "";

    val = reg_get(path + "Enable");
    obj.Enable           = (val == 1) ? true : false;

    val = reg_get(path + "ScheduleScanStatus");
    obj.ScheduleScanStatus = (val == 1) ? true : false;

    val = reg_get(path + "ScanStartTime");
    obj.ScanStartTime = Date.epoch(val);
    
    val = reg_get(path + "LastScanStartTime");
    obj.LastScanStartTime = Date.epoch(val);
    
    val = reg_get(path + "LastScanFinishTime");
    obj.LastScanFinishTime = Date.epoch(val);
    
    val = reg_get(path + "LastScanStatus");
    obj.LastScanStatus = val == 0 ? "Completed"   :
                         val == 1 ? "Interrupted" :
                         val == 2 ? "Stopped Unexpectedly" :
                         val.toString()
                       ;
    
    val = reg_get(path + "NextScanTime");
    obj.NextScanTime = Date.epoch(val);
    
    val = reg_get(path + "ScanAllFiles");
    obj.ScanAllFiles           = (val == 1) ? true : false;

    val = reg_get(path + "VirusFoundAction");
    obj.VirusFoundAction = val;

    val = reg_get(path + "ExcludedFolder");
    obj.ExcludedFolder = val;

    val = reg_get(path + "ExcludedFile");
    obj.ExcludedFile = val;

    val = reg_get(path + "ExcludedExt");
    obj.ExcludedExt = val;

    val = reg_get(path + "Frequency");
    obj.Frequency   = val == 2 ? "Weekly" :
                      val == 3 ? "Daily"  :
                      val == 4 ? "Hourly" :
                      val == 5 ? "Once"   :
                      val.toString()
                      ;

    val = reg_get(path + "DayOfWeek");
    obj.DayOfWeek = val;

    h   = reg_get(path + "Hour");
    pm  = reg_get(path + "AmPm");
    m   = reg_get(path + "Minute"); 
    h   += (pm == 2 ? 12 : 0);

    hh = ((h < 10) ? "0" : "")
       + h.toString()
       ;
    
    mm = ((m < 10) ? "0" : "")
       + m.toString()
       ;
    
    obj.ScheduleTime = hh + ":" + mm;
     
    return obj;
}

// MAIN 
WScript.Echo("Hello");

var x = offscan_schedule();
WSH.Echo("Enable             -> " + x.Enable);
WSH.Echo("ScheduleScanStatus -> " + x.ScheduleScanStatus);
WSH.Echo("ScanScanTime       -> " + x.ScanStartTime);
WSH.Echo("LastScanScanTime   -> " + x.LastScanStartTime);
WSH.Echo("LastScanStatus     -> " + x.LastScanStatus);
WSH.Echo("LastScanFinishTime -> " + x.LastScanFinishTime);
WSH.Echo("NextScanTime       -> " + x.NextScanTime);

WSH.Echo("ScanAllFiles       -> " + x.ScanAllFiles);
WSH.Echo("VirusFoundAction   -> " + x.VirusFoundAction);
WSH.Echo("ExcludedFolder     -> " + x.ExcludedFolder);
WSH.Echo("ExcludedFile       -> " + x.ExcludedFile);
WSH.Echo("ExcludedExt        -> " + x.ExcludedExt);

WSH.Echo("Frequency          -> " + x.Frequency);
WSH.Echo("DayOfWeek          -> " + x.DayOfWeek);
WSH.Echo("ScheduleTime       -> " + x.ScheduleTime);
