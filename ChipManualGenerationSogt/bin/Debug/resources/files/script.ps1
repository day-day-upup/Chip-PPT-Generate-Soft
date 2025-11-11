# =================================================================
# 方案 1: 纯 PowerShell FTP 递归下载脚本
# 注意: FTP 列表解析 (Parse-FtpListing) 相对脆弱，可能因服务器格式不同而失败。
# =================================================================

# --- 配置区 ---
$FtpUri = "ftp://192.168.1.77/"
$RemoteFolder = "Amplifier/MML806/"  # 要下载的远程根目录
$LocalPath = "F:\FTP_Downloads\"       # 本地保存路径
$Username = "yp"
$Password = "123456"
# -----------

# 确保本地目标目录存在
if (-not (Test-Path $LocalPath)) {
    New-Item -Path $LocalPath -ItemType Directory | Out-Null
}

# 凭据对象
$Credentials = New-Object System.Net.NetworkCredential($Username, $Password)

# 递归下载主函数
function Download-FtpDirectory {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RemoteDir,
        [Parameter(Mandatory=$true)]
        [string]$LocalDir
    )

    Write-Host "--> 正在处理远程目录: $RemoteDir"
    
    # 确保本地目录存在
    if (-not (Test-Path $LocalDir)) {
        New-Item -Path $LocalDir -ItemType Directory | Out-Null
    }

    # --- 1. 获取目录列表 (使用 ListDirectoryDetails) ---
    $RemoteUri = $FtpUri + $RemoteDir.Trim('/') + "/"
    # 对 URI 进行编码以处理中文和空格
    $EncodedUri = [System.Uri]::EscapeUriString($RemoteUri)
    
    try {
        $ListRequest = [System.Net.FtpWebRequest]::Create($EncodedUri)
        $ListRequest.Credentials = $Credentials
        $ListRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails
        $ListRequest.UsePassive = $true # 强烈建议使用被动模式
        $ListRequest.Timeout = 10000 
        
        $ListResponse = $ListRequest.GetResponse()
        $Reader = New-Object System.IO.StreamReader $ListResponse.GetResponseStream()
        $Listing = ($Reader.ReadToEnd() -split "`r`n")
        $Reader.Close()
        $ListResponse.Close()
    }
    catch {
        Write-Error "无法获取目录列表 [$RemoteDir]: $($_.Exception.Message)"
        return
    }

    # --- 2. 遍历列表 ---
    foreach ($Line in $Listing) {
        if ([string]::IsNullOrWhiteSpace($Line)) { continue }
        
        # ⚠️ 简易且不健壮的解析 (Unix 格式)
        $Parts = $Line -split '\s+'
        $IsDirectory = $Parts[0].StartsWith("d")
        # 尝试获取文件名 (通常是列表的最后一项)
        $Name = $Parts[$Parts.Length-1] 

        # 忽略当前目录 (.) 和父目录 (..)
        if ($Name -eq "." -or $Name -eq "..") { continue }

        $RemoteItem = $RemoteDir.Trim('/') + "/" + $Name
        $LocalItem = Join-Path $LocalDir $Name

        if ($IsDirectory) {
            # 递归处理子目录
            Download-FtpDirectory -RemoteDir $RemoteItem -LocalDir $LocalItem
        }
        else {
            # --- 下载文件 ---
            $DownloadUri = [System.Uri]::EscapeUriString($FtpUri + $RemoteItem)
            
            try {
                $WebClient = New-Object System.Net.WebClient
                $WebClient.Credentials = $Credentials
                $WebClient.DownloadFile($DownloadUri, $LocalItem)
                Write-Host "下载成功: $RemoteItem -> $LocalItem"
            }
            catch {
                Write-Error "下载文件失败 [$RemoteItem]: $($_.Exception.Message)"
            }
        }
    }
}

# --- 运行主函数 ---
Download-FtpDirectory -RemoteDir $RemoteFolder -LocalDir $LocalPath