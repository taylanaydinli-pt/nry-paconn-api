# PowerShell script to deploy the Python Flask application on Windows Server
param (
    [Parameter(Mandatory=$false)]
    [string] $AppInsightsKey = "",

    [Parameter(Mandatory=$false)]
    [string] $AppInsightsConnectionString = ""
)

# Configuration
$AppDir = "C:\nry-paconn-api"
$LogDir = "$AppDir\logs"
$PythonVersion = "3.11.8"
$PythonUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$PythonInstaller = "$env:TEMP\python-installer.exe"
$NssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
$NssmZip = "$env:TEMP\nssm.zip"
$NssmDir = "$AppDir\nssm"
$ServiceName = "NryPaconnApi"
$AppPort = 5000

Write-Output "Starting deployment of API application on Windows Server..."

# Create directories
Write-Output "Creating application directories..."
if (-Not (Test-Path $AppDir)) {
    New-Item -ItemType Directory -Path $AppDir -Force
}
if (-Not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force
}

# Download and install Python if not already installed
if (-Not (Test-Path "C:\Program Files\Python311\python.exe")) {
    Write-Output "Downloading Python installer..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $PythonUrl -OutFile $PythonInstaller

    Write-Output "Installing Python..."
    Start-Process -FilePath $PythonInstaller -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait

    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
}

# Download and extract NSSM (Non-Sucking Service Manager)
if (-Not (Test-Path $NssmDir)) {
    Write-Output "Downloading NSSM..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $NssmUrl -OutFile $NssmZip

    Write-Output "Extracting NSSM..."
    Expand-Archive -Path $NssmZip -DestinationPath $env:TEMP -Force
    New-Item -ItemType Directory -Path $NssmDir -Force
    Copy-Item -Path "$env:TEMP\nssm-2.24\win64\*" -Destination $NssmDir -Recurse -Force
}

# Clone or copy application files
Write-Output "Setting up application files..."
# For this script, we'll simulate by creating placeholder files

# Create a .env file
$EnvContent = @"
# Environment configuration for API app

# Application settings
APP_HOST=0.0.0.0
APP_PORT=$AppPort
APP_DEBUG=False

# Azure Application Insights
APPINSIGHTS_INSTRUMENTATIONKEY=$AppInsightsKey
APPINSIGHTS_CONNECTION_STRING=$AppInsightsConnectionString

# Azure Resource Group
AZURE_RESOURCE_GROUP=taylan-playground
"@

Set-Content -Path "$AppDir\.env" -Value $EnvContent

# Create a virtual environment
Write-Output "Creating Python virtual environment..."
$VenvCmd = "python -m venv $AppDir\venv"
Invoke-Expression $VenvCmd

# Install required packages
Write-Output "Installing Python packages..."
$InstallCmd = "$AppDir\venv\Scripts\pip.exe install -r requirements.txt"
Invoke-Expression $InstallCmd

# Configure as a Windows Service using NSSM
Write-Output "Configuring Windows Service using NSSM..."
$NssmExe = "$NssmDir\nssm.exe"
$ServiceExists = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

if ($ServiceExists) {
    & $NssmExe stop $ServiceName
    & $NssmExe remove $ServiceName confirm
}

& $NssmExe install $ServiceName "$AppDir\venv\Scripts\python.exe"
& $NssmExe set $ServiceName AppParameters "$AppDir\venv\Scripts\gunicorn.exe --config $AppDir\gunicorn.conf.py app.app:app"
& $NssmExe set $ServiceName AppDirectory "$AppDir"
& $NssmExe set $ServiceName DisplayName "Nry Paconn API Service"
& $NssmExe set $ServiceName Description "Long-running API endpoint service"
& $NssmExe set $ServiceName Start SERVICE_AUTO_START
& $NssmExe set $ServiceName AppStdout "$LogDir\service-stdout.log"
& $NssmExe set $ServiceName AppStderr "$LogDir\service-stderr.log"
& $NssmExe set $ServiceName AppRotateFiles 1
& $NssmExe set $ServiceName AppRotateBytes 10485760

# Install and configure IIS for reverse proxy
Write-Output "Installing IIS and required features..."
Install-WindowsFeature -Name Web-Server, Web-Mgmt-Tools, Web-WebSockets -IncludeManagementTools

# Install URL Rewrite Module for IIS
$UrlRewriteUrl = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"
$UrlRewriteMsi = "$env:TEMP\urlrewrite.msi"

if (-not (Get-Module -ListAvailable -Name IISAdministration)) {
    Write-Output "Downloading URL Rewrite module..."
    Invoke-WebRequest -Uri $UrlRewriteUrl -OutFile $UrlRewriteMsi

    Write-Output "Installing URL Rewrite module..."
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $UrlRewriteMsi /qn" -Wait
}

# Create web.config for URL Rewrite
$WebConfigContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <rewrite>
            <rules>
                <rule name="ReverseProxy" stopProcessing="true">
                    <match url="(.*)" />
                    <action type="Rewrite" url="http://localhost:$AppPort/{R:1}" />
                </rule>
            </rules>
        </rewrite>
        <handlers>
            <remove name="WebDAV" />
        </handlers>
        <modules>
            <remove name="WebDAVModule" />
        </modules>
    </system.webServer>
</configuration>
"@

# Create Default Web Site directory if it doesn't exist
$IisSitePath = "C:\inetpub\wwwroot"
if (-Not (Test-Path $IisSitePath)) {
    New-Item -ItemType Directory -Path $IisSitePath -Force
}

# Write web.config file
Set-Content -Path "$IisSitePath\web.config" -Value $WebConfigContent -Force

# Create basic gunicorn.conf.py
$GunicornConfigContent = @"
#!/usr/bin/env python3
"""
Gunicorn configuration for long-running request handling
"""

# Server socket
bind = "0.0.0.0:$AppPort"

# Worker processes
workers = 4
worker_class = "sync"
threads = 2

# Timeout configuration
timeout = 600  # 10 minutes in seconds for long-running requests
graceful_timeout = 30
keepalive = 5

# Logging
accesslog = "-"
errorlog = "-"
loglevel = "info"

# Server mechanics
daemon = False
reload = False
"@

Set-Content -Path "$AppDir\gunicorn.conf.py" -Value $GunicornConfigContent -Force

# Start the service
Write-Output "Starting the service..."
Start-Service -Name $ServiceName

# Ensure IIS is running
Write-Output "Ensuring IIS is running..."
Start-Service -Name W3SVC

Write-Output "Deployment completed successfully!"
Write-Output "API is available at http://[server-ip]/"
