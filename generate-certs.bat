@echo off
echo Generating self-signed certificate for development...

set CERT_DIR=.\certs
set KEY_FILE=%CERT_DIR%\server.key
set CERT_FILE=%CERT_DIR%\server.crt

REM Create certs directory if it doesn't exist
if not exist "%CERT_DIR%" mkdir "%CERT_DIR%"

REM Check if OpenSSL is available
where openssl >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Using OpenSSL to generate certificates...
    REM Generate private key
    openssl genrsa -out "%KEY_FILE%" 2048

    REM Generate certificate (valid for 365 days)
    openssl req -new -x509 -key "%KEY_FILE%" -out "%CERT_FILE%" -days 365 -subj "/C=JP/ST=Tokyo/L=Tokyo/O=YourCompany/CN=localhost"
    
    echo Certificate generated successfully using OpenSSL!
    goto :success
)

REM Try PowerShell method if OpenSSL is not available
echo OpenSSL not found, trying PowerShell method...
powershell -ExecutionPolicy Bypass -Command ^
"$cert = New-SelfSignedCertificate -DnsName 'localhost', '127.0.0.1' -CertStoreLocation 'cert:\CurrentUser\My' -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(1); " ^
"$certPassword = ConvertTo-SecureString -String 'password' -Force -AsPlainText; " ^
"$pfxPath = '%CERT_DIR%\server.pfx'; " ^
"Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $certPassword; " ^
"$pfx = Get-PfxCertificate -FilePath $pfxPath; " ^
"$certPem = '-----BEGIN CERTIFICATE-----' + [System.Convert]::ToBase64String($pfx.RawData, 'InsertLineBreaks') + '-----END CERTIFICATE-----'; " ^
"$certPem | Out-File -FilePath '%CERT_FILE%' -Encoding ASCII; " ^
"$rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($pfx); " ^
"$keyBytes = $rsa.ExportRSAPrivateKey(); " ^
"$keyPem = '-----BEGIN RSA PRIVATE KEY-----' + [System.Convert]::ToBase64String($keyBytes, 'InsertLineBreaks') + '-----END RSA PRIVATE KEY-----'; " ^
"$keyPem | Out-File -FilePath '%KEY_FILE%' -Encoding ASCII; " ^
"Remove-Item $pfxPath; " ^
"Remove-Item cert:\CurrentUser\My\$($cert.Thumbprint)"

if %ERRORLEVEL% EQU 0 (
    echo Certificate generated successfully using PowerShell!
    goto :success
)

echo Failed to generate certificates. Please try one of these alternatives:
echo.
echo Option 1 - Install mkcert (Recommended):
echo 1. Download mkcert from: https://github.com/FiloSottile/mkcert/releases
echo 2. Add mkcert.exe to your PATH
echo 3. Run: mkcert -install
echo 4. Run: mkcert -key-file certs\server.key -cert-file certs\server.crt localhost 127.0.0.1
echo.
echo Option 2 - Install OpenSSL:
echo 1. Download from: https://slproweb.com/products/Win32OpenSSL.html
echo 2. Add OpenSSL to your PATH
echo 3. Run this script again
echo.
echo Option 3 - Use HTTP instead of HTTPS (Less secure):
echo The development server is currently configured to use HTTP.
pause
exit /b 1

:success
echo Key file: %KEY_FILE%
echo Certificate file: %CERT_FILE%
echo.
echo Note: You may need to trust this certificate in your browser.
echo For Chrome: Go to Settings ^> Privacy and security ^> Security ^> Manage certificates
echo.
pause
