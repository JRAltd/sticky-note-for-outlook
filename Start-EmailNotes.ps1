<#
.SYNOPSIS
  Generates a local HTTPS development certificate for localhost (if not already present)
  and starts the Outlook Email Sticky Notes Add-in server on https://localhost:3000.
#>
param(
  [int]$Port = 3000,
  [switch]$HttpOnly
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$CertsDir = Join-Path $ProjectRoot "certs"
$CertPath = Join-Path $CertsDir "localhost.crt"
$KeyPath  = Join-Path $CertsDir "localhost.key"

if (-not $HttpOnly -and (-not (Test-Path $CertPath) -or -not (Test-Path $KeyPath))) {
  Write-Host "Generating self-signed localhost SSL certificate for Outlook Add-in HTTPS..." -ForegroundColor Cyan
  New-Item -ItemType Directory -Force -Path $CertsDir | Out-Null

  $csCode = @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;

public static class PemExporter {
    public static void SaveCertAndKey(X509Certificate2 cert, string certPath, string keyPath) {
        var certSb = new StringBuilder();
        certSb.AppendLine("-----BEGIN CERTIFICATE-----");
        certSb.AppendLine(Convert.ToBase64String(cert.RawData, Base64FormattingOptions.InsertLineBreaks));
        certSb.AppendLine("-----END CERTIFICATE-----");
        File.WriteAllText(certPath, certSb.ToString(), Encoding.ASCII);

        RSA rsa = cert.GetRSAPrivateKey();
        RSAParameters p = rsa.ExportParameters(true);
        byte[] pkcs1 = EncodeRsaPrivateKey(p);
        var keySb = new StringBuilder();
        keySb.AppendLine("-----BEGIN RSA PRIVATE KEY-----");
        keySb.AppendLine(Convert.ToBase64String(pkcs1, Base64FormattingOptions.InsertLineBreaks));
        keySb.AppendLine("-----END RSA PRIVATE KEY-----");
        File.WriteAllText(keyPath, keySb.ToString(), Encoding.ASCII);
    }

    private static byte[] EncodeRsaPrivateKey(RSAParameters p) {
        var body = new List<byte>();
        body.AddRange(EncodeInteger(new byte[] { 0x00 }));
        body.AddRange(EncodeInteger(p.Modulus));
        body.AddRange(EncodeInteger(p.Exponent));
        body.AddRange(EncodeInteger(p.D));
        body.AddRange(EncodeInteger(p.P));
        body.AddRange(EncodeInteger(p.Q));
        body.AddRange(EncodeInteger(p.DP));
        body.AddRange(EncodeInteger(p.DQ));
        body.AddRange(EncodeInteger(p.InverseQ));
        return EncodeSequence(body.ToArray());
    }

    private static byte[] EncodeInteger(byte[] value) {
        int i = 0;
        while (i < value.Length - 1 && value[i] == 0) i++;
        bool needPad = (value[i] & 0x80) != 0;
        int len = value.Length - i + (needPad ? 1 : 0);
        var res = new List<byte>();
        res.Add(0x02);
        res.AddRange(EncodeLength(len));
        if (needPad) res.Add(0x00);
        for (int j = i; j < value.Length; j++) res.Add(value[j]);
        return res.ToArray();
    }

    private static byte[] EncodeSequence(byte[] content) {
        var res = new List<byte>();
        res.Add(0x30);
        res.AddRange(EncodeLength(content.Length));
        res.AddRange(content);
        return res.ToArray();
    }

    private static byte[] EncodeLength(int len) {
        if (len < 0x80) return new byte[] { (byte)len };
        var bytes = new List<byte>();
        int temp = len;
        while (temp > 0) {
            bytes.Insert(0, (byte)(temp & 0xff));
            temp >>= 8;
        }
        var res = new List<byte>();
        res.Add((byte)(0x80 | bytes.Count));
        res.AddRange(bytes);
        return res.ToArray();
    }
}
"@

  if (-not ("PemExporter" -as [type])) {
    Add-Type -TypeDefinition $csCode -Language CSharp
  }

  $cert = New-SelfSignedCertificate `
    -DnsName "localhost" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

  [PemExporter]::SaveCertAndKey($cert, $CertPath, $KeyPath)
  Write-Host "SSL certificate created at $CertPath" -ForegroundColor Green
}

$env:PORT = "$Port"
if ($HttpOnly) {
  python "$ProjectRoot\server.py" --http
} else {
  python "$ProjectRoot\server.py"
}
