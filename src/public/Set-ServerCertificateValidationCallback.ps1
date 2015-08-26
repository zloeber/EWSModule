function Set-ServerCertificateValidationCallback ($CertCallback) {
    $script:modCertCallback = $CertCallback
}