Function Set-EWSSSLIgnoreWorkaround {
      <#
    .SYNOPSIS
    Sets the module to ignore SSL checking.
    .DESCRIPTION
    Sets the module to ignore SSL checking.
     
    .EXAMPLE
    Set-EWSSSLIgnoreWorkaround        

    .NOTES
    Author: Zachary Loeber
    Site: http://www.the-little-things.net/
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
    #>
    if (-not $script:IsSSLWorkAroundInPlace) {
        $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
        $Compiler=$Provider.CreateCompiler()
        $Params=New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable=$False
        $Params.GenerateInMemory=$True
        $Params.IncludeDebugInformation=$False
        $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

        $TASource=@'
          namespace Local.ToolkitExtensions.Net.CertificatePolicy{
            public class TrustAll : System.Net.ICertificatePolicy {
              public TrustAll() { 
              }
              public bool CheckValidationResult(System.Net.ServicePoint sp,
                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                System.Net.WebRequest req, int problem) {
                return true;
              }
            }
          }
'@ 
        $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
        $TAAssembly=$TAResults.CompiledAssembly

        ## We now create an instance of the TrustAll and attach it to the ServicePointManager
        $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
        [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

        $script:IsSSLWorkAroundInPlace = $true
    }
}