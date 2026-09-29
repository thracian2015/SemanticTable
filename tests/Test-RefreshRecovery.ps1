$ErrorActionPreference = 'Stop'
$source = Get-Content "$PSScriptRoot\..\src\SemanticTable\RefreshRecovery.cs" -Raw
Add-Type -TypeDefinition ($source + @"
public static class RecoveryTests
{
    public static void Run()
    {
        Check(new int[] { 1 }, "N", true);
        Check(new int[] { 0 }, "N", false);
        Check(new int[] { -1, 1, 1 }, "NRPWN", true);
        Check(new int[] { -1, 0 }, "NRP", false);
        Check(new int[] { -1, -1 }, "NRP", null);
        Check(new int[] { -1, 1, -1 }, "NRPWN", null);
        Check(new int[] { -1, 1, 0 }, "NRPWN", false);
        Check(new int[] { -2 }, "N", null);
    }
    private static void Check(int[] outcomes, string expected, bool? result)
    {
        int i = 0;
        string trace = "";
        bool previous = false;
        bool? actual = null;
        try
        {
            actual = SemanticTable.RefreshRecovery.Run(() => {
                trace += previous ? "P" : "N";
                int outcome = outcomes[i++];
                if (outcome < 0) throw new System.Runtime.InteropServices.COMException("test", outcome == -1 ? unchecked((int)0x800A03EC) : unchecked((int)0x8007000E));
                return outcome == 1;
            }, () => { trace += "R"; previous = true; },
               () => { trace += "W"; previous = false; }, message => {});
        }
        catch (System.Runtime.InteropServices.COMException) { }
        if (trace != expected || actual != result || i != outcomes.Length)
            throw new System.Exception("Unexpected recovery behavior: " + trace);
    }
}
"@)
[RecoveryTests]::Run()
Write-Output 'PASS: success, cancellation, recovery, bounded retry, and non-recoverable error (8 scenarios).'
