function OpenPowerShell() {
    // Define the path to your PowerShell script
    var psScriptPath = "C:\\Deekshith\\Signn.ps1";

    try {
        // Create a new Process object for PowerShell
        var process = Sys.OleObject("WScript.Shell").Exec("powershell.exe -File " + psScriptPath);

        // Wait for the PowerShell process to finish
        while (process.Status == 0) {
            Delay(100); // Wait for 100 milliseconds
        }

        // Check the exit code
        var exitCode = process.ExitCode;
        if (exitCode === 0) {
            // Read and log the PowerShell script output line by line
            var output = process.StdOut.ReadAll();
            var outputLines = output.split("\n");
            for (var i = 0; i < outputLines.length; i++) {
                Log.Message("PowerShell out: " + outputLines[i].trim());
            }
        } else {
            Log.Error("PowerShell script execution failed with exit code: " + exitCode);

            // Read and log the PowerShell script error output
            var errorOutput = process.StdErr.ReadAll();
            Log.Error("PowerShell Error Output: " + errorOutput);
        }
    } catch (e) {
        Log.Error("Error executing PowerShell script: " + e.message);
    }
}
