# Running UFT Tests via Command Line
A common question we get is how to exceute UFT tests via a Windows Command Line. This script was created to provide a simple, reusable script that covers most scenarios.

# Usage

Open a Windows Command Prompt (or Powershell) and execute this command:

`cscript <path_to_vbs_file>\run_uft_tests.vbs *directory-of-tests-to-run* [-flags]`

## Flags

`-w url-for-web-application` --If the test requires a URL for a web-based application, you can pass in the URL to start the test with.

`-r directory-path-for-results` -- Specifies the directory where test result files should be placed.

`-b browser-to-use` -- Specify the Browser the test should execute with.

`-e {True|False}` -- Fail on Error - Defaults to false, this tells the script to stop if a test found in the directory fails.

`-f {True|False}` -- Fail on Warning - Defaults to false, this tells the script to stop if a test found in the directory generates a warning.
