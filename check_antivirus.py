import subprocess

# Run PowerShell command to check antivirus status
command = 'powershell "Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object displayName, productState"'
result = subprocess.run(command, shell=True, capture_output=True, text=True)

print("Antivirus Check Result:")
print(result.stdout)

if "displayName" in result.stdout:
    print("✅ Antivirus detected")
else:
    print("❌ No antivirus detected")
