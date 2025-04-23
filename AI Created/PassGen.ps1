<#
.SYNOPSIS
Provides a GUI to generate secure random passwords with enhanced options.

.DESCRIPTION
Launches a Windows Forms GUI for generating secure random passwords.
Allows specifying length and character types (uppercase, lowercase, numbers, special).
Includes options to exclude ambiguous characters and use custom special characters.
Provides copy-to-clipboard functionality with feedback.

.NOTES
Author: Gemini
Requires Windows PowerShell/PowerShell 7 on Windows.
Uses WinForms and secure RNG. Compatible with PS v2+.
#>

# --- Load Required Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Password Generation Logic (Function) ---
function Generate-SecurePassword {
    param(
        [Parameter(Mandatory=$true)] [int]$Length,
        [Parameter(Mandatory=$true)] [bool]$IncludeLowercase,
        [Parameter(Mandatory=$true)] [bool]$IncludeUppercase,
        [Parameter(Mandatory=$true)] [bool]$IncludeNumbers,
        [Parameter(Mandatory=$true)] [bool]$IncludeSpecialChars,
        [Parameter(Mandatory=$false)] [bool]$ExcludeAmbiguous = $false,
        [Parameter(Mandatory=$false)] [string]$CustomSpecialChars = $null
    )

    # Base Character Sets
    $lowercaseCharsBase = 'abcdefghijklmnopqrstuvwxyz'
    $uppercaseCharsBase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $numberCharsBase = '0123456789'
    # Use custom special chars if provided, otherwise default
    $specialCharsBase = if (-not [string]::IsNullOrEmpty($CustomSpecialChars)) { $CustomSpecialChars } else { '!@#$%^&*()-_=+[{]};:''",.<>/?`~' }

    # Characters to potentially exclude
    $ambiguousCharsPattern = '[l1IO0]' # Regex pattern for ambiguous chars

    # Adjust character sets based on ExcludeAmbiguous flag
    $lowercaseChars = $lowercaseCharsBase
    $uppercaseChars = $uppercaseCharsBase
    $numberChars = $numberCharsBase
    $specialChars = $specialCharsBase

    if ($ExcludeAmbiguous) {
        $lowercaseChars = $lowercaseChars -replace $ambiguousCharsPattern
        $uppercaseChars = $uppercaseChars -replace $ambiguousCharsPattern
        $numberChars = $numberChars -replace $ambiguousCharsPattern
    }

    # Build Character Pool and Required Characters List
    $characterPool = [System.Text.StringBuilder]::new()
    $requiredChars = [System.Collections.Generic.List[char]]::new()

    # Add characters to pool and ensure one of each required type exists
    if ($IncludeLowercase -and $lowercaseChars.Length -gt 0) {
        [void]$characterPool.Append($lowercaseChars)
        $requiredChars.Add(($lowercaseChars.ToCharArray() | Get-Random -Count 1))
    }
    if ($IncludeUppercase -and $uppercaseChars.Length -gt 0) {
        [void]$characterPool.Append($uppercaseChars)
        $requiredChars.Add(($uppercaseChars.ToCharArray() | Get-Random -Count 1))
    }
    if ($IncludeNumbers -and $numberChars.Length -gt 0) {
        [void]$characterPool.Append($numberChars)
        $requiredChars.Add(($numberChars.ToCharArray() | Get-Random -Count 1))
    }
    if ($IncludeSpecialChars -and $specialChars.Length -gt 0) {
        [void]$characterPool.Append($specialChars)
        $requiredChars.Add(($specialChars.ToCharArray() | Get-Random -Count 1))
    }

    # --- Input Validation ---
    if ($characterPool.Length -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No characters available to generate password based on current selections (check exclusions and custom characters).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
    if ($Length -lt $requiredChars.Count) {
         [System.Windows.Forms.MessageBox]::Show("Password length ($Length) must be at least $($requiredChars.Count) to include one of each selected character type.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
         return $null
    }

    # --- Password Generation ---
    $poolArray = $characterPool.ToString().ToCharArray()
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $passwordChars = [char[]]::new($Length)
    $remainingLength = $Length - $requiredChars.Count

    # Generate the random part
    if ($remainingLength -gt 0) {
        $randomBytes = [byte[]]::new($remainingLength * 4)
        $rng.GetBytes($randomBytes)
        for ($i = 0; $i -lt $remainingLength; $i++) {
            $uint32Value = [System.BitConverter]::ToUInt32($randomBytes, $i * 4)
            $randomIndex = $uint32Value % $poolArray.Length
            $passwordChars[$i] = $poolArray[$randomIndex]
        }
    }

    # Add required characters
    for ($i = 0; $i -lt $requiredChars.Count; $i++) {
        $passwordChars[$remainingLength + $i] = $requiredChars[$i]
    }

    # Shuffle the array (compatible with older PS versions)
    $shuffledPasswordChars = $passwordChars | Sort-Object -Property { [guid]::NewGuid() }

    # Clean up RNG
    if ($rng -is [System.IDisposable]) {
        $rng.Dispose()
    }

    # Return the final password string
    return (-join $shuffledPasswordChars)
}

# --- Build the Form (GUI) ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Secure Password Generator"
$form.Size = New-Object System.Drawing.Size(440, 380) # Reduced height
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- GUI Controls ---

# Row 1: Length
$labelLength = New-Object System.Windows.Forms.Label
$labelLength.Location = New-Object System.Drawing.Point(20, 25)
$labelLength.Size = New-Object System.Drawing.Size(120, 20)
$labelLength.Text = "Password Length:"
$form.Controls.Add($labelLength)

$numericUpDownLength = New-Object System.Windows.Forms.NumericUpDown
$numericUpDownLength.Location = New-Object System.Drawing.Point(150, 23)
$numericUpDownLength.Size = New-Object System.Drawing.Size(60, 25)
$numericUpDownLength.Minimum = 8
$numericUpDownLength.Maximum = 128
$numericUpDownLength.Value = 16
$numericUpDownLength.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$form.Controls.Add($numericUpDownLength)

# Row 2: Character Types GroupBox
$groupBoxTypes = New-Object System.Windows.Forms.GroupBox
$groupBoxTypes.Location = New-Object System.Drawing.Point(20, 60)
$groupBoxTypes.Size = New-Object System.Drawing.Size(380, 90) # Wider
$groupBoxTypes.Text = "Include Character Types"
$form.Controls.Add($groupBoxTypes)

$checkBoxLowercase = New-Object System.Windows.Forms.CheckBox
$checkBoxLowercase.Location = New-Object System.Drawing.Point(15, 25)
$checkBoxLowercase.AutoSize = $true
$checkBoxLowercase.Text = "Lowercase (a-z)"
$checkBoxLowercase.Checked = $true
$groupBoxTypes.Controls.Add($checkBoxLowercase)

$checkBoxUppercase = New-Object System.Windows.Forms.CheckBox
$checkBoxUppercase.Location = New-Object System.Drawing.Point(190, 25) # Adjusted position
$checkBoxUppercase.AutoSize = $true
$checkBoxUppercase.Text = "Uppercase (A-Z)"
$checkBoxUppercase.Checked = $true
$groupBoxTypes.Controls.Add($checkBoxUppercase)

$checkBoxNumbers = New-Object System.Windows.Forms.CheckBox
$checkBoxNumbers.Location = New-Object System.Drawing.Point(15, 55)
$checkBoxNumbers.AutoSize = $true
$checkBoxNumbers.Text = "Numbers (0-9)"
$checkBoxNumbers.Checked = $true
$groupBoxTypes.Controls.Add($checkBoxNumbers)

$checkBoxSpecial = New-Object System.Windows.Forms.CheckBox
$checkBoxSpecial.Location = New-Object System.Drawing.Point(190, 55) # Adjusted position
$checkBoxSpecial.AutoSize = $true
$checkBoxSpecial.Text = "Special (!@#$...)" # Shortened text
$checkBoxSpecial.Checked = $true
$groupBoxTypes.Controls.Add($checkBoxSpecial)

# Row 3: Additional Options
$checkBoxExcludeAmbiguous = New-Object System.Windows.Forms.CheckBox
$checkBoxExcludeAmbiguous.Location = New-Object System.Drawing.Point(25, 160) # New row
$checkBoxExcludeAmbiguous.AutoSize = $true
$checkBoxExcludeAmbiguous.Text = "Exclude Ambiguous Characters (l, 1, I, O, 0)"
$checkBoxExcludeAmbiguous.Checked = $false # Default off
$form.Controls.Add($checkBoxExcludeAmbiguous)

$labelCustomSpecial = New-Object System.Windows.Forms.Label
$labelCustomSpecial.Location = New-Object System.Drawing.Point(20, 195) # New row
$labelCustomSpecial.Size = New-Object System.Drawing.Size(380, 20)
$labelCustomSpecial.Text = "Custom Special Characters (overrides default if not empty):"
$form.Controls.Add($labelCustomSpecial)

$textBoxCustomSpecial = New-Object System.Windows.Forms.TextBox
$textBoxCustomSpecial.Location = New-Object System.Drawing.Point(25, 215)
$textBoxCustomSpecial.Size = New-Object System.Drawing.Size(375, 25)
$textBoxCustomSpecial.Font = New-Object System.Drawing.Font("Consolas", 9)
$textBoxCustomSpecial.Text = "" # Empty by default
$form.Controls.Add($textBoxCustomSpecial)


# Row 4: Generate Button
$buttonGenerate = New-Object System.Windows.Forms.Button
$buttonGenerate.Location = New-Object System.Drawing.Point(20, 255)
$buttonGenerate.Size = New-Object System.Drawing.Size(380, 30) # Wider
$buttonGenerate.Text = "Generate Password"

# Row 5: Password Display & Copy Button
$textBoxPassword = New-Object System.Windows.Forms.TextBox
$textBoxPassword.Location = New-Object System.Drawing.Point(20, 300)
$textBoxPassword.Size = New-Object System.Drawing.Size(295, 25) # Adjusted size
$textBoxPassword.ReadOnly = $true
$textBoxPassword.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxPassword)

$buttonCopy = New-Object System.Windows.Forms.Button
$buttonCopy.Location = New-Object System.Drawing.Point(325, 298)
$buttonCopy.Size = New-Object System.Drawing.Size(75, 27)
$buttonCopy.Text = "Copy"
$buttonCopy.Enabled = $false # Disabled initially

# --- Event Handlers ---

# Generate Button Click Event
$buttonGenerate.Add_Click({
    # Get values from controls
    $length = $numericUpDownLength.Value
    $includeLower = $checkBoxLowercase.Checked
    $includeUpper = $checkBoxUppercase.Checked
    $includeNumbers = $checkBoxNumbers.Checked
    $includeSpecial = $checkBoxSpecial.Checked
    $excludeAmbiguous = $checkBoxExcludeAmbiguous.Checked
    $customSpecial = $textBoxCustomSpecial.Text

    # Call the generation function
    $generatedPassword = Generate-SecurePassword -Length $length `
        -IncludeLowercase $includeLower `
        -IncludeUppercase $includeUpper `
        -IncludeNumbers $includeNumbers `
        -IncludeSpecialChars $includeSpecial `
        -ExcludeAmbiguous $excludeAmbiguous `
        -CustomSpecialChars $customSpecial

    # Display the password if generated successfully
    if ($generatedPassword -ne $null) {
        $textBoxPassword.Text = $generatedPassword
        $buttonCopy.Enabled = $true
    } else {
        # Clear fields if generation failed
        $textBoxPassword.Text = ""
        $buttonCopy.Enabled = $false
    }
})
$form.Controls.Add($buttonGenerate)
$form.AcceptButton = $buttonGenerate # Allow Enter key to trigger generate

# Copy Button Click Event
$buttonCopy.Add_Click({
    if (-not [string]::IsNullOrEmpty($textBoxPassword.Text)) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($textBoxPassword.Text)

            # Provide visual feedback
            $originalText = $buttonCopy.Text
            $buttonCopy.Text = "Copied!"
            $buttonCopy.Enabled = $false
            $form.Refresh() # Update UI
            Start-Sleep -Milliseconds 750 # Pause briefly
            # Check if form still exists before updating UI back
            if ($form -and -not $form.IsDisposed) {
                $buttonCopy.Text = $originalText
                $buttonCopy.Enabled = $true
                $form.Refresh()
            }

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not copy password to clipboard. Error: $($_.Exception.Message)", "Clipboard Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
})
$form.Controls.Add($buttonCopy)


# --- Show the Form ---
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog() # Show modally

# --- Clean up Form Resources ---
# Check if form exists and is not disposed before disposing
if ($form -and -not $form.IsDisposed) {
    $form.Dispose()
}
