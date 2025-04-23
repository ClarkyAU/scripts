<#
.SYNOPSIS
Provides a GUI to generate secure random passwords or passphrases with enhanced options.

.DESCRIPTION
Launches a Windows Forms GUI for generating secure random passwords or passphrases.
Password Mode: Allows specifying length, character types (lowercase/uppercase), exact number of digits and special chars, excluding ambiguous chars, and custom special chars.
Passphrase Mode: Downloads the EFF Large Wordlist (in the background) and allows specifying number of words, separator, capitalization, and inserting exact numbers of digits/symbols randomly before/after each word.
Provides copy-to-clipboard functionality with feedback.

.NOTES
Author: Gemini
Requires Windows PowerShell/PowerShell 7 on Windows with internet access for passphrase mode.
Uses WinForms and secure RNG. Compatible with PS v2+.
Downloads EFF Large Wordlist (approx 120KB) on first use of Passphrase mode per session.
Uses a Timer to check background download status while the form is open.
#>

# --- Load Required Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Global Variables ---
$Script:EffWordList = $null # To store the downloaded word list
$Script:WordListUrl = "https://www.eff.org/files/2016/07/18/eff_large_wordlist.txt" # URL for the word list
$Script:DefaultSpecialChars = '!@#$%^&*()-_=+[{]};:''",.<>/?`~'
$Script:PassphraseSymbols = '!@#$%^&*()-_=+?' # Smaller set for passphrase appending/inserting
$Script:DownloadJob = $null # Variable to hold the background job object

# --- Password Generation Logic (Function) ---
function Generate-SecurePassword {
    param(
        [Parameter(Mandatory=$true)] [int]$Length,
        [Parameter(Mandatory=$true)] [bool]$IncludeLowercase,
        [Parameter(Mandatory=$true)] [bool]$IncludeUppercase,
        [Parameter(Mandatory=$true)] [int]$NumberOfNumbers,
        [Parameter(Mandatory=$true)] [int]$NumberOfSpecialChars,
        [Parameter(Mandatory=$false)] [bool]$ExcludeAmbiguous = $false,
        [Parameter(Mandatory=$false)] [string]$CustomSpecialChars = $null
    )

    # Base Character Sets
    $lowercaseCharsBase = 'abcdefghijklmnopqrstuvwxyz'
    $uppercaseCharsBase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $numberCharsBase = '0123456789'
    $specialCharsBase = if (-not [string]::IsNullOrEmpty($CustomSpecialChars)) { $CustomSpecialChars } else { $Script:DefaultSpecialChars }
    $ambiguousCharsPattern = '[l1IO0]' # Regex for ambiguous chars

    # Adjust character sets based on ExcludeAmbiguous flag
    $lowercaseChars = $lowercaseCharsBase
    $uppercaseChars = $uppercaseCharsBase
    $numberChars = $numberCharsBase
    $specialChars = $specialCharsBase
    if ($ExcludeAmbiguous) {
        $lowercaseChars = $lowercaseChars -replace $ambiguousCharsPattern
        $uppercaseChars = $uppercaseChars -replace $ambiguousCharsPattern
        $numberChars = $numberChars -replace $ambiguousCharsPattern
        # Note: Ambiguous check not typically applied to special chars unless custom set contains them
    }

    # --- Validation ---
    # Check if character sets required for specified counts are available
    if ($NumberOfNumbers -gt 0 -and $numberChars.Length -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Cannot include numbers: Number character set is empty (possibly due to exclusions).", "Error", 'OK', 'Error'); return $null
    }
    if ($NumberOfSpecialChars -gt 0 -and $specialChars.Length -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Cannot include special characters: Special character set is empty (check custom set and exclusions).", "Error", 'OK', 'Error'); return $null
    }

    # Calculate minimum required length based on selections using if/else
    $minRequiredForLetters = 0
    if ($IncludeLowercase) { $minRequiredForLetters++ }
    if ($IncludeUppercase) { $minRequiredForLetters++ }

    $totalRequiredChars = $NumberOfNumbers + $NumberOfSpecialChars + $minRequiredForLetters

    # Check if length is sufficient
    if ($Length -lt $totalRequiredChars) {
         [System.Windows.Forms.MessageBox]::Show("Password length ($Length) is too short. It must be at least $totalRequiredChars to include the specified number of digits/symbols and at least one of each selected letter type.", "Error", 'OK', 'Error')
         return $null
    }
    # Check if at least one type is selected overall
    if (-not $IncludeLowercase -and -not $IncludeUppercase -and $NumberOfNumbers -eq 0 -and $NumberOfSpecialChars -eq 0) {
         [System.Windows.Forms.MessageBox]::Show("Please select at least one character type (lowercase, uppercase) or specify a number of digits/symbols.", "Error", 'OK', 'Error')
         return $null
    }


    # --- Generation ---
    $passwordChars = [System.Collections.Generic.List[char]]::new()
    $characterPool = [System.Text.StringBuilder]::new()

    # 1. Add required numbers
    if ($NumberOfNumbers -gt 0) {
        [void]$characterPool.Append($numberChars)
        for ($i = 0; $i -lt $NumberOfNumbers; $i++) {
            $passwordChars.Add(($numberChars.ToCharArray() | Get-Random -Count 1))
        }
    }
    # 2. Add required special characters
    if ($NumberOfSpecialChars -gt 0) {
        [void]$characterPool.Append($specialChars)
        for ($i = 0; $i -lt $NumberOfSpecialChars; $i++) {
            $passwordChars.Add(($specialChars.ToCharArray() | Get-Random -Count 1))
        }
    }
    # 3. Add required letters (at least one if selected)
    if ($IncludeLowercase -and $lowercaseChars.Length -gt 0) {
        [void]$characterPool.Append($lowercaseChars)
        $passwordChars.Add(($lowercaseChars.ToCharArray() | Get-Random -Count 1))
    }
     if ($IncludeUppercase -and $uppercaseChars.Length -gt 0) {
        [void]$characterPool.Append($uppercaseChars)
        $passwordChars.Add(($uppercaseChars.ToCharArray() | Get-Random -Count 1))
    }

    # 4. Fill remaining length with random chars from the pool
    $remainingLength = $Length - $passwordChars.Count
    if ($remainingLength -gt 0) {
        if ($characterPool.Length -eq 0) {
            # Fallback check: Rebuild pool if only numbers/symbols filled the exact length initially
            if ($IncludeLowercase -and $lowercaseChars.Length -gt 0) { [void]$characterPool.Append($lowercaseChars) }
            if ($IncludeUppercase -and $uppercaseChars.Length -gt 0) { [void]$characterPool.Append($uppercaseChars) }
            if ($characterPool.Length -eq 0) {
                 [System.Windows.Forms.MessageBox]::Show("Cannot fill remaining password length: Character pool is empty based on selections.", "Error", 'OK', 'Error'); return $null
            }
        }
        $poolArray = $characterPool.ToString().ToCharArray()
        $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
        $randomBytes = [byte[]]::new($remainingLength * 4); $rng.GetBytes($randomBytes)
        for ($i = 0; $i -lt $remainingLength; $i++) {
            $uint32Value = [System.BitConverter]::ToUInt32($randomBytes, $i * 4)
            $randomIndex = $uint32Value % $poolArray.Length
            $passwordChars.Add($poolArray[$randomIndex])
        }
         if ($rng -is [System.IDisposable]) { $rng.Dispose() }
    }

    # 5. Shuffle the final list
    $finalPasswordArray = $passwordChars.ToArray() | Sort-Object -Property { [guid]::NewGuid() }

    # Return password string
    return (-join $finalPasswordArray)
}

# --- Passphrase Generation Logic (Function) ---
function Generate-Passphrase {
    param(
        [Parameter(Mandatory=$true)] [int]$NumberOfWords,
        [Parameter(Mandatory=$true)] [string]$Separator,
        [Parameter(Mandatory=$true)] [bool]$Capitalize,
        [Parameter(Mandatory=$true)] [int]$NumberOfNumbers,
        [Parameter(Mandatory=$true)] [int]$NumberOfSymbols
    )

    # Check if word list is loaded
    if ($Script:EffWordList -eq $null -or $Script:EffWordList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Word list is not loaded. Cannot generate passphrase.", "Error", 'OK', 'Error')
        return $null
    }

    # Select random words from the loaded list
    $selectedWords = $Script:EffWordList | Get-Random -Count $NumberOfWords

    # Apply capitalization if requested
    if ($Capitalize) {
        $selectedWords = $selectedWords | ForEach-Object {
            if ($_.Length -gt 0) {
                $_.Substring(0, 1).ToUpper() + $_.Substring(1) # Capitalise first letter
            } else { $_ }
        }
    }

    # --- MODIFIED LOGIC for inserting numbers/symbols randomly before/after words ---

    # 1. Generate the list of extra characters (numbers and symbols)
    $extraChars = [System.Collections.Generic.List[char]]::new()
    if ($NumberOfNumbers -gt 0) {
        $numberChars = '0123456789'.ToCharArray()
        for ($i = 0; $i -lt $NumberOfNumbers; $i++) {
            $extraChars.Add(($numberChars | Get-Random -Count 1))
        }
    }
    if ($NumberOfSymbols -gt 0) {
        $symbolChars = $Script:PassphraseSymbols.ToCharArray()
        if ($symbolChars.Length -gt 0) {
             for ($i = 0; $i -lt $NumberOfSymbols; $i++) {
                $extraChars.Add(($symbolChars | Get-Random -Count 1))
            }
        } else {
            Write-Warning "Cannot add symbols to passphrase: Symbol set is empty."
        }
    }

    # 2. Shuffle the extra characters (if any exist)
    $shuffledExtraChars = $null
    if ($extraChars.Count -gt 0) {
        $shuffledExtraChars = $extraChars.ToArray() | Sort-Object -Property { [guid]::NewGuid() }
    }

    # --- CORRECTED QUEUE CREATION for older PowerShell ---
    # Create an empty queue using New-Object
    $extraCharsQueue = New-Object 'System.Collections.Generic.Queue[char]'
    # Enqueue items only if the shuffled array exists and has items
    if ($null -ne $shuffledExtraChars) {
        foreach ($char in $shuffledExtraChars) {
            $extraCharsQueue.Enqueue($char)
        }
    }
    # --- END CORRECTION ---


    # 3. Build the final passphrase by interleaving randomly before/after words
    $resultBuilder = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $selectedWords.Count; $i++) {
        $currentWord = $selectedWords[$i]
        $prefix = ""
        $suffix = ""

        # Try to add an extra character either before or after
        if ($extraCharsQueue.Count -gt 0) {
            $extraChar = $extraCharsQueue.Dequeue()
            # Randomly decide placement (0 = before, 1 = after)
            $placement = Get-Random -Minimum 0 -Maximum 2
            if ($placement -eq 0) {
                $prefix = $extraChar
            } else {
                $suffix = $extraChar
            }
        }

        # Append prefix, word, suffix
        [void]$resultBuilder.Append($prefix)
        [void]$resultBuilder.Append($currentWord)
        [void]$resultBuilder.Append($suffix)

        # Append the separator unless it's the last word
        if ($i -lt ($selectedWords.Count - 1)) {
            [void]$resultBuilder.Append($Separator)
        }
    }

    # 4. Append any remaining extra characters (if more extras than words)
    while ($extraCharsQueue.Count -gt 0) {
         [void]$resultBuilder.Append($extraCharsQueue.Dequeue())
    }

    # Return the final passphrase string
    return $resultBuilder.ToString()
}


# --- Build the Form (GUI) ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Secure Password & Passphrase Generator"
$form.Size = New-Object System.Drawing.Size(410, 420) # UPDATED FORM SIZE
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- GUI Controls ---

# Mode Selection Panel
$panelMode = New-Object System.Windows.Forms.Panel
$panelMode.Location = New-Object System.Drawing.Point(10, 10)
$panelMode.Size = New-Object System.Drawing.Size(440, 30)
$form.Controls.Add($panelMode)

$radioPasswordMode = New-Object System.Windows.Forms.RadioButton
$radioPasswordMode.Location = New-Object System.Drawing.Point(10, 5)
$radioPasswordMode.AutoSize = $true
$radioPasswordMode.Text = "Password Mode"
$radioPasswordMode.Checked = $true # Default mode
$panelMode.Controls.Add($radioPasswordMode)

$radioPassphraseMode = New-Object System.Windows.Forms.RadioButton
$radioPassphraseMode.Location = New-Object System.Drawing.Point(180, 5)
$radioPassphraseMode.AutoSize = $true
$radioPassphraseMode.Text = "Passphrase Mode"
$panelMode.Controls.Add($radioPassphraseMode)

# --- Password Options GroupBox ---
$groupBoxPasswordOpts = New-Object System.Windows.Forms.GroupBox
$groupBoxPasswordOpts.Location = New-Object System.Drawing.Point(10, 45)
$groupBoxPasswordOpts.Size = New-Object System.Drawing.Size(440, 265) # Increased size
$groupBoxPasswordOpts.Text = "Password Options"
$form.Controls.Add($groupBoxPasswordOpts)

# Length (inside Password GroupBox)
$labelLength = New-Object System.Windows.Forms.Label
$labelLength.Location = New-Object System.Drawing.Point(10, 25)
$labelLength.Size = New-Object System.Drawing.Size(120, 20)
$labelLength.Text = "Password Length:"
$groupBoxPasswordOpts.Controls.Add($labelLength)

$numericUpDownLength = New-Object System.Windows.Forms.NumericUpDown
$numericUpDownLength.Location = New-Object System.Drawing.Point(140, 23)
$numericUpDownLength.Size = New-Object System.Drawing.Size(60, 25)
$numericUpDownLength.Minimum = 8
$numericUpDownLength.Maximum = 128
$numericUpDownLength.Value = 16
$numericUpDownLength.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$groupBoxPasswordOpts.Controls.Add($numericUpDownLength)

# Letter Types GroupBox (Nested)
$groupBoxLetters = New-Object System.Windows.Forms.GroupBox
$groupBoxLetters.Location = New-Object System.Drawing.Point(10, 55)
$groupBoxLetters.Size = New-Object System.Drawing.Size(420, 60) # Adjusted size
$groupBoxLetters.Text = "Include Letters (at least one of each selected type)"
$groupBoxPasswordOpts.Controls.Add($groupBoxLetters)

$checkBoxLowercase = New-Object System.Windows.Forms.CheckBox
$checkBoxLowercase.Location = New-Object System.Drawing.Point(15, 25); $checkBoxLowercase.AutoSize = $true
$checkBoxLowercase.Text = "Lowercase (a-z)"; $checkBoxLowercase.Checked = $true
$groupBoxLetters.Controls.Add($checkBoxLowercase)
$checkBoxUppercase = New-Object System.Windows.Forms.CheckBox
$checkBoxUppercase.Location = New-Object System.Drawing.Point(210, 25); $checkBoxUppercase.AutoSize = $true
$checkBoxUppercase.Text = "Uppercase (A-Z)"; $checkBoxUppercase.Checked = $true
$groupBoxLetters.Controls.Add($checkBoxUppercase)

# Number/Symbol Counts (inside Password GroupBox)
$labelNumNumbers = New-Object System.Windows.Forms.Label
$labelNumNumbers.Location = New-Object System.Drawing.Point(10, 125)
$labelNumNumbers.Size = New-Object System.Drawing.Size(120, 20)
$labelNumNumbers.Text = "Number of Digits:"
$groupBoxPasswordOpts.Controls.Add($labelNumNumbers)

$numericUpDownNumNumbers = New-Object System.Windows.Forms.NumericUpDown
$numericUpDownNumNumbers.Location = New-Object System.Drawing.Point(140, 123)
$numericUpDownNumNumbers.Size = New-Object System.Drawing.Size(60, 25)
$numericUpDownNumNumbers.Minimum = 0
$numericUpDownNumNumbers.Maximum = 20 # Arbitrary max, adjust as needed
$numericUpDownNumNumbers.Value = 1 # Default count
$numericUpDownNumNumbers.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$groupBoxPasswordOpts.Controls.Add($numericUpDownNumNumbers)

$labelNumSpecial = New-Object System.Windows.Forms.Label
$labelNumSpecial.Location = New-Object System.Drawing.Point(220, 125) # Position next to numbers
$labelNumSpecial.Size = New-Object System.Drawing.Size(130, 20)
$labelNumSpecial.Text = "Number of Symbols:"
$groupBoxPasswordOpts.Controls.Add($labelNumSpecial)

$numericUpDownNumSpecial = New-Object System.Windows.Forms.NumericUpDown
$numericUpDownNumSpecial.Location = New-Object System.Drawing.Point(355, 123) # Position next to label
$numericUpDownNumSpecial.Size = New-Object System.Drawing.Size(60, 25)
$numericUpDownNumSpecial.Minimum = 0
$numericUpDownNumSpecial.Maximum = 20 # Arbitrary max
$numericUpDownNumSpecial.Value = 1 # Default count
$numericUpDownNumSpecial.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$groupBoxPasswordOpts.Controls.Add($numericUpDownNumSpecial)


# Other Password Options (inside Password GroupBox)
$checkBoxExcludeAmbiguous = New-Object System.Windows.Forms.CheckBox
$checkBoxExcludeAmbiguous.Location = New-Object System.Drawing.Point(15, 160) # Position below counts
$checkBoxExcludeAmbiguous.AutoSize = $true
$checkBoxExcludeAmbiguous.Text = "Exclude Ambiguous Characters (l, 1, I, O, 0)"
$checkBoxExcludeAmbiguous.Checked = $false
$groupBoxPasswordOpts.Controls.Add($checkBoxExcludeAmbiguous)

$labelCustomSpecial = New-Object System.Windows.Forms.Label
$labelCustomSpecial.Location = New-Object System.Drawing.Point(10, 195) # Position below checkbox
$labelCustomSpecial.Size = New-Object System.Drawing.Size(420, 20)
$labelCustomSpecial.Text = "Custom Special Characters (overrides default if not empty):"
$groupBoxPasswordOpts.Controls.Add($labelCustomSpecial)

$textBoxCustomSpecial = New-Object System.Windows.Forms.TextBox
$textBoxCustomSpecial.Location = New-Object System.Drawing.Point(15, 215) # Position below label
$textBoxCustomSpecial.Size = New-Object System.Drawing.Size(410, 25)
$textBoxCustomSpecial.Font = New-Object System.Drawing.Font("Consolas", 9)
$textBoxCustomSpecial.Text = ""
$groupBoxPasswordOpts.Controls.Add($textBoxCustomSpecial)


# --- Passphrase Options GroupBox ---
$groupBoxPassphraseOpts = New-Object System.Windows.Forms.GroupBox
$groupBoxPassphraseOpts.Location = New-Object System.Drawing.Point(10, 45) # Same location, will toggle visibility
$groupBoxPassphraseOpts.Size = New-Object System.Drawing.Size(440, 220) # Adjusted size
$groupBoxPassphraseOpts.Text = "Passphrase Options"
$groupBoxPassphraseOpts.Visible = $false # Initially hidden
$form.Controls.Add($groupBoxPassphraseOpts)

# Number of Words
$labelPpNumWords = New-Object System.Windows.Forms.Label
$labelPpNumWords.Location = New-Object System.Drawing.Point(10, 25)
$labelPpNumWords.Size = New-Object System.Drawing.Size(120, 20)
$labelPpNumWords.Text = "Number of Words:"
$groupBoxPassphraseOpts.Controls.Add($labelPpNumWords)

$numericUpDownPpNumWords = New-Object System.Windows.Forms.NumericUpDown
$numericUpDownPpNumWords.Location = New-Object System.Drawing.Point(140, 23)
$numericUpDownPpNumWords.Size = New-Object System.Drawing.Size(60, 25)
$numericUpDownPpNumWords.Minimum = 3
$numericUpDownPpNumWords.Maximum = 10
$numericUpDownPpNumWords.Value = 4 # Default words
$numericUpDownPpNumWords.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$groupBoxPassphraseOpts.Controls.Add($numericUpDownPpNumWords)

# Separator
$labelPpSeparator = New-Object System.Windows.Forms.Label
$labelPpSeparator.Location = New-Object System.Drawing.Point(220, 25)
$labelPpSeparator.Size = New-Object System.Drawing.Size(70, 20)
$labelPpSeparator.Text = "Separator:"
$groupBoxPassphraseOpts.Controls.Add($labelPpSeparator)

$textBoxPpSeparator = New-Object System.Windows.Forms.TextBox
$textBoxPpSeparator.Location = New-Object System.Drawing.Point(295, 23)
$textBoxPpSeparator.Size = New-Object System.Drawing.Size(40, 25)
$textBoxPpSeparator.Text = "-" # Default separator
$textBoxPpSeparator.MaxLength = 3
$groupBoxPassphraseOpts.Controls.Add($textBoxPpSeparator)

# Capitalize Option
$checkBoxPpCapitalize = New-Object System.Windows.Forms.CheckBox
$checkBoxPpCapitalize.Location = New-Object System.Drawing.Point(15, 60)
$checkBoxPpCapitalize.AutoSize = $true
$checkBoxPpCapitalize.Text = "Capitalise Each Word"
$checkBoxPpCapitalize.Checked = $true
$groupBoxPassphraseOpts.Controls.Add($checkBoxPpCapitalize)

# Insert Counts (Changed from Append)
$labelPpNumNumbers = New-Object System.Windows.Forms.Label
$labelPpNumNumbers.Location = New-Object System.Drawing.Point(10, 95)
$labelPpNumNumbers.Size = New-Object System.Drawing.Size(160, 20)
$labelPpNumNumbers.Text = "Insert # of Digits (0-9):" # Changed text
$groupBoxPassphraseOpts.Controls.Add($labelPpNumNumbers)

$numericUpDownPpNumNumbers = New-Object System.Windows.Forms.NumericUpDown
$numericUpDownPpNumNumbers.Location = New-Object System.Drawing.Point(175, 93)
$numericUpDownPpNumNumbers.Size = New-Object System.Drawing.Size(60, 25)
$numericUpDownPpNumNumbers.Minimum = 0
$numericUpDownPpNumNumbers.Maximum = 5 # Max insert count
$numericUpDownPpNumNumbers.Value = 0 # Default insert count
$numericUpDownPpNumNumbers.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$groupBoxPassphraseOpts.Controls.Add($numericUpDownPpNumNumbers)

$labelPpNumSymbols = New-Object System.Windows.Forms.Label
$labelPpNumSymbols.Location = New-Object System.Drawing.Point(10, 125) # Below numbers
$labelPpNumSymbols.Size = New-Object System.Drawing.Size(160, 20)
$labelPpNumSymbols.Text = "Insert # of Symbols:" # Changed text
$groupBoxPassphraseOpts.Controls.Add($labelPpNumSymbols)

$numericUpDownPpNumSymbols = New-Object System.Windows.Forms.NumericUpDown
$numericUpDownPpNumSymbols.Location = New-Object System.Drawing.Point(175, 123) # Below numbers input
$numericUpDownPpNumSymbols.Size = New-Object System.Drawing.Size(60, 25)
$numericUpDownPpNumSymbols.Minimum = 0
$numericUpDownPpNumSymbols.Maximum = 5 # Max insert count
$numericUpDownPpNumSymbols.Value = 0 # Default insert count
$numericUpDownPpNumSymbols.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$groupBoxPassphraseOpts.Controls.Add($numericUpDownPpNumSymbols)

# Word List Status Label (inside Passphrase GroupBox)
$labelWordListStatus = New-Object System.Windows.Forms.Label
$labelWordListStatus.Location = New-Object System.Drawing.Point(10, 185) # Adjusted position
$labelWordListStatus.Size = New-Object System.Drawing.Size(420, 20)
$labelWordListStatus.Text = "Word list status: Not loaded"
$labelWordListStatus.ForeColor = [System.Drawing.Color]::Gray
$groupBoxPassphraseOpts.Controls.Add($labelWordListStatus)


# --- Common Controls (Below Options) ---

# Generate Button
$buttonGenerate = New-Object System.Windows.Forms.Button
$buttonGenerate.Location = New-Object System.Drawing.Point(20, 325) # Adjusted Y position
$buttonGenerate.Size = New-Object System.Drawing.Size(420, 35) # Wider/Taller
$buttonGenerate.Text = "Generate" # Changed text slightly
$form.Controls.Add($buttonGenerate)

# Result Display Textbox
$textBoxResult = New-Object System.Windows.Forms.TextBox
$textBoxResult.Location = New-Object System.Drawing.Point(20, 375) # Adjusted Y position
$textBoxResult.Size = New-Object System.Drawing.Size(335, 25) # Adjusted size
$textBoxResult.ReadOnly = $true
$textBoxResult.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxResult)

# Copy Button
$buttonCopy = New-Object System.Windows.Forms.Button
$buttonCopy.Location = New-Object System.Drawing.Point(365, 373) # Adjusted Y/X position
$buttonCopy.Size = New-Object System.Drawing.Size(75, 27)
$buttonCopy.Text = "Copy"
$buttonCopy.Enabled = $false # Disabled initially
$form.Controls.Add($buttonCopy)

# --- Timer for Checking Download Job ---
$timerDownloadCheck = New-Object System.Windows.Forms.Timer
$timerDownloadCheck.Interval = 500 # Check every 500ms

$timerDownloadCheck_Tick = {
    # Check if the job exists and is running
    if ($Script:DownloadJob -ne $null -and $Script:DownloadJob.State -ne 'Running') {
        # Job finished (Completed or Failed)
        $timerDownloadCheck.Stop() # Stop the timer

        $jobResult = $Script:DownloadJob | Receive-Job -Keep

        # Ensure the form hasn't been closed
        if ($form -and -not $form.IsDisposed) {
            if ($jobResult) {
                if ($jobResult.Success) {
                    $Script:EffWordList = $jobResult.Words
                    # Update UI (runs on UI thread from Timer Tick)
                    $labelWordListStatus.Text = "Word list downloaded successfully ($($jobResult.Count) words)."
                    $labelWordListStatus.ForeColor = [System.Drawing.Color]::Green
                    $buttonGenerate.Enabled = $true
                } else {
                    # Update UI on failure
                    $labelWordListStatus.Text = "Word list download failed."
                    $labelWordListStatus.ForeColor = [System.Drawing.Color]::Red
                    $buttonGenerate.Enabled = $false
                    [System.Windows.Forms.MessageBox]::Show($form, "Failed to download word list.`nError: $($jobResult.ErrorMessage)", "Word List Error", 'OK', 'Error')
                }
            } else {
                 $labelWordListStatus.Text = "Job completed but no result received."
                 $labelWordListStatus.ForeColor = [System.Drawing.Color]::Red
            }
        }
        # Clean up the job
        $Script:DownloadJob | Remove-Job -Force
        $Script:DownloadJob = $null
    }
    # If job is still running, the timer will tick again
}
$timerDownloadCheck.Add_Tick($timerDownloadCheck_Tick)


# --- Event Handlers ---

# Mode Radio Button Change Event
$ModeChangedHandler = {
    param($sender, $e)
    # Stop timer if it's running
    if ($timerDownloadCheck.Enabled) {
        $timerDownloadCheck.Stop()
    }
    # Clean up any existing job if switching modes
    if ($Script:DownloadJob -ne $null) {
        Write-Warning "Download job cancelled due to mode change."
        $Script:DownloadJob | Remove-Job -Force
        $Script:DownloadJob = $null
    }

    if ($radioPasswordMode.Checked) {
        $groupBoxPasswordOpts.Visible = $true
        $groupBoxPassphraseOpts.Visible = $false
        $buttonGenerate.Text = "Generate Password"
        $buttonGenerate.Enabled = $true
    } else { # Passphrase mode
        $groupBoxPasswordOpts.Visible = $false
        $groupBoxPassphraseOpts.Visible = $true
        $buttonGenerate.Text = "Generate Passphrase"
        # Attempt to fetch word list if not loaded
        if ($Script:EffWordList -eq $null) {
            $buttonGenerate.Enabled = $false
            $labelWordListStatus.Text = "Downloading word list (background)..."
            $labelWordListStatus.ForeColor = [System.Drawing.Color]::Orange
            $labelWordListStatus.Refresh()

            # Start the background job
            $Script:DownloadJob = Start-Job -ScriptBlock {
                param($Url)
                $WordList = $null
                try {
                    Write-Host "Job: Downloading word list from $Url"
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls
                    $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 30
                    if ($response.StatusCode -eq 200) {
                        $lines = $response.Content -split '[\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d{5}\s+\w+' }
                        $WordList = $lines | ForEach-Object { ($_ -split '\s+', 2)[1] }
                        Write-Host "Job: Downloaded $($WordList.Count) words."
                        return @{ Success = $true; Words = $WordList; Count = $WordList.Count }
                    } else {
                        throw "Job: Failed download. Status: $($response.StatusCode)"
                    }
                } catch {
                    Write-Error "Job Error: $($_.Exception.Message)"
                    return @{ Success = $false; ErrorMessage = $_.Exception.Message }
                }
            } -ArgumentList $Script:WordListUrl

            # Start the timer to check the job status
            $timerDownloadCheck.Start()

        } else {
             # Word list already loaded
             $labelWordListStatus.Text = "Word list ready ($($Script:EffWordList.Count) words)."
             $labelWordListStatus.ForeColor = [System.Drawing.Color]::Green
             $buttonGenerate.Enabled = $true
        }
    }
    # Clear previous result on mode change
    $textBoxResult.Text = ""
    $buttonCopy.Enabled = $false
}
# Attach the handler to both radio buttons
$radioPasswordMode.add_CheckedChanged($ModeChangedHandler)
$radioPassphraseMode.add_CheckedChanged($ModeChangedHandler)

# Generate Button Click Event
$buttonGenerate.Add_Click({
    $generatedResult = $null

    if ($radioPasswordMode.Checked) {
        # --- Password Mode ---
        $length = $numericUpDownLength.Value
        $includeLower = $checkBoxLowercase.Checked
        $includeUpper = $checkBoxUppercase.Checked
        $numNumbers = $numericUpDownNumNumbers.Value # Get count from NumericUpDown
        $numSpecial = $numericUpDownNumSpecial.Value # Get count from NumericUpDown
        $excludeAmbiguous = $checkBoxExcludeAmbiguous.Checked
        $customSpecial = $textBoxCustomSpecial.Text
        # Call password generation function with counts
        $generatedResult = Generate-SecurePassword -Length $length `
            -IncludeLowercase $includeLower -IncludeUppercase $includeUpper `
            -NumberOfNumbers $numNumbers -NumberOfSpecialChars $numSpecial `
            -ExcludeAmbiguous $excludeAmbiguous -CustomSpecialChars $customSpecial

    } else {
        # --- Passphrase Mode ---
        if ($Script:EffWordList -eq $null -or $Script:EffWordList.Count -eq 0) {
             [System.Windows.Forms.MessageBox]::Show("Word list is not available. Cannot generate passphrase.", "Error", 'OK', 'Error')
             return
        }
        # Get passphrase options including counts
        $numWords = $numericUpDownPpNumWords.Value
        $separator = $textBoxPpSeparator.Text
        $capitalize = $checkBoxPpCapitalize.Checked
        $numPpNumbers = $numericUpDownPpNumNumbers.Value # Get count from NumericUpDown
        $numPpSymbols = $numericUpDownPpNumSymbols.Value # Get count from NumericUpDown
        # Call passphrase generation function with counts
        $generatedResult = Generate-Passphrase -NumberOfWords $numWords `
            -Separator $separator -Capitalize $capitalize `
            -NumberOfNumbers $numPpNumbers -NumberOfSymbols $numPpSymbols
    }

    # Display the result if generated successfully
    if ($generatedResult -ne $null) {
        $textBoxResult.Text = $generatedResult
        $buttonCopy.Enabled = $true
    } else {
        # Clear result if generation failed
        $textBoxResult.Text = ""
        $buttonCopy.Enabled = $false
    }
})

# Copy Button Click Event
$buttonCopy.Add_Click({
    if (-not [string]::IsNullOrEmpty($textBoxResult.Text)) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($textBoxResult.Text)
            # Provide visual feedback ("Copied!")
            $originalText = $buttonCopy.Text; $buttonCopy.Text = "Copied!"; $buttonCopy.Enabled = $false; $form.Refresh()
            Start-Sleep -Milliseconds 750 # Pause briefly
            # Restore button text if form still exists
            if ($form -and -not $form.IsDisposed) { $buttonCopy.Text = $originalText; $buttonCopy.Enabled = $true; $form.Refresh() }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not copy to clipboard. Error: $($_.Exception.Message)", "Clipboard Error", 'OK', 'Warning')
        }
    }
})

# --- Show the Form ---
$form.Add_Shown({
    $form.Activate()
    # Trigger mode check on load to set initial state correctly
    $ModeChangedHandler.Invoke($null, $null)
})
# Add Form Closing event to clean up timer and job
$form.Add_FormClosing({
    # Stop the timer if it's running
    if ($timerDownloadCheck.Enabled) {
        $timerDownloadCheck.Stop()
    }
    # Remove any lingering job
    if ($Script:DownloadJob -ne $null) {
        Write-Host "Cleaning up download job on form close."
        $Script:DownloadJob | Remove-Job -Force
        $Script:DownloadJob = $null
    }
    # Dispose the timer
    $timerDownloadCheck.Dispose()
})

# Display the form as a modal dialog
[void]$form.ShowDialog()

# --- Clean up Form Resources ---
# Form object is disposed automatically after ShowDialog() returns

# Clean up any jobs that might somehow persist
Get-Job -State Completed -EA SilentlyContinue | Remove-Job -Force
Get-Job -State Failed -EA SilentlyContinue | Remove-Job -Force
