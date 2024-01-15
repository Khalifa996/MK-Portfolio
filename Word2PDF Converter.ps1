# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Initialize main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Word2PDF Converter"
$mainForm.Size = New-Object System.Drawing.Size(400, 200)
$mainForm.StartPosition = 'CenterScreen'

# Attach FormClosed event
$mainForm.Add_FormClosed({ [Environment]::Exit(0) })

# Initialize Open File Dialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Word Documents|*.doc;*.docx"

# Initialize button to open file dialog
$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = "Select File to be converted"
$openButton.Size = New-Object System.Drawing.Size(200, 40)
$openButton.Location = New-Object System.Drawing.Point(($mainForm.Size.Width / 2 - $openButton.Size.Width / 2), ($mainForm.Size.Height / 2 - $openButton.Size.Height / 2))
$openButton.Add_Click({
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $filePath = $openFileDialog.FileName
        Convert-WordToPdf -InputPath $filePath
    }
})

# Add button to form
$mainForm.Controls.Add($openButton)

# Function to convert Word to PDF
function Convert-WordToPdf {
    param (
        [string]$InputPath
    )

    $word = $null
    $doc = $null

    try {
        # Initialize Word COM Object in headless mode
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false  # Make Word invisible

        # Open Word Document
        $doc = $word.Documents.Open($InputPath)

        # Define PDF save format (17)
        $pdfFormat = 17

        # Generate PDF path
        $pdfPath = [System.IO.Path]::ChangeExtension($InputPath, 'pdf')

        # Save as PDF
        $doc.SaveAs([ref]$pdfPath, [ref]$pdfFormat)

        # Prevent changes to Normal.dotm
        $word.NormalTemplate.Saved = $true

        [System.Windows.Forms.MessageBox]::Show("Successfully converted to PDF!")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_")
    } finally {
        Start-Sleep -Seconds 1
        if ($doc) {
            $doc.Saved = $true  # Explicitly mark the document as saved
            $doc.Close([ref]$false)  # Close without saving changes
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
    
        Start-Sleep -Seconds 1
        if ($word) {
            $word.NormalTemplate.Saved = $true  # Explicitly mark the template as saved
            $word.Quit([ref]$false)  # Quit without saving changes
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }    
}

# Show form
$mainForm.ShowDialog()

