#**********************************************************
#*Carter C. 1/9/17
#*
#*Connects to IP Address Table
#*Allows user to edit data in table and save those changes
#**********************************************************


Param
(
    $Instance = 'XXXX.XXXX.XXXX.com', #***CHANGE TO CORRECT SERVER***
    $Database = 'ipaddress',
    $Query = 'select * from PrinterAssets order by CAST(PARSENAME([ReservationIP], 4) AS INT),
CAST(PARSENAME([ReservationIP], 3) AS INT),
CAST(PARSENAME([ReservationIP], 2) AS INT),
CAST(PARSENAME([ReservationIP], 1) AS INT)' 
    #$Query2 = "select * from IPAddresses where ipaddress= 'XXX.XXX.XXX.XXX'"
    #$Query2 = $userText
)

#********************************************************************************************
#********************CLICKITY CLACK ACTION PACKS*********************************************
#********************************************************************************************
$SelectIP_Click2 =          #calls function that searches by IP address entered in text box
{
        $userText = $ip2TxtSQLQuery.Text
        #$form1.Close()
        (Show-SelectIP)
} #End SelectIP_Click2

$SelectIP_Click23 =         #calls function that searches by device entered in text box
{
        $TEST = $ip2TxtSQLQuery353.text #removed .txt
        #$form1.Close()
        #Write-host $TEST
        (Show-SelectIP23)
} #End SelectIP_Click2

$Quit2=                     #closes window
{
    Write-Debug "closing the form"
    #$form1.Close()
} #End Quit2 

function Refresh            #refreshes window to reflect most recent data
{ 
    $form1.Close()
    $form1.Dispose()
    GenerateForm
}

$Refi_Click =               #edit click -- opens edittable datagridview
{
    GenFormRefi
    #$form1.Close()
} #End Refi_Click

$Scanners_Click = 
{
    GenFormScanners
}

$Printers_Click = 
{
    GenFormPrinters
}

$GenForm_Click =            #refresh button on edittable datagrid - opens another parent window
{
    $form3.Close()
    GenerateForm
} #End GenForm_Click

$datagridview1_Click =      #double click on ID to call function that pulls data for that specific ID
{
    #[void][System.Windows.Forms.MessageBox]::Show($datagridview1.SelectedCells[0].FormattedValue, 'You Chose')
    $selectedCell = $datagridview1.SelectedCells[0].FormattedValue
    Show-SelectIP35
} #End datagridview1_Click

#*********************************NO MAS CLICKITY CLACK ACTION PACKS :,(**********************

#*********************************************************************************************
#**************GENFORMREFI - EDITABLE SHEET THAT WILL NO CLOSE PARENT SHEET*******************
#**************CLICKING EDIT ON PARENT SHEET TRIGGERS THIS FORM*******************************
#*********************************************************************************************

function GenFormRefi{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::EnableVisualStyles()             # Whole bunch of stuff I don't know how to explain
    $form3 = New-Object 'System.Windows.Forms.Form'                      # but it works
    $datagridview1 = New-Object 'System.Windows.Forms.DataGridView'      #  
    $buttonOK = New-Object 'System.Windows.Forms.Button'                 #New-Object creation section for this form
    $ip2TxtSQLQuery = New-Object 'System.Windows.Forms.TextBox'          # 
    $ip2TxtSQLQuery353 = New-Object 'System.Windows.Forms.TextBox'
    $btn353 = New-Object 'System.Windows.Forms.Button'      
    $btnQuit2 = New-Object 'System.Windows.Forms.Button'                 #   
    $btn35 = New-Object 'System.Windows.Forms.Button'                    #
    $btnGenForm = New-Object 'System.Windows.Forms.Button'               #

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState' #does something

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=True" #SSPI

    $form3_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)
    }

    #****************THIS IS THE buttonOK CLICK ACTION***SAVES AND CLOSES USER EDITS TO DATA***********************
    #************THE PopUP ARE YOU SURE BUTTON IS A LIAR, IT WILL SAVE WHETHER YOU LIKE IT OR NOT******************
    $buttonOK_Click = {
        if ($script:DVGhasChanged -and [System.Windows.Forms.MessageBox]::Show("*Click* Yes.. I Dare you..", 'Data Changed', 'YesNo'))
        {            
            $script:adapter.Update($datagridview1.DataSource)
        }
    }

    $datagridview1_CurrentCellDirtyStateChanged = {
        $script:DVGhasChanged = $true
    }
    $Form_StateCorrection_Load = {
        $form3.WindowState = $InitialFormWindowState
    }
    
    $form3.SuspendLayout()

    #*******************FORM3**WILL RENAME LATER*********************************************
    $form3.Controls.Add($datagridview1)
    $form3.Controls.Add($buttonOK)
    #$form3.AcceptButton = $buttonOK
    $form3.Controls.Add($btnQuit2)
    $form3.Controls.Add($btnGenForm)
    $form3.Controls.Add($ip2TxtSQLQuery)
    $form3.Controls.Add($btn35)
    $form3.ClientSize = '900, 600'
    $form3.FormBorderStyle = 'Sizable'
    $form3.BackColor = "Pink" #[System.Drawing.Color]::FromArgb(255,185,209,234)
    $form3.MaximizeBox = $True
    $form3.MinimizeBox = $True
    $form3.Name = 'form3'
    $form3.StartPosition = 'CenterScreen'
    $form3.Text = '***EDIT MODE***'
    $form3.KeyPreview = $True
    $form3.add_Load($form3_Load)

    #***********TEXT BOX THAT WILL SEARCH QUERY FOR User Enter IPADDRESS - 10.10.60.10 Is AN EXAMPLE***********************
    $ip2TxtSQLQuery.Text = "10.10.60.10"
    $ip2TxtSQLQuery.Name = 'txtSQLQuery'
    $ip2TxtSQLQuery.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 20
    $ip2TxtSQLQuery.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 46
    $ip2TxtSQLQuery.Location = $System_Drawing_Point
    $ip2TxtSQLQuery.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn35.PerformClick()
        }
    })

    #****BUTTON THAT GOES WITH THE ABOVE TEXT BOX - INITIATES QUERY FOR THAT IP AND OPENS NEW WINDOW - SEE CLICK ACTION*********
    $btn35.Anchor = 'Top, Left'
    $btn35.UseVisualStyleBackColor = $True
    $btn35.Text = 'Select IP'
    $btn35.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn35.TabIndex = 3
    $btn35.Name = 'btn35'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btn35.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 45
    $btn35.Location = $System_Drawing_Point
    $btn35.add_Click($SelectIP_Click2)
    
    #****************

    $ip2TxtSQLQuery353.Text = "Enter   Asset Number"
    $ip2TxtSQLQuery353.Name = 'txtSQLQuery353'
    $ip2TxtSQLQuery353.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 20
    $ip2TxtSQLQuery353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 16
    $ip2TxtSQLQuery353.Location = $System_Drawing_Point
    $ip2TxtSQLQuery353.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery353.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery353.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn353.PerformClick()
        }
    })
    
    $form3.Controls.Add($ip2TxtSQLQuery353)

    $btn353.Anchor = 'Top, Left'
    $btn353.UseVisualStyleBackColor = $True
    $btn353.Text = 'Select Device'
    $btn353.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn353.TabIndex = 4
    $btn353.Name = 'btn353'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 90
    $System_Drawing_Size.Height = 23
    $btn353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 15
    $btn353.Location = $System_Drawing_Point
    $btn353.add_Click($SelectIP_Click23)

    $form3.Controls.Add($btn353)

    
    #***********************DATA GRID VIEW UNO*******************************************************************

    $dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridView1.DefaultCellStyle.BackColor = "White"
    $dataGridView1.BackgroundColor = "White"
    $dataGridView1.Name = 'dataGridView'
    $dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridView1.ReadOnly = $False
    $dataGridView1.AllowUserToDeleteRows = $True
    $dataGridView1.RowHeadersVisible = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 870
    $System_Drawing_Size.Height = 480
    $dataGridView1.Size = $System_Drawing_Size
    $dataGridView1.TabIndex = 8
    $dataGridView1.Anchor = 15
    $dataGridView1.AutoSizeColumnsMode = 'AllCells'
    $dataGridView1.AllowUserToAddRows = $True
    $dataGridView1.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 70
    $dataGridView1.Location = $System_Drawing_Point
    $dataGridView1.AllowUserToOrderColumns = $True
    $dataGridView1.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridView1AutoSizeColumnsMode.AllCells
    #$datagridview1.add_DoubleClick($datagridview1_Click)  #Disabled because I don't want it there, but might be handy one day
    $datagridview1.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    $datagridview1.add_CurrentCellDirtyStateChanged($datagridview1_CurrentCellDirtyStateChanged)

    #*****************buttonOK Creation*****************
    $buttonOK.Anchor = 'Bottom, Left'
    $buttonOK.DialogResult = 'Ok'
    $buttonOK.Location = '13,565'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '105,23' #75
    $buttonOK.Text = '&Save and Close'
    $buttonOK.UseVisualStyleBackColor = $True
    $buttonOK.add_Click($buttonOK_Click)

    #*********btnQuit2 Creation -- see click action for more**********
    $btnQuit2.Anchor = 'Bottom, Left'
    $btnQuit2.DialogResult = 'Ok'
    $btnQuit2.Location = '125,565'
    $btnQuit2.Name = 'buttonOK'
    $btnQuit2.Size = '75,23'
    $btnQuit2.Text = '&Close'
    $btnQuit2.UseVisualStyleBackColor = $True
    $btnQuit2.add_Click($Quit2)

    #*********btnGenForm Creation -- see click action for more*********
    $btnGenForm.UseVisualStyleBackColor = $True
    $btnGenForm.Text = 'Refresh'
    $btnGenForm.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnGenForm.TabIndex = 1
    $btnGenForm.Name = 'btnGenForm'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btnGenForm.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 15
    $btnGenForm.Location = $System_Drawing_Point
    $btnGenForm.add_Click($GenForm_Click)

    $form3.ResumeLayout()   #** no idea what this does, but it does important stuff

    $InitialFormWindowState = $form3.WindowState         #** no idea what this does, but it does important stuff
    $form3.add_Load($Form_StateCorrection_Load)          #** no idea what this does, but it does important stuff
    $form3.ShowDialog()                        

}

function GenFormScanners{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::EnableVisualStyles()             # Whole bunch of stuff I don't know how to explain
    $formSCN = New-Object 'System.Windows.Forms.Form'                      # but it works
    $datagridview1 = New-Object 'System.Windows.Forms.DataGridView'      #  
    $buttonOK = New-Object 'System.Windows.Forms.Button'                 #New-Object creation section for this form
    $ip2TxtSQLQuery = New-Object 'System.Windows.Forms.TextBox'          # 
    $ip2TxtSQLQuery353 = New-Object 'System.Windows.Forms.TextBox'
    $btn353 = New-Object 'System.Windows.Forms.Button'      
    $btnQuit2 = New-Object 'System.Windows.Forms.Button'                 #   
    $btn35 = New-Object 'System.Windows.Forms.Button'                    #
    $btnGenForm = New-Object 'System.Windows.Forms.Button'               #

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState' #does something

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=True" #SSPI

    $formSCN_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select * FROM PrinterAssets where AssetType= 'Scanner' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)
    }

    #****************THIS IS THE buttonOK CLICK ACTION***SAVES AND CLOSES USER EDITS TO DATA***********************
    #************THE PopUP ARE YOU SURE BUTTON IS A LIAR, IT WILL SAVE WHETHER YOU LIKE IT OR NOT******************
    $buttonOK_Click = {
        if ($script:DVGhasChanged -and [System.Windows.Forms.MessageBox]::Show("*Click* Yes.. I Dare you..", 'Data Changed', 'YesNo'))
        {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing
            $FormT = New-Object system.Windows.Forms.Form
            $FormT.Width = '500' #400
            $FormT.Height = '300' #200
            $file2 = (Get-item "K:\powershell\WARNING.jpg")
            $Image2 = [System.Drawing.Image]::FromFile($file2)
            $FormT.BackgroundImage  = $Image2
            $FormT.BackgroundImageLayout = 'Stretch'
            $FormT.BackColor = "White"
            $file = (Get-Item "K:\powershell\boom-boomm.jpg")
            $Image = [system.drawing.image]::FromFile($file)
            $script:Label = New-Object System.Windows.Forms.Label
            $script:Label.ForeColor = "Black"
            $script:Label.BackColor = "Red"
            $script:Label.Font = New-Object System.Drawing.Font("Arial",20,[System.Drawing.FontStyle]::Bold)
            $FormT.Controls.Add($boom)
            $script:Label.AutoSize = $true
            $script:Label.Location = '25,120'
            $FormT.Controls.Add($Label)
            $Timer = New-Object System.Windows.Forms.Timer
            $Timer.Interval = 1000
            $script:CountDown = 2
            $Timer.add_Tick(
            {
                $script:Label.Text = "          Self destruct initiated.
                 
Pete's System will be wiped in :
                  $CountDown seconds"
                $script:CountDown--
                if($script:CountDown -eq -1)
                {
                    $FormT.Controls.Remove($Label)
                    $FormT.BackgroundImage = $Image
                    $FormT.BackgroundImageLayout = "Stretch"
                }
                if($script:CountDown -eq -2)
                {
                    $Timer.Stop()
                    $FormT.close()
                }
            }
            )
            $Timer.Start()
            $FormT.ShowDialog()
            
            $script:adapter.Update($datagridview1.DataSource)
        }
    }

    $datagridview1_CurrentCellDirtyStateChanged = {
        $script:DVGhasChanged = $true
    }
    $Form_StateCorrection_Load = {
        $formSCN.WindowState = $InitialFormWindowState
    }
    
    $formSCN.SuspendLayout()

    #*******************FORM3**WILL RENAME LATER*********************************************
    $formSCN.Controls.Add($datagridview1)
    $formSCN.Controls.Add($buttonOK)
    #$form3.AcceptButton = $buttonOK
    $formSCN.Controls.Add($btnQuit2)
    $formSCN.Controls.Add($btnGenForm)
    $formSCN.Controls.Add($ip2TxtSQLQuery)
    $formSCN.Controls.Add($btn35)
    $formSCN.ClientSize = '900, 600'
    $formSCN.FormBorderStyle = 'Sizable'
    $formSCN.BackColor = "Pink" #[System.Drawing.Color]::FromArgb(255,185,209,234)
    $formSCN.MaximizeBox = $True
    $formSCN.MinimizeBox = $True
    $formSCN.Name = 'form3'
    $formSCN.StartPosition = 'CenterScreen'
    $formSCN.Text = '***EDIT SCANNERS MODE***'
    $formSCN.KeyPreview = $True
    $formSCN.add_Load($formSCN_Load)

    #***********TEXT BOX THAT WILL SEARCH QUERY FOR User Enter IPADDRESS - 10.10.60.10 Is AN EXAMPLE***********************
    $ip2TxtSQLQuery.Text = "10.10.60.10"
    $ip2TxtSQLQuery.Name = 'txtSQLQuery'
    $ip2TxtSQLQuery.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 20
    $ip2TxtSQLQuery.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 46
    $ip2TxtSQLQuery.Location = $System_Drawing_Point
    $ip2TxtSQLQuery.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn35.PerformClick()
        }
    })

    #****BUTTON THAT GOES WITH THE ABOVE TEXT BOX - INITIATES QUERY FOR THAT IP AND OPENS NEW WINDOW - SEE CLICK ACTION*********
    $btn35.Anchor = 'Top, Left'
    $btn35.UseVisualStyleBackColor = $True
    $btn35.Text = 'Select IP'
    $btn35.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn35.TabIndex = 3
    $btn35.Name = 'btn35'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btn35.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 45
    $btn35.Location = $System_Drawing_Point
    $btn35.add_Click($SelectIP_Click2)
    
    #****************

    $ip2TxtSQLQuery353.Text = "Enter   Asset Number"
    $ip2TxtSQLQuery353.Name = 'txtSQLQuery353'
    $ip2TxtSQLQuery353.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 20
    $ip2TxtSQLQuery353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 16
    $ip2TxtSQLQuery353.Location = $System_Drawing_Point
    $ip2TxtSQLQuery353.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery353.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery353.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn353.PerformClick()
        }
    })
    
    $formSCN.Controls.Add($ip2TxtSQLQuery353)

    $btn353.Anchor = 'Top, Left'
    $btn353.UseVisualStyleBackColor = $True
    $btn353.Text = 'Select Device'
    $btn353.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn353.TabIndex = 4
    $btn353.Name = 'btn353'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 90
    $System_Drawing_Size.Height = 23
    $btn353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 15
    $btn353.Location = $System_Drawing_Point
    $btn353.add_Click($SelectIP_Click23)

    $formSCN.Controls.Add($btn353)

    
    #***********************DATA GRID VIEW UNO*******************************************************************

    $dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridView1.DefaultCellStyle.BackColor = "White"
    $dataGridView1.BackgroundColor = "White"
    $dataGridView1.Name = 'dataGridView'
    $dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridView1.ReadOnly = $False
    $dataGridView1.AllowUserToDeleteRows = $True
    $dataGridView1.RowHeadersVisible = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 870
    $System_Drawing_Size.Height = 480
    $dataGridView1.Size = $System_Drawing_Size
    $dataGridView1.TabIndex = 8
    $dataGridView1.Anchor = 15
    $dataGridView1.AutoSizeColumnsMode = 'AllCells'
    $dataGridView1.AllowUserToAddRows = $True
    $dataGridView1.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 70
    $dataGridView1.Location = $System_Drawing_Point
    $dataGridView1.AllowUserToOrderColumns = $True
    $dataGridView1.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridView1AutoSizeColumnsMode.AllCells
    #$datagridview1.add_DoubleClick($datagridview1_Click)  #Disabled because I don't want it there, but might be handy one day
    $datagridview1.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    $datagridview1.add_CurrentCellDirtyStateChanged($datagridview1_CurrentCellDirtyStateChanged)

    #*****************buttonOK Creation*****************
    $buttonOK.Anchor = 'Bottom, Left'
    $buttonOK.DialogResult = 'Ok'
    $buttonOK.Location = '13,565'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '105,23' #75
    $buttonOK.Text = '&Save and Close'
    $buttonOK.UseVisualStyleBackColor = $True
    $buttonOK.add_Click($buttonOK_Click)

    #*********btnQuit2 Creation -- see click action for more**********
    $btnQuit2.Anchor = 'Bottom, Left'
    $btnQuit2.DialogResult = 'Ok'
    $btnQuit2.Location = '125,565'
    $btnQuit2.Name = 'buttonOK'
    $btnQuit2.Size = '75,23'
    $btnQuit2.Text = '&Close'
    $btnQuit2.UseVisualStyleBackColor = $True
    $btnQuit2.add_Click($Quit2)

    #*********btnGenForm Creation -- see click action for more*********
    $btnGenForm.UseVisualStyleBackColor = $True
    $btnGenForm.Text = 'Refresh'
    $btnGenForm.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnGenForm.TabIndex = 1
    $btnGenForm.Name = 'btnGenForm'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btnGenForm.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 15
    $btnGenForm.Location = $System_Drawing_Point
    $btnGenForm.add_Click($GenForm_Click)

    $formSCN.ResumeLayout()   #** no idea what this does, but it does important stuff

    $InitialFormWindowState = $formSCN.WindowState         #** no idea what this does, but it does important stuff
    $formSCN.add_Load($Form_StateCorrection_Load)          #** no idea what this does, but it does important stuff
    $formSCN.ShowDialog()                        

}

function GenFormPrinters{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::EnableVisualStyles()             # Whole bunch of stuff I don't know how to explain
    $formPRT = New-Object 'System.Windows.Forms.Form'                      # but it works
    $datagridview1 = New-Object 'System.Windows.Forms.DataGridView'      #  
    $buttonOK = New-Object 'System.Windows.Forms.Button'                 #New-Object creation section for this form
    $ip2TxtSQLQuery = New-Object 'System.Windows.Forms.TextBox'          # 
    $ip2TxtSQLQuery353 = New-Object 'System.Windows.Forms.TextBox'
    $btn353 = New-Object 'System.Windows.Forms.Button'      
    $btnQuit2 = New-Object 'System.Windows.Forms.Button'                 #   
    $btn35 = New-Object 'System.Windows.Forms.Button'                    #
    $btnGenForm = New-Object 'System.Windows.Forms.Button'               #

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState' #does something

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=True" #SSPI

    $formPRT_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select * FROM PrinterAssets where AssetType= 'Printer' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)
    }

    #****************THIS IS THE buttonOK CLICK ACTION***SAVES AND CLOSES USER EDITS TO DATA***********************
    #************THE PopUP ARE YOU SURE BUTTON IS A LIAR, IT WILL SAVE WHETHER YOU LIKE IT OR NOT******************
    $buttonOK_Click = {
        if ($script:DVGhasChanged -and [System.Windows.Forms.MessageBox]::Show("*Click* Yes.. I Dare you..", 'Data Changed', 'YesNo'))
        {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing
            $FormT = New-Object system.Windows.Forms.Form
            $FormT.Width = '500' #400
            $FormT.Height = '300' #200
            $file2 = (Get-item "K:\powershell\WARNING.jpg")
            $Image2 = [System.Drawing.Image]::FromFile($file2)
            $FormT.BackgroundImage  = $Image2
            $FormT.BackgroundImageLayout = 'Stretch'
            $FormT.BackColor = "White"
            $file = (Get-Item "K:\powershell\boom-boomm.jpg")
            $Image = [system.drawing.image]::FromFile($file)
            $script:Label = New-Object System.Windows.Forms.Label
            $script:Label.ForeColor = "Black"
            $script:Label.BackColor = "Red"
            $script:Label.Font = New-Object System.Drawing.Font("Arial",20,[System.Drawing.FontStyle]::Bold)
            $FormT.Controls.Add($boom)
            $script:Label.AutoSize = $true
            $script:Label.Location = '25,120'
            $FormT.Controls.Add($Label)
            $Timer = New-Object System.Windows.Forms.Timer
            $Timer.Interval = 1000
            $script:CountDown = 2
            $Timer.add_Tick(
            {
                $script:Label.Text = "          Self destruct initiated.
                 
Pete's System will be wiped in :
                  $CountDown seconds"
                $script:CountDown--
                if($script:CountDown -eq -1)
                {
                    $FormT.Controls.Remove($Label)
                    $FormT.BackgroundImage = $Image
                    $FormT.BackgroundImageLayout = "Stretch"
                }
                if($script:CountDown -eq -2)
                {
                    $Timer.Stop()
                    $FormT.close()
                }
            }
            )
            $Timer.Start()
            $FormT.ShowDialog()
            
            $script:adapter.Update($datagridview1.DataSource)
        }
    }

    $datagridview1_CurrentCellDirtyStateChanged = {
        $script:DVGhasChanged = $true
    }
    $Form_StateCorrection_Load = {
        $formPRT.WindowState = $InitialFormWindowState
    }
    
    $formPRT.SuspendLayout()

    #*******************FORM3**WILL RENAME LATER*********************************************
    $formPRT.Controls.Add($datagridview1)
    $formPRT.Controls.Add($buttonOK)
    #$form3.AcceptButton = $buttonOK
    $formPRT.Controls.Add($btnQuit2)
    $formPRT.Controls.Add($btnGenForm)
    $formPRT.Controls.Add($ip2TxtSQLQuery)
    $formPRT.Controls.Add($btn35)
    $formPRT.ClientSize = '900, 600'
    $formPRT.FormBorderStyle = 'Sizable'
    $formPRT.BackColor = "Pink" #[System.Drawing.Color]::FromArgb(255,185,209,234)
    $formPRT.MaximizeBox = $True
    $formPRT.MinimizeBox = $True
    $formPRT.Name = 'form3'
    $formPRT.StartPosition = 'CenterScreen'
    $formPRT.Text = '***EDIT PRINTERS MODE***'
    $formPRT.KeyPreview = $True
    $formPRT.add_Load($formPRT_Load)

    #***********TEXT BOX THAT WILL SEARCH QUERY FOR User Enter IPADDRESS - 10.10.60.10 Is AN EXAMPLE***********************
    $ip2TxtSQLQuery.Text = "10.10.60.10"
    $ip2TxtSQLQuery.Name = 'txtSQLQuery'
    $ip2TxtSQLQuery.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 20
    $ip2TxtSQLQuery.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 46
    $ip2TxtSQLQuery.Location = $System_Drawing_Point
    $ip2TxtSQLQuery.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn35.PerformClick()
        }
    })

    #****BUTTON THAT GOES WITH THE ABOVE TEXT BOX - INITIATES QUERY FOR THAT IP AND OPENS NEW WINDOW - SEE CLICK ACTION*********
    $btn35.Anchor = 'Top, Left'
    $btn35.UseVisualStyleBackColor = $True
    $btn35.Text = 'Select IP'
    $btn35.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn35.TabIndex = 3
    $btn35.Name = 'btn35'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btn35.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 45
    $btn35.Location = $System_Drawing_Point
    $btn35.add_Click($SelectIP_Click2)
    
    #****************

    $ip2TxtSQLQuery353.Text = "Enter  Asset Number"
    $ip2TxtSQLQuery353.Name = 'txtSQLQuery353'
    $ip2TxtSQLQuery353.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 20
    $ip2TxtSQLQuery353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 16
    $ip2TxtSQLQuery353.Location = $System_Drawing_Point
    $ip2TxtSQLQuery353.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery353.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery353.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn353.PerformClick()
        }
    })
    
    $formPRT.Controls.Add($ip2TxtSQLQuery353)

    $btn353.Anchor = 'Top, Left'
    $btn353.UseVisualStyleBackColor = $True
    $btn353.Text = 'Select Device'
    $btn353.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn353.TabIndex = 4
    $btn353.Name = 'btn353'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 90
    $System_Drawing_Size.Height = 23
    $btn353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 15
    $btn353.Location = $System_Drawing_Point
    $btn353.add_Click($SelectIP_Click23)

    $formPRT.Controls.Add($btn353)

    
    #***********************DATA GRID VIEW UNO*******************************************************************

    $dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridView1.DefaultCellStyle.BackColor = "White"
    $dataGridView1.BackgroundColor = "White"
    $dataGridView1.Name = 'dataGridView'
    $dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridView1.ReadOnly = $False
    $dataGridView1.AllowUserToDeleteRows = $True
    $dataGridView1.RowHeadersVisible = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 870
    $System_Drawing_Size.Height = 480
    $dataGridView1.Size = $System_Drawing_Size
    $dataGridView1.TabIndex = 8
    $dataGridView1.Anchor = 15
    $dataGridView1.AutoSizeColumnsMode = 'AllCells'
    $dataGridView1.AllowUserToAddRows = $True
    $dataGridView1.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 70
    $dataGridView1.Location = $System_Drawing_Point
    $dataGridView1.AllowUserToOrderColumns = $True
    $dataGridView1.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridView1AutoSizeColumnsMode.AllCells
    #$datagridview1.add_DoubleClick($datagridview1_Click)  #Disabled because I don't want it there, but might be handy one day
    $datagridview1.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    $datagridview1.add_CurrentCellDirtyStateChanged($datagridview1_CurrentCellDirtyStateChanged)

    #*****************buttonOK Creation*****************
    $buttonOK.Anchor = 'Bottom, Left'
    $buttonOK.DialogResult = 'Ok'
    $buttonOK.Location = '13,565'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '105,23' #75
    $buttonOK.Text = '&Save and Close'
    $buttonOK.UseVisualStyleBackColor = $True
    $buttonOK.add_Click($buttonOK_Click)

    #*********btnQuit2 Creation -- see click action for more**********
    $btnQuit2.Anchor = 'Bottom, Left'
    $btnQuit2.DialogResult = 'Ok'
    $btnQuit2.Location = '125,565'
    $btnQuit2.Name = 'buttonOK'
    $btnQuit2.Size = '75,23'
    $btnQuit2.Text = '&Close'
    $btnQuit2.UseVisualStyleBackColor = $True
    $btnQuit2.add_Click($Quit2)

    #*********btnGenForm Creation -- see click action for more*********
    $btnGenForm.UseVisualStyleBackColor = $True
    $btnGenForm.Text = 'Refresh'
    $btnGenForm.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnGenForm.TabIndex = 1
    $btnGenForm.Name = 'btnGenForm'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btnGenForm.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 15
    $btnGenForm.Location = $System_Drawing_Point
    $btnGenForm.add_Click($GenForm_Click)

    $formPRT.ResumeLayout()   #** no idea what this does, but it does important stuff

    $InitialFormWindowState = $formPRT.WindowState         #** no idea what this does, but it does important stuff
    $formPRT.add_Load($Form_StateCorrection_Load)          #** no idea what this does, but it does important stuff
    $formPRT.ShowDialog()                        

}

#*********************************************************************************************
#**************SelectIP35 - EDITABLE SHEET THAT WILL NOT CLOSE PARENT SHEET*******************
#**************DOUBLE CLICK ID ON PARENT SHEET TRIGGERS THIS FORM*****************************
#*********************************************************************************************
function Show-SelectIP35{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form1 = New-Object 'System.Windows.Forms.Form'
    $datagridview1 = New-Object 'System.Windows.Forms.DataGridView'
    $buttonOK = New-Object 'System.Windows.Forms.Button'
    $btnQuit2 = New-Object 'System.Windows.Forms.Button'
    $datagridviewTest = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewIP = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewVers = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewDesc = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewNot = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewMod = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewModN = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewSerN = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPII = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewBrd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPrd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewMT = New-Object 'System.Windows.Forms.DataGridView'
    
    $labelIP = New-Object 'System.Windows.Forms.Label'
    $labelMAC = New-Object 'System.Windows.Forms.Label'
    $labelDVC = New-Object 'System.Windows.Forms.Label'
    $labelVers = New-Object 'System.Windows.Forms.Label'
    $labelDesc = New-Object 'System.Windows.Forms.Label'
    $labelNot = New-Object 'System.Windows.Forms.Label'
    $labelMod = New-Object 'System.Windows.Forms.Label'
    $labelModN = New-Object 'System.Windows.Forms.Label'
    $labelSerN = New-Object 'System.Windows.Forms.Label'
    $labelPII = New-Object 'System.Windows.Forms.Label'
    $labelBrd = New-Object 'System.Windows.Forms.Label'
    $labelPrd = New-Object 'System.Windows.Forms.Label'
    $labelMT = New-Object 'System.Windows.Forms.Label'
    

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=SSPI"

    $form1_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select id,PrinterName FROM PrinterAssets where id= '$selectedCell' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)
                
        $cmd2 = $conn.CreateCommand()
        $cmd2.CommandText = "Select id,AssetNumber FROM PrinterAssets where id= '$selectedCell' "
        $script:adapter2 = New-Object System.Data.SqlClient.SqlDataAdapter($cmd2)
        $dt2 = New-Object System.Data.DataTable
        $script:adapter2.Fill($dt2)
        $datagridviewTest.DataSource = $dt2
        $cmdBldr2 = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter2)

        $cmdIP = $conn.CreateCommand()
        $cmdIP.CommandText = "Select id,ReservationIP FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterIP = New-Object System.Data.SqlClient.SqlDataAdapter($cmdIP)
        $dtIP = New-Object System.Data.DataTable
        $script:adapterIP.Fill($dtIP)
        $datagridviewIP.DataSource = $dtIP
        $cmdBldrIP = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterIP)

        $cmdVers = $conn.CreateCommand()
        $cmdVers.CommandText = "Select id,Manufacturer FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterVers = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVers)
        $dtVers = New-Object System.Data.DataTable
        $script:adapterVers.Fill($dtVers)
        $datagridviewVers.DataSource = $dtVers
        $cmdBldrVers = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVers)

        $cmdDesc = $conn.CreateCommand()
        $cmdDesc.CommandText = "Select id,Model FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterDesc = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDesc)
        $dtDesc = New-Object System.Data.DataTable
        $script:adapterDesc.Fill($dtDesc)
        $datagridviewDesc.DataSource = $dtDesc
        $cmdBldrDesc = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDesc)

        $cmdNot = $conn.CreateCommand()
        $cmdNot.CommandText = "Select id,Notes FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterNot = New-Object System.Data.SqlClient.SqlDataAdapter($cmdNot)
        $dtNot = New-Object System.Data.DataTable
        $script:adapterNot.Fill($dtNot)
        $datagridviewNot.DataSource = $dtNot
        $cmdBldrNot = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterNot)

        $cmdMod = $conn.CreateCommand()
        $cmdMod.CommandText = "Select id,Location FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterMod = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMod)
        $dtMod = New-Object System.Data.DataTable
        $script:adapterMod.Fill($dtMod)
        $datagridviewMod.DataSource = $dtMod
        $cmdBldrMod = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMod)

        $cmdModN = $conn.CreateCommand()
        $cmdModN.CommandText = "Select id,ModelNumber FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterModN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdModN)
        $dtModN = New-Object System.Data.DataTable
        $script:adapterModN.Fill($dtModN)
        $datagridviewModN.DataSource = $dtModN
        $cmdBldrModN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterModN)

        $cmdSerN = $conn.CreateCommand()
        $cmdSerN.CommandText = "Select id,SerialNumber FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterSerN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdSerN)
        $dtSerN = New-Object System.Data.DataTable
        $script:adapterSerN.Fill($dtSerN)
        $datagridviewSerN.DataSource = $dtSerN
        $cmdBldrSerN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterSerN)

        $cmdPII = $conn.CreateCommand()
        $cmdPII.CommandText = "Select id,LeaseNumber FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterPII = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPII)
        $dtPII = New-Object System.Data.DataTable
        $script:adapterPII.Fill($dtPII)
        $datagridviewPII.DataSource = $dtPII
        $cmdBldrPII = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPII)

        $cmdBrd = $conn.CreateCommand()
        $cmdBrd.CommandText = "Select id,LeaseVendor FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterBrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdBrd)
        $dtBrd = New-Object System.Data.DataTable
        $script:adapterBrd.Fill($dtBrd)
        $datagridviewBrd.DataSource = $dtBrd
        $cmdBldrBrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterBrd)

        $cmdPrd = $conn.CreateCommand()
        $cmdPrd.CommandText = "Select id,PurchaseDate FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterPrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrd)
        $dtPrd = New-Object System.Data.DataTable
        $script:adapterPrd.Fill($dtPrd)
        $datagridviewPrd.DataSource = $dtPrd
        $cmdBldrPrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrd)

        $cmdMT = $conn.CreateCommand()
        $cmdMT.CommandText = "Select id,AssetType FROM PrinterAssets where id= '$selectedCell' "
        $script:adapterMT = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMT)
        $dtMT = New-Object System.Data.DataTable
        $script:adapterMT.Fill($dtMT)
        $datagridviewMT.DataSource = $dtMT
        $cmdBldrMT = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMT)

        
    }

    $buttonOK_Click = {
            [System.Windows.Forms.MessageBox]::Show("If you *Click* No, everything will blow up.. don't do it..", 'Data Changed', 'YesNo')
            $script:adapter.Update($datagridview1.DataSource)
            $script:adapter2.Update($datagridviewTest.DataSource)
            $script:adapterIP.Update($datagridviewIP.DataSource)
            $script:adapterVers.Update($dataGridViewVers.DataSource)
            $script:adapterDesc.Update($dataGridViewDesc.DataSource)
            $script:adapterNot.Update($dataGridViewNot.DataSource)
            $script:adapterMod.Update($dataGridViewMod.DataSource)
            $script:adapterModN.Update($dataGridViewModN.DataSource)
            $script:adapterSerN.Update($dataGridViewSerN.DataSource)
            $script:adapterPII.Update($dataGridViewPII.DataSource)
            $script:adapterBrd.Update($dataGridViewBrd.DataSource)
            $script:adapterPrd.Update($dataGridViewPrd.DataSource)
            $script:adapterMT.Update($dataGridViewMT.DataSource)
            
    }
    $datagridview1_CurrentCellDirtyStateChanged = {
        $script:DVGhasChanged = $true
    }
    $Form_StateCorrection_Load = {
        $form1.WindowState = $InitialFormWindowState
    }
    
    $form1.SuspendLayout()

    #form1
    
    $form1.Controls.Add($datagridview1)
    $form1.Controls.Add($datagridviewTest)
    $form1.Controls.Add($dataGridViewIP)
    $form1.Controls.Add($dataGridViewVers)
    $form1.Controls.Add($dataGridViewDesc)
    $form1.Controls.Add($dataGridViewNot)
    $form1.Controls.Add($dataGridViewMod)
    $form1.Controls.Add($dataGridViewModN)
    $form1.Controls.add($dataGridViewSerN)
    $form1.Controls.Add($dataGridViewPII)
    $form1.Controls.Add($dataGridViewBrd)
    $form1.Controls.Add($dataGridViewPrd)
    $form1.Controls.Add($dataGridViewMT)
 

    $form1.Controls.Add($buttonOK)
    #$form1.AcceptButton = $buttonOK
    $form1.Controls.Add($btnQuit2)
    $form1.Controls.Add($labelIP)
    $form1.Controls.Add($labelMAC)
    $form1.Controls.Add($labelDVC)
    $form1.Controls.add($labelVers)
    $form1.Controls.Add($labelDesc)
    $form1.Controls.Add($labelNot)
    $form1.Controls.Add($labelMod)
    $form1.Controls.Add($labelModN)
    $form1.Controls.Add($labelSerN)
    $form1.Controls.Add($labelPII)
    $form1.Controls.Add($labelBrd)
    $form1.Controls.Add($labelPrd)
    $form1.Controls.Add($labelMT)

    

    $labelDVC.Name = "Device"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDVC.size = $System_Drawing_Size
    $labelDVC.text = "ID   | Device"
    $labelDVC.Location = '5,2'

    $labelIP.Name = "IPLabel"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelIP.size = $System_Drawing_Size
    $labelIP.text = "ID   | IP Address"
    $labelIP.Location = '5,103'

    $labelMAC.Name = "MAC"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelMAC.size = $System_Drawing_Size
    $labelMAC.text = "ID   |  Asset Number"
    $labelMAC.Location = '5,53'
    
    $labelVers.Name = "Version"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVers.size = $System_Drawing_Size
    $labelVers.text = "ID   | Manufacturer"
    $labelVers.Location = '5,153'

    $labelDesc.Name = "Description"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDesc.size = $System_Drawing_Size
    $labelDesc.text = "ID   | Model"
    $labelDesc.Location = '5,203'

    $labelNot.Name = "Notes"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelNot.size = $System_Drawing_Size
    $labelNot.text = "ID   | Notes"
    $labelNot.Location = '5,268'

    $labelMod.Name = "Model"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMod.size = $System_Drawing_Size
    $labelMod.text = "ID   | Location"
    $labelMod.Location = '5,333'

    $labelModN.Name = "Model #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelModN.size = $System_Drawing_Size
    $labelModN.text = "ID   | Model Number"
    $labelModN.Location = '5,383'

    $labelSerN.Name = "Serial Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelSerN.size = $System_Drawing_Size
    $labelSerN.text = "ID   | Serial Number"
    $labelSerN.Location = '5,433'

    $labelPII.Name = "PII"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelPII.size = $System_Drawing_Size
    $labelPII.text = "ID   |  Lease Number"
    $labelPII.Location = '250,2'

    $labelBrd.Name = "Brand"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelBrd.size = $System_Drawing_Size
    $labelBrd.text = "ID   | Vendor"
    $labelBrd.Location = '250,53'

    $labelPrd.Name = "Product ID"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelPrd.size = $System_Drawing_Size
    $labelPrd.text = "ID   | Purchase Date"
    $labelPrd.Location = '250,103'

    $labelMT.Name = "MT"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMT.size = $System_Drawing_Size
    $labelMT.text = "ID   | Asset Type"
    $labelMT.Location = '250,153'


    #$form1.AcceptButton = $btnQuit2
    $form1.ClientSize = '455, 515' #725, 515
    $form1.FormBorderStyle = 'FixedDialog'
    $form1.BackColor = [System.Drawing.Color]::FromArgb(255,185,209,234)
    $form1.MaximizeBox = $False
    $form1.MinimizeBox = $True
    $form1.Name = 'form1'
    $form1.StartPosition = 'CenterScreen'
    $form1.Text = '***ID DBL CLICK***'
    $form1.KeyPreview = $True
    $form1.add_Load($form1_Load)

    #Device
    $dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridView1.DefaultCellStyle.BackColor = "White"
    $dataGridView1.BackgroundColor = "White"
    $dataGridView1.Name = 'dataGridView1'
    $dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridView1.ReadOnly = $False
    $dataGridView1.AllowUserToDeleteRows = $False
    $dataGridView1.RowHeadersVisible = $false
    $dataGridView1.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridView1.Size = $System_Drawing_Size
    $dataGridView1.TabIndex = 8
    $dataGridView1.Anchor = 15
    $dataGridView1.AutoSizeColumnsMode = 'AllCells'
    $dataGridView1.AllowUserToAddRows = $False
    $dataGridView1.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 15
    $dataGridView1.Location = $System_Drawing_Point
    $dataGridView1.AllowUserToOrderColumns = $True
    $dataGridView1.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridView1AutoSizeColumnsMode.AllCells
    $datagridview1.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    $datagridview1.add_CurrentCellDirtyStateChanged($datagridview1_CurrentCellDirtyStateChanged)
    
    #****Data MAC
    $dataGridViewTest.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewTest.DefaultCellStyle.BackColor = "White"
    $dataGridViewTest.BackgroundColor = "White"
    $dataGridViewTest.Name = 'dataGridView1'
    $dataGridViewTest.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewTest.ReadOnly = $False
    $dataGridViewTest.AllowUserToDeleteRows = $False
    $dataGridViewTest.RowHeadersVisible = $false
    $dataGridViewTest.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewTest.Size = $System_Drawing_Size
    $dataGridViewTest.TabIndex = 8
    $dataGridViewTest.Anchor = 15
    $dataGridViewTest.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewTest.AllowUserToAddRows = $false
    $dataGridViewTest.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 65
    $dataGridViewTest.Location = $System_Drawing_Point
    $dataGridViewTest.AllowUserToOrderColumns = $True
    $dataGridViewTest.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewTestAutoSizeColumnsMode.AllCells
    $datagridviewTest.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Data IP
    $dataGridViewIP.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewIP.DefaultCellStyle.BackColor = "White"
    $dataGridViewIP.BackgroundColor = "White"
    $dataGridViewIP.Name = 'dataGridViewIP'
    $dataGridViewIP.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewIP.ReadOnly = $False
    $dataGridViewIP.AllowUserToDeleteRows = $False
    $dataGridViewIP.ColumnHeadersVisible = $false
    $dataGridViewIP.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewIP.Size = $System_Drawing_Size
    $dataGridViewIP.TabIndex = 8
    $dataGridViewIP.Anchor = 15
    $dataGridViewIP.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewIP.AllowUserToAddRows = $false
    $dataGridViewIP.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 115
    $dataGridViewIP.Location = $System_Drawing_Point
    $dataGridViewIP.AllowUserToOrderColumns = $True
    $dataGridViewIP.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewIPAutoSizeColumnsMode.AllCells
    $datagridviewIP.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    
    #**Version
    $dataGridViewVers.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVers.DefaultCellStyle.BackColor = "White"
    $dataGridViewVers.BackgroundColor = "White"
    $dataGridViewVers.Name = 'dataGridViewVers'
    $dataGridViewVers.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVers.ReadOnly = $False
    $dataGridViewVers.AllowUserToDeleteRows = $False
    $dataGridViewVers.ColumnHeadersVisible = $false
    $dataGridViewVers.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVers.Size = $System_Drawing_Size
    $dataGridViewVers.TabIndex = 8
    $dataGridViewVers.Anchor = 15
    $dataGridViewVers.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVers.AllowUserToAddRows = $false
    $dataGridViewVers.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 165
    $dataGridViewVers.Location = $System_Drawing_Point
    $dataGridViewVers.AllowUserToOrderColumns = $True
    $dataGridViewVers.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVersAutoSizeColumnsMode.AllCells
    $datagridviewVers.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Description
    $dataGridViewDesc.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDesc.DefaultCellStyle.BackColor = "White"
    $dataGridViewDesc.BackgroundColor = "White"
    $dataGridViewDesc.Name = 'dataGridViewDesc'
    $dataGridViewDesc.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDesc.ReadOnly = $False
    $dataGridViewDesc.AllowUserToDeleteRows = $False
    $dataGridViewDesc.ColumnHeadersVisible = $false
    $dataGridViewDesc.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 445
    $System_Drawing_Size.Height = 40
    $dataGridViewDesc.Size = $System_Drawing_Size
    $dataGridViewDesc.TabIndex = 8
    $dataGridViewDesc.Anchor = 15
    $dataGridViewDesc.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDesc.AllowUserToAddRows = $false
    $dataGridViewDesc.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 215
    $dataGridViewDesc.Location = $System_Drawing_Point
    $dataGridViewDesc.AllowUserToOrderColumns = $True
    $dataGridViewDesc.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDescAutoSizeColumnsMode.AllCells
    $datagridviewDesc.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Notes
    $dataGridViewNot.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewNot.DefaultCellStyle.BackColor = "White"
    $dataGridViewNot.BackgroundColor = "White"
    $dataGridViewNot.Name = 'dataGridViewNot'
    $dataGridViewNot.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewNot.ReadOnly = $False
    $dataGridViewNot.AllowUserToDeleteRows = $False
    $dataGridViewNot.ColumnHeadersVisible = $false
    $dataGridViewNot.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 445
    $System_Drawing_Size.Height = 40
    $dataGridViewNot.Size = $System_Drawing_Size
    $dataGridViewNot.TabIndex = 8
    $dataGridViewNot.Anchor = 15
    $dataGridViewNot.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewNot.AllowUserToAddRows = $false
    $dataGridViewNot.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 280
    $dataGridViewNot.Location = $System_Drawing_Point
    $dataGridViewNot.AllowUserToOrderColumns = $True
    $dataGridViewNot.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewNotAutoSizeColumnsMode.AllCells
    $datagridviewNot.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Model
    $dataGridViewMod.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewMod.DefaultCellStyle.BackColor = "White"
    $dataGridViewMod.BackgroundColor = "White"
    $dataGridViewMod.Name = 'dataGridViewMod'
    $dataGridViewMod.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewMod.ReadOnly = $False
    $dataGridViewMod.AllowUserToDeleteRows = $False
    $dataGridViewMod.ColumnHeadersVisible = $false
    $dataGridViewMod.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewMod.Size = $System_Drawing_Size
    $dataGridViewMod.TabIndex = 8
    $dataGridViewMod.Anchor = 15
    $dataGridViewMod.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewMod.AllowUserToAddRows = $false
    $dataGridViewMod.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 345
    $dataGridViewMod.Location = $System_Drawing_Point
    $dataGridViewMod.AllowUserToOrderColumns = $True
    $dataGridViewMod.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewModAutoSizeColumnsMode.AllCells
    $datagridviewMod.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**ModelNumber
    $dataGridViewModN.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewModN.DefaultCellStyle.BackColor = "White"
    $dataGridViewModN.BackgroundColor = "White"
    $dataGridViewModN.Name = 'dataGridViewModN'
    $dataGridViewModN.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewModN.ReadOnly = $False
    $dataGridViewModN.AllowUserToDeleteRows = $False
    $dataGridViewModN.ColumnHeadersVisible = $false
    $dataGridViewModN.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewModN.Size = $System_Drawing_Size
    $dataGridViewModN.TabIndex = 8
    $dataGridViewModN.Anchor = 15
    $dataGridViewModN.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewModN.AllowUserToAddRows = $false
    $dataGridViewModN.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 395
    $dataGridViewModN.Location = $System_Drawing_Point
    $dataGridViewModN.AllowUserToOrderColumns = $True
    $dataGridViewModN.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewModNAutoSizeColumnsMode.AllCells
    $datagridviewModN.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })


    #**SerialNumber
    $dataGridViewSerN.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewSerN.DefaultCellStyle.BackColor = "White"
    $dataGridViewSerN.BackgroundColor = "White"
    $dataGridViewSerN.Name = 'dataGridViewModN'
    $dataGridViewSerN.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewSerN.ReadOnly = $False
    $dataGridViewSerN.AllowUserToDeleteRows = $False
    $dataGridViewSerN.ColumnHeadersVisible = $false
    $dataGridViewSerN.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewSerN.Size = $System_Drawing_Size
    $dataGridViewSerN.TabIndex = 8
    $dataGridViewSerN.Anchor = 15
    $dataGridViewSerN.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewSerN.AllowUserToAddRows = $false
    $dataGridViewSerN.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 445
    $dataGridViewSerN.Location = $System_Drawing_Point
    $dataGridViewSerN.AllowUserToOrderColumns = $True
    $dataGridViewSerN.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewSerNAutoSizeColumnsMode.AllCells
    $datagridviewSerN.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PII
    $dataGridViewPII.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPII.DefaultCellStyle.BackColor = "White"
    $dataGridViewPII.BackgroundColor = "White"
    $dataGridViewPII.Name = 'dataGridViewModN'
    $dataGridViewPII.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPII.ReadOnly = $False
    $dataGridViewPII.AllowUserToDeleteRows = $False
    $dataGridViewPII.ColumnHeadersVisible = $false
    $dataGridViewPII.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPII.Size = $System_Drawing_Size
    $dataGridViewPII.TabIndex = 8
    $dataGridViewPII.Anchor = 15
    $dataGridViewPII.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPII.AllowUserToAddRows = $false
    $dataGridViewPII.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 15
    $dataGridViewPII.Location = $System_Drawing_Point
    $dataGridViewPII.AllowUserToOrderColumns = $True
    $dataGridViewPII.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPIIAutoSizeColumnsMode.AllCells
    $datagridviewPII.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

     #**Brand
    $dataGridViewBrd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewBrd.DefaultCellStyle.BackColor = "White"
    $dataGridViewBrd.BackgroundColor = "White"
    $dataGridViewBrd.Name = 'dataGridViewBrd'
    $dataGridViewBrd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewBrd.ReadOnly = $False
    $dataGridViewBrd.AllowUserToDeleteRows = $False
    $dataGridViewBrd.ColumnHeadersVisible = $false
    $dataGridViewBrd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewBrd.Size = $System_Drawing_Size
    $dataGridViewBrd.TabIndex = 8
    $dataGridViewBrd.Anchor = 15
    $dataGridViewBrd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewBrd.AllowUserToAddRows = $false
    $dataGridViewBrd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 65
    $dataGridViewBrd.Location = $System_Drawing_Point
    $dataGridViewBrd.AllowUserToOrderColumns = $True
    $dataGridViewBrd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewBrdAutoSizeColumnsMode.AllCells
    $datagridviewBrd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PrdID
    $dataGridViewPrd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPrd.DefaultCellStyle.BackColor = "White"
    $dataGridViewPrd.BackgroundColor = "White"
    $dataGridViewPrd.Name = 'dataGridViewBrd'
    $dataGridViewPrd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPrd.ReadOnly = $False
    $dataGridViewPrd.AllowUserToDeleteRows = $False
    $dataGridViewPrd.ColumnHeadersVisible = $false
    $dataGridViewPrd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPrd.Size = $System_Drawing_Size
    $dataGridViewPrd.TabIndex = 8
    $dataGridViewPrd.Anchor = 15
    $dataGridViewPrd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPrd.AllowUserToAddRows = $false
    $dataGridViewPrd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 115
    $dataGridViewPrd.Location = $System_Drawing_Point
    $dataGridViewPrd.AllowUserToOrderColumns = $True
    $dataGridViewPrd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPrdAutoSizeColumnsMode.AllCells
    $datagridviewPrd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**MT
    $dataGridViewMT.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewMT.DefaultCellStyle.BackColor = "White"
    $dataGridViewMT.BackgroundColor = "White"
    $dataGridViewMT.Name = 'dataGridViewMT'
    $dataGridViewMT.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewMT.ReadOnly = $False
    $dataGridViewMT.AllowUserToDeleteRows = $False
    $dataGridViewMT.ColumnHeadersVisible = $false
    $dataGridViewMT.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewMT.Size = $System_Drawing_Size
    $dataGridViewMT.TabIndex = 8
    $dataGridViewMT.Anchor = 15
    $dataGridViewMT.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewMT.AllowUserToAddRows = $false
    $dataGridViewMT.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 165
    $dataGridViewMT.Location = $System_Drawing_Point
    $dataGridViewMT.AllowUserToOrderColumns = $True
    $dataGridViewMT.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewMTAutoSizeColumnsMode.AllCells
    $datagridviewMT.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    $buttonOK.Anchor = 'Bottom, Right'
    $buttonOK.DialogResult = 'Ok'
    $buttonOK.Location = '13,485'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '105,23' #75
    $buttonOK.Text = '&Save and Close'
    $buttonOK.UseVisualStyleBackColor = $True
    $buttonOK.add_Click($buttonOK_Click)

    $btnQuit2.Anchor = 'Bottom, Left'
    $btnQuit2.DialogResult = 'Ok'
    $btnQuit2.Location = '125,485'
    $btnQuit2.Name = 'buttonOK'
    $btnQuit2.Size = '75,23'
    $btnQuit2.Text = '&Close'
    $btnQuit2.UseVisualStyleBackColor = $True
    $btnQuit2.add_Click($Quit2)

    $form1.ResumeLayout()

    $InitialFormWindowState = $form1.WindowState
    $form1.add_Load($Form_StateCorrection_Load)
    $form1.ShowDialog()

}

#*********************************************************************************************
#**************SelectIP23 - EDITABLE SHEET THAT WILL NOT CLOSE PARENT SHEET*******************
#**************SEARCHING BY DEVICE TRIGGERS THIS FORM*****************************************
#*********************************************************************************************
function Show-SelectIP23{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form1 = New-Object 'System.Windows.Forms.Form'
    $datagridview1 = New-Object 'System.Windows.Forms.DataGridView'
    $buttonOK = New-Object 'System.Windows.Forms.Button'
    $btnQuit2 = New-Object 'System.Windows.Forms.Button'
    $datagridviewTest = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewIP = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewVers = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewDesc = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewNot = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewMod = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewModN = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewSerN = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPII = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewBrd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPrd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewMT = New-Object 'System.Windows.Forms.DataGridView'
    
    $labelIP = New-Object 'System.Windows.Forms.Label'
    $labelMAC = New-Object 'System.Windows.Forms.Label'
    $labelDVC = New-Object 'System.Windows.Forms.Label'
    $labelVers = New-Object 'System.Windows.Forms.Label'
    $labelDesc = New-Object 'System.Windows.Forms.Label'
    $labelNot = New-Object 'System.Windows.Forms.Label'
    $labelMod = New-Object 'System.Windows.Forms.Label'
    $labelModN = New-Object 'System.Windows.Forms.Label'
    $labelSerN = New-Object 'System.Windows.Forms.Label'
    $labelPII = New-Object 'System.Windows.Forms.Label'
    $labelBrd = New-Object 'System.Windows.Forms.Label'
    $labelPrd = New-Object 'System.Windows.Forms.Label'
    $labelMT = New-Object 'System.Windows.Forms.Label'

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=SSPI"

    $form1_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select id,PrinterName FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)

        
        $cmd2 = $conn.CreateCommand()
        $cmd2.CommandText = "Select id,AssetNumber FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapter2 = New-Object System.Data.SqlClient.SqlDataAdapter($cmd2)
        $dt2 = New-Object System.Data.DataTable
        $script:adapter2.Fill($dt2)
        $datagridviewTest.DataSource = $dt2
        $cmdBldr2 = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter2)

        $cmdIP = $conn.CreateCommand()
        $cmdIP.CommandText = "Select id,ReservationIP FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterIP = New-Object System.Data.SqlClient.SqlDataAdapter($cmdIP)
        $dtIP = New-Object System.Data.DataTable
        $script:adapterIP.Fill($dtIP)
        $datagridviewIP.DataSource = $dtIP
        $cmdBldrIP = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterIP)

        $cmdVers = $conn.CreateCommand()
        $cmdVers.CommandText = "Select id,Manufacturer FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterVers = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVers)
        $dtVers = New-Object System.Data.DataTable
        $script:adapterVers.Fill($dtVers)
        $datagridviewVers.DataSource = $dtVers
        $cmdBldrVers = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVers)

        $cmdDesc = $conn.CreateCommand()
        $cmdDesc.CommandText = "Select id,Model FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterDesc = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDesc)
        $dtDesc = New-Object System.Data.DataTable
        $script:adapterDesc.Fill($dtDesc)
        $datagridviewDesc.DataSource = $dtDesc
        $cmdBldrDesc = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDesc)

        $cmdNot = $conn.CreateCommand()
        $cmdNot.CommandText = "Select id,Notes FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterNot = New-Object System.Data.SqlClient.SqlDataAdapter($cmdNot)
        $dtNot = New-Object System.Data.DataTable
        $script:adapterNot.Fill($dtNot)
        $datagridviewNot.DataSource = $dtNot
        $cmdBldrNot = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterNot)

        $cmdMod = $conn.CreateCommand()
        $cmdMod.CommandText = "Select id,Location FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterMod = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMod)
        $dtMod = New-Object System.Data.DataTable
        $script:adapterMod.Fill($dtMod)
        $datagridviewMod.DataSource = $dtMod
        $cmdBldrMod = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMod)

        $cmdModN = $conn.CreateCommand()
        $cmdModN.CommandText = "Select id,ModelNumber FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterModN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdModN)
        $dtModN = New-Object System.Data.DataTable
        $script:adapterModN.Fill($dtModN)
        $datagridviewModN.DataSource = $dtModN
        $cmdBldrModN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterModN)

        $cmdSerN = $conn.CreateCommand()
        $cmdSerN.CommandText = "Select id,SerialNumber FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterSerN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdSerN)
        $dtSerN = New-Object System.Data.DataTable
        $script:adapterSerN.Fill($dtSerN)
        $datagridviewSerN.DataSource = $dtSerN
        $cmdBldrSerN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterSerN)

        $cmdPII = $conn.CreateCommand()
        $cmdPII.CommandText = "Select id,LeaseNumber FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterPII = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPII)
        $dtPII = New-Object System.Data.DataTable
        $script:adapterPII.Fill($dtPII)
        $datagridviewPII.DataSource = $dtPII
        $cmdBldrPII = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPII)

        $cmdBrd = $conn.CreateCommand()
        $cmdBrd.CommandText = "Select id,LeaseVendor FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterBrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdBrd)
        $dtBrd = New-Object System.Data.DataTable
        $script:adapterBrd.Fill($dtBrd)
        $datagridviewBrd.DataSource = $dtBrd
        $cmdBldrBrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterBrd)

        $cmdPrd = $conn.CreateCommand()
        $cmdPrd.CommandText = "Select id,PurchaseDate FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterPrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrd)
        $dtPrd = New-Object System.Data.DataTable
        $script:adapterPrd.Fill($dtPrd)
        $datagridviewPrd.DataSource = $dtPrd
        $cmdBldrPrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrd)

        $cmdMT = $conn.CreateCommand()
        $cmdMT.CommandText = "Select id,AssetType FROM PrinterAssets where AssetNumber= '$TEST' "
        $script:adapterMT = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMT)
        $dtMT = New-Object System.Data.DataTable
        $script:adapterMT.Fill($dtMT)
        $datagridviewMT.DataSource = $dtMT
        $cmdBldrMT = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMT)
    }

    $buttonOK_Click = {
            [System.Windows.Forms.MessageBox]::Show("Self-destruct initiated.. Just kidding.. *Click* Yes to finish saving changes", 'Data Changed', 'YesNo')
            $script:adapter.Update($datagridview1.DataSource)
            $script:adapter2.Update($datagridviewTest.DataSource)
            $script:adapterIP.Update($datagridviewIP.DataSource)
            $script:adapterVers.Update($dataGridViewVers.DataSource)
            $script:adapterDesc.Update($dataGridViewDesc.DataSource)
            $script:adapterNot.Update($dataGridViewNot.DataSource)
            $script:adapterMod.Update($dataGridViewMod.DataSource)
            $script:adapterModN.Update($dataGridViewModN.DataSource)
            $script:adapterSerN.Update($dataGridViewSerN.DataSource)
            $script:adapterPII.Update($dataGridViewPII.DataSource)
            $script:adapterBrd.Update($dataGridViewBrd.DataSource)
            $script:adapterPrd.Update($dataGridViewPrd.DataSource)
            $script:adapterMT.Update($dataGridViewMT.DataSource)
    }
    $datagridview1_CurrentCellDirtyStateChanged = {
        $script:DVGhasChanged = $true
    }
    $Form_StateCorrection_Load = {
        $form1.WindowState = $InitialFormWindowState
    }
    
    $form1.SuspendLayout()

    #form1
    
    $form1.Controls.Add($datagridview1)
    $form1.Controls.Add($datagridviewTest)
    $form1.Controls.Add($dataGridViewIP)
    $form1.Controls.Add($dataGridViewVers)
    $form1.Controls.Add($dataGridViewDesc)
    $form1.Controls.Add($dataGridViewNot)
    $form1.Controls.Add($dataGridViewMod)
    $form1.Controls.Add($dataGridViewModN)
    $form1.Controls.add($dataGridViewSerN)
    $form1.Controls.Add($dataGridViewPII)
    $form1.Controls.Add($dataGridViewBrd)
    $form1.Controls.Add($dataGridViewPrd)
    $form1.Controls.Add($dataGridViewMT)

    $form1.Controls.Add($buttonOK)
    #$form1.AcceptButton = $buttonOK
    $form1.Controls.Add($btnQuit2)
    $form1.Controls.Add($labelIP)
    $form1.Controls.Add($labelMAC)
    $form1.Controls.Add($labelDVC)
    $form1.Controls.add($labelVers)
    $form1.Controls.Add($labelDesc)
    $form1.Controls.Add($labelNot)
    $form1.Controls.Add($labelMod)
    $form1.Controls.Add($labelModN)
    $form1.Controls.Add($labelSerN)
    $form1.Controls.Add($labelPII)
    $form1.Controls.Add($labelBrd)
    $form1.Controls.Add($labelPrd)
    $form1.Controls.Add($labelMT)

    $labelDVC.Name = "Device"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDVC.size = $System_Drawing_Size
    $labelDVC.text = "ID   | Device"
    $labelDVC.Location = '5,2'

    $labelIP.Name = "IPLabel"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelIP.size = $System_Drawing_Size
    $labelIP.text = "ID   | IP Address"
    $labelIP.Location = '5,103'

    $labelMAC.Name = "MAC"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelMAC.size = $System_Drawing_Size
    $labelMAC.text = "ID   |  Asset Number"
    $labelMAC.Location = '5,53'
    
    $labelVers.Name = "Version"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVers.size = $System_Drawing_Size
    $labelVers.text = "ID   | Manufacturer"
    $labelVers.Location = '5,153'

    $labelDesc.Name = "Description"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDesc.size = $System_Drawing_Size
    $labelDesc.text = "ID   | Model"
    $labelDesc.Location = '5,203'

    $labelNot.Name = "Notes"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelNot.size = $System_Drawing_Size
    $labelNot.text = "ID   | Notes"
    $labelNot.Location = '5,268'

    $labelMod.Name = "Model"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMod.size = $System_Drawing_Size
    $labelMod.text = "ID   | Location"
    $labelMod.Location = '5,333'

    $labelModN.Name = "Model #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelModN.size = $System_Drawing_Size
    $labelModN.text = "ID   | Model Number"
    $labelModN.Location = '5,383'

    $labelSerN.Name = "Serial Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelSerN.size = $System_Drawing_Size
    $labelSerN.text = "ID   | Serial Number"
    $labelSerN.Location = '5,433'

    $labelPII.Name = "PII"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelPII.size = $System_Drawing_Size
    $labelPII.text = "ID   |  Lease Number"
    $labelPII.Location = '250,2'

    $labelBrd.Name = "Brand"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelBrd.size = $System_Drawing_Size
    $labelBrd.text = "ID   | Vendor"
    $labelBrd.Location = '250,53'

    $labelPrd.Name = "Product ID"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelPrd.size = $System_Drawing_Size
    $labelPrd.text = "ID   | Purchase Date"
    $labelPrd.Location = '250,103'

    $labelMT.Name = "MT"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMT.size = $System_Drawing_Size
    $labelMT.text = "ID   | Asset Type"
    $labelMT.Location = '250,153'

    #$form1.AcceptButton = $btnQuit2
    $form1.ClientSize = '455, 515' #900,800
    $form1.FormBorderStyle = 'FixedDialog'
    $form1.BackColor = [System.Drawing.Color]::FromArgb(255,185,209,234)
    $form1.MaximizeBox = $False
    $form1.MinimizeBox = $True
    $form1.Name = 'form1'
    $form1.StartPosition = 'CenterScreen'
    $form1.Text = '***Device Search***'
    $form1.KeyPreview = $True
    $form1.add_Load($form1_Load)

    #Device
    $dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridView1.DefaultCellStyle.BackColor = "White"
    $dataGridView1.BackgroundColor = "White"
    $dataGridView1.Name = 'dataGridView1'
    $dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridView1.ReadOnly = $False
    $dataGridView1.AllowUserToDeleteRows = $False
    $dataGridView1.RowHeadersVisible = $false
    $dataGridView1.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridView1.Size = $System_Drawing_Size
    $dataGridView1.TabIndex = 8
    $dataGridView1.Anchor = 15
    $dataGridView1.AutoSizeColumnsMode = 'AllCells'
    $dataGridView1.AllowUserToAddRows = $False
    $dataGridView1.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 15
    $dataGridView1.Location = $System_Drawing_Point
    $dataGridView1.AllowUserToOrderColumns = $True
    $dataGridView1.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridView1AutoSizeColumnsMode.AllCells
    $datagridview1.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    $datagridview1.add_CurrentCellDirtyStateChanged($datagridview1_CurrentCellDirtyStateChanged)
    
    #****Data MAC
    $dataGridViewTest.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewTest.DefaultCellStyle.BackColor = "White"
    $dataGridViewTest.BackgroundColor = "White"
    $dataGridViewTest.Name = 'dataGridView1'
    $dataGridViewTest.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewTest.ReadOnly = $False
    $dataGridViewTest.AllowUserToDeleteRows = $False
    $dataGridViewTest.RowHeadersVisible = $false
    $dataGridViewTest.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewTest.Size = $System_Drawing_Size
    $dataGridViewTest.TabIndex = 8
    $dataGridViewTest.Anchor = 15
    $dataGridViewTest.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewTest.AllowUserToAddRows = $false
    $dataGridViewTest.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 65
    $dataGridViewTest.Location = $System_Drawing_Point
    $dataGridViewTest.AllowUserToOrderColumns = $True
    $dataGridViewTest.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewTestAutoSizeColumnsMode.AllCells
    $datagridviewTest.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Data IP
    $dataGridViewIP.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewIP.DefaultCellStyle.BackColor = "White"
    $dataGridViewIP.BackgroundColor = "White"
    $dataGridViewIP.Name = 'dataGridViewIP'
    $dataGridViewIP.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewIP.ReadOnly = $False
    $dataGridViewIP.AllowUserToDeleteRows = $False
    $dataGridViewIP.ColumnHeadersVisible = $false
    $dataGridViewIP.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewIP.Size = $System_Drawing_Size
    $dataGridViewIP.TabIndex = 8
    $dataGridViewIP.Anchor = 15
    $dataGridViewIP.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewIP.AllowUserToAddRows = $false
    $dataGridViewIP.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 115
    $dataGridViewIP.Location = $System_Drawing_Point
    $dataGridViewIP.AllowUserToOrderColumns = $True
    $dataGridViewIP.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewIPAutoSizeColumnsMode.AllCells
    $datagridviewIP.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    
    #**Version
    $dataGridViewVers.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVers.DefaultCellStyle.BackColor = "White"
    $dataGridViewVers.BackgroundColor = "White"
    $dataGridViewVers.Name = 'dataGridViewVers'
    $dataGridViewVers.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVers.ReadOnly = $False
    $dataGridViewVers.AllowUserToDeleteRows = $False
    $dataGridViewVers.ColumnHeadersVisible = $false
    $dataGridViewVers.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVers.Size = $System_Drawing_Size
    $dataGridViewVers.TabIndex = 8
    $dataGridViewVers.Anchor = 15
    $dataGridViewVers.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVers.AllowUserToAddRows = $false
    $dataGridViewVers.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 165
    $dataGridViewVers.Location = $System_Drawing_Point
    $dataGridViewVers.AllowUserToOrderColumns = $True
    $dataGridViewVers.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVersAutoSizeColumnsMode.AllCells
    $datagridviewVers.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Description
    $dataGridViewDesc.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDesc.DefaultCellStyle.BackColor = "White"
    $dataGridViewDesc.BackgroundColor = "White"
    $dataGridViewDesc.Name = 'dataGridViewDesc'
    $dataGridViewDesc.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDesc.ReadOnly = $False
    $dataGridViewDesc.AllowUserToDeleteRows = $False
    $dataGridViewDesc.ColumnHeadersVisible = $false
    $dataGridViewDesc.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 445
    $System_Drawing_Size.Height = 40
    $dataGridViewDesc.Size = $System_Drawing_Size
    $dataGridViewDesc.TabIndex = 8
    $dataGridViewDesc.Anchor = 15
    $dataGridViewDesc.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDesc.AllowUserToAddRows = $false
    $dataGridViewDesc.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 215
    $dataGridViewDesc.Location = $System_Drawing_Point
    $dataGridViewDesc.AllowUserToOrderColumns = $True
    $dataGridViewDesc.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDescAutoSizeColumnsMode.AllCells
    $datagridviewDesc.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Notes
    $dataGridViewNot.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewNot.DefaultCellStyle.BackColor = "White"
    $dataGridViewNot.BackgroundColor = "White"
    $dataGridViewNot.Name = 'dataGridViewNot'
    $dataGridViewNot.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewNot.ReadOnly = $False
    $dataGridViewNot.AllowUserToDeleteRows = $False
    $dataGridViewNot.ColumnHeadersVisible = $false
    $dataGridViewNot.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 445
    $System_Drawing_Size.Height = 40
    $dataGridViewNot.Size = $System_Drawing_Size
    $dataGridViewNot.TabIndex = 8
    $dataGridViewNot.Anchor = 15
    $dataGridViewNot.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewNot.AllowUserToAddRows = $false
    $dataGridViewNot.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 280
    $dataGridViewNot.Location = $System_Drawing_Point
    $dataGridViewNot.AllowUserToOrderColumns = $True
    $dataGridViewNot.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewNotAutoSizeColumnsMode.AllCells
    $datagridviewNot.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Model
    $dataGridViewMod.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewMod.DefaultCellStyle.BackColor = "White"
    $dataGridViewMod.BackgroundColor = "White"
    $dataGridViewMod.Name = 'dataGridViewMod'
    $dataGridViewMod.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewMod.ReadOnly = $False
    $dataGridViewMod.AllowUserToDeleteRows = $False
    $dataGridViewMod.ColumnHeadersVisible = $false
    $dataGridViewMod.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewMod.Size = $System_Drawing_Size
    $dataGridViewMod.TabIndex = 8
    $dataGridViewMod.Anchor = 15
    $dataGridViewMod.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewMod.AllowUserToAddRows = $false
    $dataGridViewMod.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 345
    $dataGridViewMod.Location = $System_Drawing_Point
    $dataGridViewMod.AllowUserToOrderColumns = $True
    $dataGridViewMod.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewModAutoSizeColumnsMode.AllCells
    $datagridviewMod.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**ModelNumber
    $dataGridViewModN.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewModN.DefaultCellStyle.BackColor = "White"
    $dataGridViewModN.BackgroundColor = "White"
    $dataGridViewModN.Name = 'dataGridViewModN'
    $dataGridViewModN.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewModN.ReadOnly = $False
    $dataGridViewModN.AllowUserToDeleteRows = $False
    $dataGridViewModN.ColumnHeadersVisible = $false
    $dataGridViewModN.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewModN.Size = $System_Drawing_Size
    $dataGridViewModN.TabIndex = 8
    $dataGridViewModN.Anchor = 15
    $dataGridViewModN.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewModN.AllowUserToAddRows = $false
    $dataGridViewModN.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 395
    $dataGridViewModN.Location = $System_Drawing_Point
    $dataGridViewModN.AllowUserToOrderColumns = $True
    $dataGridViewModN.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewModNAutoSizeColumnsMode.AllCells
    $datagridviewModN.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })


    #**SerialNumber
    $dataGridViewSerN.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewSerN.DefaultCellStyle.BackColor = "White"
    $dataGridViewSerN.BackgroundColor = "White"
    $dataGridViewSerN.Name = 'dataGridViewModN'
    $dataGridViewSerN.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewSerN.ReadOnly = $False
    $dataGridViewSerN.AllowUserToDeleteRows = $False
    $dataGridViewSerN.ColumnHeadersVisible = $false
    $dataGridViewSerN.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewSerN.Size = $System_Drawing_Size
    $dataGridViewSerN.TabIndex = 8
    $dataGridViewSerN.Anchor = 15
    $dataGridViewSerN.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewSerN.AllowUserToAddRows = $false
    $dataGridViewSerN.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 445
    $dataGridViewSerN.Location = $System_Drawing_Point
    $dataGridViewSerN.AllowUserToOrderColumns = $True
    $dataGridViewSerN.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewSerNAutoSizeColumnsMode.AllCells
    $datagridviewSerN.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PII
    $dataGridViewPII.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPII.DefaultCellStyle.BackColor = "White"
    $dataGridViewPII.BackgroundColor = "White"
    $dataGridViewPII.Name = 'dataGridViewModN'
    $dataGridViewPII.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPII.ReadOnly = $False
    $dataGridViewPII.AllowUserToDeleteRows = $False
    $dataGridViewPII.ColumnHeadersVisible = $false
    $dataGridViewPII.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPII.Size = $System_Drawing_Size
    $dataGridViewPII.TabIndex = 8
    $dataGridViewPII.Anchor = 15
    $dataGridViewPII.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPII.AllowUserToAddRows = $false
    $dataGridViewPII.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 15
    $dataGridViewPII.Location = $System_Drawing_Point
    $dataGridViewPII.AllowUserToOrderColumns = $True
    $dataGridViewPII.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPIIAutoSizeColumnsMode.AllCells
    $datagridviewPII.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Brand
    $dataGridViewBrd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewBrd.DefaultCellStyle.BackColor = "White"
    $dataGridViewBrd.BackgroundColor = "White"
    $dataGridViewBrd.Name = 'dataGridViewBrd'
    $dataGridViewBrd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewBrd.ReadOnly = $False
    $dataGridViewBrd.AllowUserToDeleteRows = $False
    $dataGridViewBrd.ColumnHeadersVisible = $false
    $dataGridViewBrd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewBrd.Size = $System_Drawing_Size
    $dataGridViewBrd.TabIndex = 8
    $dataGridViewBrd.Anchor = 15
    $dataGridViewBrd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewBrd.AllowUserToAddRows = $false
    $dataGridViewBrd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 65
    $dataGridViewBrd.Location = $System_Drawing_Point
    $dataGridViewBrd.AllowUserToOrderColumns = $True
    $dataGridViewBrd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewBrdAutoSizeColumnsMode.AllCells
    $datagridviewBrd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PrdID
    $dataGridViewPrd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPrd.DefaultCellStyle.BackColor = "White"
    $dataGridViewPrd.BackgroundColor = "White"
    $dataGridViewPrd.Name = 'dataGridViewBrd'
    $dataGridViewPrd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPrd.ReadOnly = $False
    $dataGridViewPrd.AllowUserToDeleteRows = $False
    $dataGridViewPrd.ColumnHeadersVisible = $false
    $dataGridViewPrd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPrd.Size = $System_Drawing_Size
    $dataGridViewPrd.TabIndex = 8
    $dataGridViewPrd.Anchor = 15
    $dataGridViewPrd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPrd.AllowUserToAddRows = $false
    $dataGridViewPrd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 115
    $dataGridViewPrd.Location = $System_Drawing_Point
    $dataGridViewPrd.AllowUserToOrderColumns = $True
    $dataGridViewPrd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPrdAutoSizeColumnsMode.AllCells
    $datagridviewPrd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**MT
    $dataGridViewMT.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewMT.DefaultCellStyle.BackColor = "White"
    $dataGridViewMT.BackgroundColor = "White"
    $dataGridViewMT.Name = 'dataGridViewMT'
    $dataGridViewMT.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewMT.ReadOnly = $False
    $dataGridViewMT.AllowUserToDeleteRows = $False
    $dataGridViewMT.ColumnHeadersVisible = $false
    $dataGridViewMT.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewMT.Size = $System_Drawing_Size
    $dataGridViewMT.TabIndex = 8
    $dataGridViewMT.Anchor = 15
    $dataGridViewMT.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewMT.AllowUserToAddRows = $false
    $dataGridViewMT.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 165
    $dataGridViewMT.Location = $System_Drawing_Point
    $dataGridViewMT.AllowUserToOrderColumns = $True
    $dataGridViewMT.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewMTAutoSizeColumnsMode.AllCells
    $datagridviewMT.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })


    $buttonOK.Anchor = 'Bottom, Right'
    $buttonOK.DialogResult = 'Ok'
    $buttonOK.Location = '13,485'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '105,23' #75
    $buttonOK.Text = '&Save and Close'
    $buttonOK.UseVisualStyleBackColor = $True
    $buttonOK.add_Click($buttonOK_Click)

    $btnQuit2.Anchor = 'Bottom, Left'
    $btnQuit2.DialogResult = 'Ok'
    $btnQuit2.Location = '125,485'
    $btnQuit2.Name = 'buttonOK'
    $btnQuit2.Size = '75,23'
    $btnQuit2.Text = '&Close'
    $btnQuit2.UseVisualStyleBackColor = $True
    $btnQuit2.add_Click($Quit2)

    $form1.ResumeLayout()

    $InitialFormWindowState = $form1.WindowState
    $form1.add_Load($Form_StateCorrection_Load)
    $form1.ShowDialog()

}

#*********************************************************************************************
#****************SelectIP - EDITABLE SHEET THAT WILL NOT CLOSE PARENT SHEET*******************
#****************SEARCHING BY IPADDRESS TRIGGERS THIS FORM************************************
#*********************************************************************************************
function Show-SelectIP{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form1 = New-Object 'System.Windows.Forms.Form'
    $datagridview1 = New-Object 'System.Windows.Forms.DataGridView'
    $buttonOK = New-Object 'System.Windows.Forms.Button'
    $btnQuit2 = New-Object 'System.Windows.Forms.Button'
    $datagridviewTest = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewIP = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewVers = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewDesc = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewNot = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewMod = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewModN = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewSerN = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPII = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewBrd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPrd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewMT = New-Object 'System.Windows.Forms.DataGridView'

    $labelIP = New-Object 'System.Windows.Forms.Label'
    $labelMAC = New-Object 'System.Windows.Forms.Label'
    $labelDVC = New-Object 'System.Windows.Forms.Label'
    $labelVers = New-Object 'System.Windows.Forms.Label'
    $labelDesc = New-Object 'System.Windows.Forms.Label'
    $labelNot = New-Object 'System.Windows.Forms.Label'
    $labelMod = New-Object 'System.Windows.Forms.Label'
    $labelModN = New-Object 'System.Windows.Forms.Label'
    $labelSerN = New-Object 'System.Windows.Forms.Label'
    $labelPII = New-Object 'System.Windows.Forms.Label'
    $labelBrd = New-Object 'System.Windows.Forms.Label'
    $labelPrd = New-Object 'System.Windows.Forms.Label'
    $labelMT = New-Object 'System.Windows.Forms.Label'

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=SSPI"

    $form1_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select id,PrinterName FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)

        
        $cmd2 = $conn.CreateCommand()
        $cmd2.CommandText = "Select id,AssetNumber FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapter2 = New-Object System.Data.SqlClient.SqlDataAdapter($cmd2)
        $dt2 = New-Object System.Data.DataTable
        $script:adapter2.Fill($dt2)
        $datagridviewTest.DataSource = $dt2
        $cmdBldr2 = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter2)

        $cmdIP = $conn.CreateCommand()
        $cmdIP.CommandText = "Select id,ReservationIP FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterIP = New-Object System.Data.SqlClient.SqlDataAdapter($cmdIP)
        $dtIP = New-Object System.Data.DataTable
        $script:adapterIP.Fill($dtIP)
        $datagridviewIP.DataSource = $dtIP
        $cmdBldrIP = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterIP)

        $cmdVers = $conn.CreateCommand()
        $cmdVers.CommandText = "Select id,Manufacturer FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterVers = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVers)
        $dtVers = New-Object System.Data.DataTable
        $script:adapterVers.Fill($dtVers)
        $datagridviewVers.DataSource = $dtVers
        $cmdBldrVers = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVers)

        $cmdDesc = $conn.CreateCommand()
        $cmdDesc.CommandText = "Select id,Model FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterDesc = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDesc)
        $dtDesc = New-Object System.Data.DataTable
        $script:adapterDesc.Fill($dtDesc)
        $datagridviewDesc.DataSource = $dtDesc
        $cmdBldrDesc = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDesc)

        $cmdNot = $conn.CreateCommand()
        $cmdNot.CommandText = "Select id,Notes FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterNot = New-Object System.Data.SqlClient.SqlDataAdapter($cmdNot)
        $dtNot = New-Object System.Data.DataTable
        $script:adapterNot.Fill($dtNot)
        $datagridviewNot.DataSource = $dtNot
        $cmdBldrNot = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterNot)

        $cmdMod = $conn.CreateCommand()
        $cmdMod.CommandText = "Select id,Location FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterMod = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMod)
        $dtMod = New-Object System.Data.DataTable
        $script:adapterMod.Fill($dtMod)
        $datagridviewMod.DataSource = $dtMod
        $cmdBldrMod = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMod)

        $cmdModN = $conn.CreateCommand()
        $cmdModN.CommandText = "Select id,ModelNumber FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterModN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdModN)
        $dtModN = New-Object System.Data.DataTable
        $script:adapterModN.Fill($dtModN)
        $datagridviewModN.DataSource = $dtModN
        $cmdBldrModN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterModN)

        $cmdSerN = $conn.CreateCommand()
        $cmdSerN.CommandText = "Select id,SerialNumber FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterSerN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdSerN)
        $dtSerN = New-Object System.Data.DataTable
        $script:adapterSerN.Fill($dtSerN)
        $datagridviewSerN.DataSource = $dtSerN
        $cmdBldrSerN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterSerN)

        $cmdPII = $conn.CreateCommand()
        $cmdPII.CommandText = "Select id,LeaseNumber FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterPII = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPII)
        $dtPII = New-Object System.Data.DataTable
        $script:adapterPII.Fill($dtPII)
        $datagridviewPII.DataSource = $dtPII
        $cmdBldrPII = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPII)

        $cmdBrd = $conn.CreateCommand()
        $cmdBrd.CommandText = "Select id,LeaseVendor FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterBrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdBrd)
        $dtBrd = New-Object System.Data.DataTable
        $script:adapterBrd.Fill($dtBrd)
        $datagridviewBrd.DataSource = $dtBrd
        $cmdBldrBrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterBrd)

        $cmdPrd = $conn.CreateCommand()
        $cmdPrd.CommandText = "Select id,PurchaseDate FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterPrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrd)
        $dtPrd = New-Object System.Data.DataTable
        $script:adapterPrd.Fill($dtPrd)
        $datagridviewPrd.DataSource = $dtPrd
        $cmdBldrPrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrd)

        $cmdMT = $conn.CreateCommand()
        $cmdMT.CommandText = "Select id,AssetType FROM PrinterAssets where ReservationIP= '$userText' "
        $script:adapterMT = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMT)
        $dtMT = New-Object System.Data.DataTable
        $script:adapterMT.Fill($dtMT)
        $datagridviewMT.DataSource = $dtMT
        $cmdBldrMT = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMT)
    }

    $buttonOK_Click = {
            [System.Windows.Forms.MessageBox]::Show("Changes have been saved. Clicking No won't save you now", 'Data Changed', 'YesNo')
            $script:adapter.Update($datagridview1.DataSource)
            $script:adapter2.Update($datagridviewTest.DataSource)
            $script:adapterIP.Update($datagridviewIP.DataSource)
            $script:adapterVers.Update($dataGridViewVers.DataSource)
            $script:adapterDesc.Update($dataGridViewDesc.DataSource)
            $script:adapterNot.Update($dataGridViewNot.DataSource)
            $script:adapterMod.Update($dataGridViewMod.DataSource)
            $script:adapterModN.Update($dataGridViewModN.DataSource)
            $script:adapterSerN.Update($dataGridViewSerN.DataSource)
            $script:adapterPII.Update($dataGridViewPII.DataSource)
            $script:adapterBrd.Update($dataGridViewBrd.DataSource)
            $script:adapterPrd.Update($dataGridViewPrd.DataSource)
            $script:adapterMT.Update($dataGridViewMT.DataSource)
    }
    $datagridview1_CurrentCellDirtyStateChanged = {
        $script:DVGhasChanged = $true
    }
    $Form_StateCorrection_Load = {
        $form1.WindowState = $InitialFormWindowState
    }
    
    $form1.SuspendLayout()

    #form1
    
    $form1.Controls.Add($datagridview1)
    $form1.Controls.Add($datagridviewTest)
    $form1.Controls.Add($dataGridViewIP)
    $form1.Controls.Add($dataGridViewVers)
    $form1.Controls.Add($dataGridViewDesc)
    $form1.Controls.Add($dataGridViewNot)
    $form1.Controls.Add($dataGridViewMod)
    $form1.Controls.Add($dataGridViewModN)
    $form1.Controls.add($dataGridViewSerN)
    $form1.Controls.Add($dataGridViewPII)
    $form1.Controls.Add($dataGridViewBrd)
    $form1.Controls.Add($dataGridViewPrd)
    $form1.Controls.Add($dataGridViewMT)

    $form1.Controls.Add($buttonOK)
    #$form1.AcceptButton = $buttonOK
    $form1.Controls.Add($btnQuit2)
    $form1.Controls.Add($labelIP)
    $form1.Controls.Add($labelMAC)
    $form1.Controls.Add($labelDVC)
    $form1.Controls.add($labelVers)
    $form1.Controls.Add($labelDesc)
    $form1.Controls.Add($labelNot)
    $form1.Controls.Add($labelMod)
    $form1.Controls.Add($labelModN)
    $form1.Controls.Add($labelSerN)
    $form1.Controls.Add($labelPII)
    $form1.Controls.Add($labelBrd)
    $form1.Controls.Add($labelPrd)
    $form1.Controls.Add($labelMT)

    $labelDVC.Name = "Device"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDVC.size = $System_Drawing_Size
    $labelDVC.text = "ID   | Device"
    $labelDVC.Location = '5,2'

    $labelIP.Name = "IPLabel"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelIP.size = $System_Drawing_Size
    $labelIP.text = "ID   | IP Address"
    $labelIP.Location = '5,103'

    $labelMAC.Name = "MAC"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelMAC.size = $System_Drawing_Size
    $labelMAC.text = "ID   |  Asset Number"
    $labelMAC.Location = '5,53'
    
    $labelVers.Name = "Version"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVers.size = $System_Drawing_Size
    $labelVers.text = "ID   | Manufacturer"
    $labelVers.Location = '5,153'

    $labelDesc.Name = "Description"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDesc.size = $System_Drawing_Size
    $labelDesc.text = "ID   | Model"
    $labelDesc.Location = '5,203'

    $labelNot.Name = "Notes"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelNot.size = $System_Drawing_Size
    $labelNot.text = "ID   | Notes"
    $labelNot.Location = '5,268'

    $labelMod.Name = "Model"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMod.size = $System_Drawing_Size
    $labelMod.text = "ID   | Location"
    $labelMod.Location = '5,333'

    $labelModN.Name = "Model #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelModN.size = $System_Drawing_Size
    $labelModN.text = "ID   | Model Number"
    $labelModN.Location = '5,383'

    $labelSerN.Name = "Serial Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelSerN.size = $System_Drawing_Size
    $labelSerN.text = "ID   | Serial Number"
    $labelSerN.Location = '5,433'

    $labelPII.Name = "PII"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 15
    $labelPII.size = $System_Drawing_Size
    $labelPII.text = "ID   |  Lease Number"
    $labelPII.Location = '250,2'

    $labelBrd.Name = "Brand"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelBrd.size = $System_Drawing_Size
    $labelBrd.text = "ID   | Vendor"
    $labelBrd.Location = '250,53'

    $labelPrd.Name = "Product ID"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelPrd.size = $System_Drawing_Size
    $labelPrd.text = "ID   | Purchase Date"
    $labelPrd.Location = '250,103'

    $labelMT.Name = "MT"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMT.size = $System_Drawing_Size
    $labelMT.text = "ID   | Asset Type"
    $labelMT.Location = '250,153'

    #$form1.AcceptButton = $btnQuit2
    $form1.ClientSize = '455, 515' #900,800
    $form1.FormBorderStyle = 'FixedDialog'
    $form1.BackColor = [System.Drawing.Color]::FromArgb(255,185,209,234)
    $form1.MaximizeBox = $False
    $form1.MinimizeBox = $True
    $form1.Name = 'form1'
    $form1.StartPosition = 'CenterScreen'
    $form1.Text = '***IP Search***'
    $form1.KeyPreview = $True
    $form1.add_Load($form1_Load)

    #Device
    $dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridView1.DefaultCellStyle.BackColor = "White"
    $dataGridView1.BackgroundColor = "White"
    $dataGridView1.Name = 'dataGridView1'
    $dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridView1.ReadOnly = $False
    $dataGridView1.AllowUserToDeleteRows = $False
    $dataGridView1.RowHeadersVisible = $false
    $dataGridView1.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridView1.Size = $System_Drawing_Size
    $dataGridView1.TabIndex = 8
    $dataGridView1.Anchor = 15
    $dataGridView1.AutoSizeColumnsMode = 'AllCells'
    $dataGridView1.AllowUserToAddRows = $False
    $dataGridView1.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 15
    $dataGridView1.Location = $System_Drawing_Point
    $dataGridView1.AllowUserToOrderColumns = $True
    $dataGridView1.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridView1AutoSizeColumnsMode.AllCells
    $datagridview1.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    $datagridview1.add_CurrentCellDirtyStateChanged($datagridview1_CurrentCellDirtyStateChanged)
    
    #****Data MAC
    $dataGridViewTest.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewTest.DefaultCellStyle.BackColor = "White"
    $dataGridViewTest.BackgroundColor = "White"
    $dataGridViewTest.Name = 'dataGridView1'
    $dataGridViewTest.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewTest.ReadOnly = $False
    $dataGridViewTest.AllowUserToDeleteRows = $False
    $dataGridViewTest.RowHeadersVisible = $false
    $dataGridViewTest.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewTest.Size = $System_Drawing_Size
    $dataGridViewTest.TabIndex = 8
    $dataGridViewTest.Anchor = 15
    $dataGridViewTest.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewTest.AllowUserToAddRows = $false
    $dataGridViewTest.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 65
    $dataGridViewTest.Location = $System_Drawing_Point
    $dataGridViewTest.AllowUserToOrderColumns = $True
    $dataGridViewTest.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewTestAutoSizeColumnsMode.AllCells
    $datagridviewTest.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Data IP
    $dataGridViewIP.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewIP.DefaultCellStyle.BackColor = "White"
    $dataGridViewIP.BackgroundColor = "White"
    $dataGridViewIP.Name = 'dataGridViewIP'
    $dataGridViewIP.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewIP.ReadOnly = $False
    $dataGridViewIP.AllowUserToDeleteRows = $False
    $dataGridViewIP.ColumnHeadersVisible = $false
    $dataGridViewIP.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewIP.Size = $System_Drawing_Size
    $dataGridViewIP.TabIndex = 8
    $dataGridViewIP.Anchor = 15
    $dataGridViewIP.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewIP.AllowUserToAddRows = $false
    $dataGridViewIP.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 115
    $dataGridViewIP.Location = $System_Drawing_Point
    $dataGridViewIP.AllowUserToOrderColumns = $True
    $dataGridViewIP.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewIPAutoSizeColumnsMode.AllCells
    $datagridviewIP.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })
    
    #**Version
    $dataGridViewVers.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVers.DefaultCellStyle.BackColor = "White"
    $dataGridViewVers.BackgroundColor = "White"
    $dataGridViewVers.Name = 'dataGridViewVers'
    $dataGridViewVers.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVers.ReadOnly = $False
    $dataGridViewVers.AllowUserToDeleteRows = $False
    $dataGridViewVers.ColumnHeadersVisible = $false
    $dataGridViewVers.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVers.Size = $System_Drawing_Size
    $dataGridViewVers.TabIndex = 8
    $dataGridViewVers.Anchor = 15
    $dataGridViewVers.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVers.AllowUserToAddRows = $false
    $dataGridViewVers.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 165
    $dataGridViewVers.Location = $System_Drawing_Point
    $dataGridViewVers.AllowUserToOrderColumns = $True
    $dataGridViewVers.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVersAutoSizeColumnsMode.AllCells
    $datagridviewVers.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Description
    $dataGridViewDesc.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDesc.DefaultCellStyle.BackColor = "White"
    $dataGridViewDesc.BackgroundColor = "White"
    $dataGridViewDesc.Name = 'dataGridViewDesc'
    $dataGridViewDesc.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDesc.ReadOnly = $False
    $dataGridViewDesc.AllowUserToDeleteRows = $False
    $dataGridViewDesc.ColumnHeadersVisible = $false
    $dataGridViewDesc.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 445
    $System_Drawing_Size.Height = 40
    $dataGridViewDesc.Size = $System_Drawing_Size
    $dataGridViewDesc.TabIndex = 8
    $dataGridViewDesc.Anchor = 15
    $dataGridViewDesc.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDesc.AllowUserToAddRows = $false
    $dataGridViewDesc.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 215
    $dataGridViewDesc.Location = $System_Drawing_Point
    $dataGridViewDesc.AllowUserToOrderColumns = $True
    $dataGridViewDesc.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDescAutoSizeColumnsMode.AllCells
    $datagridviewDesc.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Notes
    $dataGridViewNot.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewNot.DefaultCellStyle.BackColor = "White"
    $dataGridViewNot.BackgroundColor = "White"
    $dataGridViewNot.Name = 'dataGridViewNot'
    $dataGridViewNot.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewNot.ReadOnly = $False
    $dataGridViewNot.AllowUserToDeleteRows = $False
    $dataGridViewNot.ColumnHeadersVisible = $false
    $dataGridViewNot.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 445
    $System_Drawing_Size.Height = 40
    $dataGridViewNot.Size = $System_Drawing_Size
    $dataGridViewNot.TabIndex = 8
    $dataGridViewNot.Anchor = 15
    $dataGridViewNot.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewNot.AllowUserToAddRows = $false
    $dataGridViewNot.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 280
    $dataGridViewNot.Location = $System_Drawing_Point
    $dataGridViewNot.AllowUserToOrderColumns = $True
    $dataGridViewNot.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewNotAutoSizeColumnsMode.AllCells
    $datagridviewNot.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Model
    $dataGridViewMod.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewMod.DefaultCellStyle.BackColor = "White"
    $dataGridViewMod.BackgroundColor = "White"
    $dataGridViewMod.Name = 'dataGridViewMod'
    $dataGridViewMod.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewMod.ReadOnly = $False
    $dataGridViewMod.AllowUserToDeleteRows = $False
    $dataGridViewMod.ColumnHeadersVisible = $false
    $dataGridViewMod.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewMod.Size = $System_Drawing_Size
    $dataGridViewMod.TabIndex = 8
    $dataGridViewMod.Anchor = 15
    $dataGridViewMod.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewMod.AllowUserToAddRows = $false
    $dataGridViewMod.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 345
    $dataGridViewMod.Location = $System_Drawing_Point
    $dataGridViewMod.AllowUserToOrderColumns = $True
    $dataGridViewMod.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewModAutoSizeColumnsMode.AllCells
    $datagridviewMod.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**ModelNumber
    $dataGridViewModN.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewModN.DefaultCellStyle.BackColor = "White"
    $dataGridViewModN.BackgroundColor = "White"
    $dataGridViewModN.Name = 'dataGridViewModN'
    $dataGridViewModN.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewModN.ReadOnly = $False
    $dataGridViewModN.AllowUserToDeleteRows = $False
    $dataGridViewModN.ColumnHeadersVisible = $false
    $dataGridViewModN.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewModN.Size = $System_Drawing_Size
    $dataGridViewModN.TabIndex = 8
    $dataGridViewModN.Anchor = 15
    $dataGridViewModN.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewModN.AllowUserToAddRows = $false
    $dataGridViewModN.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 395
    $dataGridViewModN.Location = $System_Drawing_Point
    $dataGridViewModN.AllowUserToOrderColumns = $True
    $dataGridViewModN.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewModNAutoSizeColumnsMode.AllCells
    $datagridviewModN.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })


    #**SerialNumber
    $dataGridViewSerN.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewSerN.DefaultCellStyle.BackColor = "White"
    $dataGridViewSerN.BackgroundColor = "White"
    $dataGridViewSerN.Name = 'dataGridViewModN'
    $dataGridViewSerN.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewSerN.ReadOnly = $False
    $dataGridViewSerN.AllowUserToDeleteRows = $False
    $dataGridViewSerN.ColumnHeadersVisible = $false
    $dataGridViewSerN.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewSerN.Size = $System_Drawing_Size
    $dataGridViewSerN.TabIndex = 8
    $dataGridViewSerN.Anchor = 15
    $dataGridViewSerN.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewSerN.AllowUserToAddRows = $false
    $dataGridViewSerN.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 445
    $dataGridViewSerN.Location = $System_Drawing_Point
    $dataGridViewSerN.AllowUserToOrderColumns = $True
    $dataGridViewSerN.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewSerNAutoSizeColumnsMode.AllCells
    $datagridviewSerN.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PII
    $dataGridViewPII.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPII.DefaultCellStyle.BackColor = "White"
    $dataGridViewPII.BackgroundColor = "White"
    $dataGridViewPII.Name = 'dataGridViewModN'
    $dataGridViewPII.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPII.ReadOnly = $False
    $dataGridViewPII.AllowUserToDeleteRows = $False
    $dataGridViewPII.ColumnHeadersVisible = $false
    $dataGridViewPII.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPII.Size = $System_Drawing_Size
    $dataGridViewPII.TabIndex = 8
    $dataGridViewPII.Anchor = 15
    $dataGridViewPII.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPII.AllowUserToAddRows = $false
    $dataGridViewPII.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 15
    $dataGridViewPII.Location = $System_Drawing_Point
    $dataGridViewPII.AllowUserToOrderColumns = $True
    $dataGridViewPII.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPIIAutoSizeColumnsMode.AllCells
    $datagridviewPII.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Brand
    $dataGridViewBrd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewBrd.DefaultCellStyle.BackColor = "White"
    $dataGridViewBrd.BackgroundColor = "White"
    $dataGridViewBrd.Name = 'dataGridViewBrd'
    $dataGridViewBrd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewBrd.ReadOnly = $False
    $dataGridViewBrd.AllowUserToDeleteRows = $False
    $dataGridViewBrd.ColumnHeadersVisible = $false
    $dataGridViewBrd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewBrd.Size = $System_Drawing_Size
    $dataGridViewBrd.TabIndex = 8
    $dataGridViewBrd.Anchor = 15
    $dataGridViewBrd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewBrd.AllowUserToAddRows = $false
    $dataGridViewBrd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 65
    $dataGridViewBrd.Location = $System_Drawing_Point
    $dataGridViewBrd.AllowUserToOrderColumns = $True
    $dataGridViewBrd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewBrdAutoSizeColumnsMode.AllCells
    $datagridviewBrd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PrdID
    $dataGridViewPrd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPrd.DefaultCellStyle.BackColor = "White"
    $dataGridViewPrd.BackgroundColor = "White"
    $dataGridViewPrd.Name = 'dataGridViewBrd'
    $dataGridViewPrd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPrd.ReadOnly = $False
    $dataGridViewPrd.AllowUserToDeleteRows = $False
    $dataGridViewPrd.ColumnHeadersVisible = $false
    $dataGridViewPrd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPrd.Size = $System_Drawing_Size
    $dataGridViewPrd.TabIndex = 8
    $dataGridViewPrd.Anchor = 15
    $dataGridViewPrd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPrd.AllowUserToAddRows = $false
    $dataGridViewPrd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 115
    $dataGridViewPrd.Location = $System_Drawing_Point
    $dataGridViewPrd.AllowUserToOrderColumns = $True
    $dataGridViewPrd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPrdAutoSizeColumnsMode.AllCells
    $datagridviewPrd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**MT
    $dataGridViewMT.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewMT.DefaultCellStyle.BackColor = "White"
    $dataGridViewMT.BackgroundColor = "White"
    $dataGridViewMT.Name = 'dataGridViewMT'
    $dataGridViewMT.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewMT.ReadOnly = $False
    $dataGridViewMT.AllowUserToDeleteRows = $False
    $dataGridViewMT.ColumnHeadersVisible = $false
    $dataGridViewMT.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewMT.Size = $System_Drawing_Size
    $dataGridViewMT.TabIndex = 8
    $dataGridViewMT.Anchor = 15
    $dataGridViewMT.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewMT.AllowUserToAddRows = $false
    $dataGridViewMT.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 165
    $dataGridViewMT.Location = $System_Drawing_Point
    $dataGridViewMT.AllowUserToOrderColumns = $True
    $dataGridViewMT.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewMTAutoSizeColumnsMode.AllCells
    $datagridviewMT.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    $buttonOK.Anchor = 'Bottom, Right'
    $buttonOK.DialogResult = 'Ok'
    $buttonOK.Location = '13,485'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '105,23' #75
    $buttonOK.Text = '&Save and Close'
    $buttonOK.UseVisualStyleBackColor = $True
    $buttonOK.add_Click($buttonOK_Click)

    $btnQuit2.Anchor = 'Bottom, Left'
    $btnQuit2.DialogResult = 'Ok'
    $btnQuit2.Location = '125,485'
    $btnQuit2.Name = 'buttonOK'
    $btnQuit2.Size = '75,23'
    $btnQuit2.Text = '&Close'
    $btnQuit2.UseVisualStyleBackColor = $True
    $btnQuit2.add_Click($Quit2)

    $form1.ResumeLayout()

    $InitialFormWindowState = $form1.WindowState
    $form1.add_Load($Form_StateCorrection_Load)
    $form1.ShowDialog()

}

#*********************************************************************************************
#**************GenerateForm - PARENT SHEET - Read Only****************************************
#**************STRIKING REFRESH BUTTON WILL REFRESH CHANGES MADE ON OTHER FORMS***************
#*********************************************************************************************
function GenerateForm{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form1 = New-Object 'System.Windows.Forms.Form'
    $datagridview1 = New-Object 'System.Windows.Forms.DataGridView'
    $buttonOK = New-Object 'System.Windows.Forms.Button'
    $ip2TxtSQLQuery = New-Object 'System.Windows.Forms.TextBox'
    $ip2TxtSQLQuery353 = New-Object 'System.Windows.Forms.TextBox'
    $btn353 = New-Object 'System.Windows.Forms.Button'
    $btnQuit2 = New-Object 'System.Windows.Forms.Button'
    $btn35 = New-Object 'System.Windows.Forms.Button'
    $btnRefi = New-Object 'System.Windows.Forms.Button'
    $btnEdit = New-Object 'System.Windows.Forms.Button'
    $btnScanners = New-Object 'System.Windows.Forms.Button'
    $btnPrinters = New-Object 'System.Windows.Forms.Button'

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=True" #SSPI

    $form1_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)
    }

    $buttonOK_Click = {
        if ($script:DVGhasChanged -and [System.Windows.Forms.MessageBox]::Show('BOOM', 'Data Changed', 'YesNo')){
            $script:adapter.Update($datagridview1.DataSource)
        }
    }

    $datagridview1_CurrentCellDirtyStateChanged = {
        $script:DVGhasChanged = $true
    }
    $Form_StateCorrection_Load = {
        $form1.WindowState = "Maximized" #$InitialFormWindowState 
    }
    
    $form1.SuspendLayout()

    #form1
    $form1.Controls.Add($datagridview1)
    #$form1.Controls.Add($buttonOK)
    #$form1.AcceptButton = $buttonOK
    $form1.Controls.Add($btnQuit2)
    $form1.Controls.Add($btnRefi)
    $form1.Controls.Add($btnEdit)
    $form1.Controls.Add($btnScanners)
    $form1.Controls.Add($btnPrinters)
    $form1.ClientSize = '890, 374'
    $form1.FormBorderStyle = 'Sizable'
    $form1.BackColor = "Gray" #[System.Drawing.Color]::FromArgb(255,185,209,234)
    $form1.MaximizeBox = $True
    $form1.MinimizeBox = $True
    $form1.Name = 'form1'
    $form1.StartPosition = 'CenterScreen'
    $form1.Text = 'Query All'
    $form1.KeyPreview = $True
    $form1.add_Load($form1_Load)

    #***************************************************************************************
    $ip2TxtSQLQuery.Text = "10.10.60.10"
    $ip2TxtSQLQuery.Name = 'txtSQLQuery'
    $ip2TxtSQLQuery.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 20
    $ip2TxtSQLQuery.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 46
    $ip2TxtSQLQuery.Location = $System_Drawing_Point
    $ip2TxtSQLQuery.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn35.PerformClick()
        }
    })
    
    $form1.Controls.Add($ip2TxtSQLQuery)

    $btn35.Anchor = 'Top, Left'
    $btn35.UseVisualStyleBackColor = $True
    $btn35.Text = 'Select IP'
    $btn35.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn35.TabIndex = 3
    $btn35.Name = 'btn3'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 90
    $System_Drawing_Size.Height = 23
    $btn35.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 45
    $btn35.Location = $System_Drawing_Point
    $btn35.add_Click($SelectIP_Click2)

    $form1.Controls.Add($btn35)
    #******************************************************************************************
    $ip2TxtSQLQuery353.Text = "Enter  Asset Number"
    $ip2TxtSQLQuery353.Name = 'txtSQLQuery353'
    $ip2TxtSQLQuery353.TabIndex = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 23
    $ip2TxtSQLQuery353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 205
    $System_Drawing_Point.Y = 17
    $ip2TxtSQLQuery353.Location = $System_Drawing_Point
    $ip2TxtSQLQuery353.DataBindings.DefaultDataSourceUpdateMode = 0
    $ip2TxtSQLQuery353.Anchor = 'top, Left' #was Bottom Left
    $ip2TxtSQLQuery353.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $btn353.PerformClick()
        }
    })
    
    $form1.Controls.Add($ip2TxtSQLQuery353)

    $btn353.Anchor = 'Top, Left'
    $btn353.UseVisualStyleBackColor = $True
    $btn353.Text = 'Select Device'
    $btn353.DataBindings.DefaultDataSourceUpdateMode = 0
    $btn353.TabIndex = 4
    $btn353.Name = 'btn353'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 90
    $System_Drawing_Size.Height = 23
    $btn353.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 15
    $btn353.Location = $System_Drawing_Point
    $btn353.add_Click($SelectIP_Click23)

    $form1.Controls.Add($btn353)

    #******************************************************************************************

    $dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridView1.DefaultCellStyle.BackColor = "White"
    $dataGridView1.BackgroundColor = "White"
    $dataGridView1.Name = 'dataGridView'
    $dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridView1.ReadOnly = $True
    $dataGridView1.AllowUserToDeleteRows = $False
    $dataGridView1.RowHeadersVisible = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 870
    $System_Drawing_Size.Height = 260
    $dataGridView1.Size = $System_Drawing_Size
    $dataGridView1.TabIndex = 8
    $dataGridView1.Anchor = 15
    $dataGridView1.AutoSizeColumnsMode = 'AllCells'
    $dataGridView1.AllowUserToAddRows = $True
    $dataGridView1.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 70
    $dataGridView1.Location = $System_Drawing_Point
    $dataGridView1.AllowUserToOrderColumns = $True
    $dataGridView1.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridView1AutoSizeColumnsMode.AllCells
    $datagridview1.add_DoubleClick($datagridview1_Click)
    
    $datagridview1.add_CurrentCellDirtyStateChanged($datagridview1_CurrentCellDirtyStateChanged)

    $buttonOK.Anchor = 'Bottom, Left'
    $buttonOK.DialogResult = 'Ok'
    $buttonOK.Location = '13,339'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '105,23' #75
    $buttonOK.Text = '&Save and Close'
    $buttonOK.UseVisualStyleBackColor = $True
    $buttonOK.add_Click($buttonOK_Click)

    $btnQuit2.Anchor = 'Bottom, Left'
    $btnQuit2.DialogResult = 'Ok'
    $btnQuit2.Location = '125,339'
    $btnQuit2.Name = 'buttonOK'
    $btnQuit2.Size = '75,23'
    $btnQuit2.Text = '&Close'
    $btnQuit2.UseVisualStyleBackColor = $True
    $btnQuit2.add_Click($Quit2)

    $btnRefi.UseVisualStyleBackColor = $True
    $btnRefi.Text = 'Refresh'
    $btnRefi.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnRefi.TabIndex = 1
    $btnRefi.Name = 'btnAll'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 90
    $System_Drawing_Size.Height = 23
    $btnRefi.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 15
    $btnRefi.Location = $System_Drawing_Point
    $btnRefi.add_Click({Refresh})

    $btnEdit.UseVisualStyleBackColor = $True
    $btnEdit.Text = 'Edit All'
    $btnEdit.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnEdit.TabIndex = 1
    $btnEdit.Name = 'btnEdit'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 90
    $System_Drawing_Size.Height = 23
    $btnEdit.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 45
    $btnEdit.Location = $System_Drawing_Point
    $btnEdit.add_Click($Refi_Click)

    $btnScanners.UseVisualStyleBackColor = $True
    $btnScanners.Text = 'Edit Scanners'
    $btnScanners.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnScanners.TabIndex = 8
    $btnScanners.Name = 'btnScanners'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 85
    $System_Drawing_Size.Height = 23
    $btnScanners.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 338
    $System_Drawing_Point.Y = 15
    $btnScanners.Location = $System_Drawing_Point
    $btnScanners.add_Click($Scanners_Click)

    $btnPrinters.UseVisualStyleBackColor = $True
    $btnPrinters.Text = 'Edit Printers'
    $btnPrinters.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnPrinters.TabIndex = 8
    $btnPrinters.Name = 'btnPrinters'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 85
    $System_Drawing_Size.Height = 23
    $btnPrinters.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 338
    $System_Drawing_Point.Y = 45
    $btnPrinters.Location = $System_Drawing_Point
    $btnPrinters.add_Click($Printers_Click)

    $form1.ResumeLayout()

    $InitialFormWindowState = $form1.WindowState
    $form1.add_Load($Form_StateCorrection_Load)
    $form1.ShowDialog()

}

#**********************SHOW/HIDE PS CONSOLE WINDOW************************
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()

    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

Hide-Console
#**************************************************************************

GenerateForm




