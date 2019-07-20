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
    $Query = 'select * from IPAddresses order by CAST(PARSENAME([ipaddress], 4) AS INT),
CAST(PARSENAME([ipaddress], 3) AS INT),
CAST(PARSENAME([ipaddress], 2) AS INT),
CAST(PARSENAME([ipaddress], 1) AS INT)'
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
        $TEST = $ip2TxtSQLQuery353.Text
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

    #***********TEXT BOX THAT WILL SEARCH QUERY FOR User Enter IPADDRESS - 192.168.1.98 Is AN EXAMPLE***********************
    $ip2TxtSQLQuery.Text = "192.168.1.1"
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

    $ip2TxtSQLQuery353.Text = "Enter Device Name"
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
    
    $dataGridViewVrtl = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewHtSv = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPrch = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewAst = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewVnd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewInv = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewDeco = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewDeDt = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewWarV = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewWarE = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewWarS = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewLInv = New-Object 'System.Windows.Forms.DataGridView'
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
    
    $labelVrtl = New-Object 'System.Windows.Forms.Label'
    $labelHtSv = New-Object 'System.Windows.Forms.Label'
    $labelPrch = New-Object 'System.Windows.Forms.Label'
    
    $labelAst = New-Object 'System.Windows.Forms.Label'
    $labelVnd = New-Object 'System.Windows.Forms.Label'
    $labelInv = New-Object 'System.Windows.Forms.Label'
    
    $labelDeco = New-Object 'System.Windows.Forms.Label'
    $labelDeDt = New-Object 'System.Windows.Forms.Label'
    
    $labelWarV = New-Object 'System.Windows.Forms.Label'
    $labelWarE = New-Object 'System.Windows.Forms.Label'
    $labelWarS = New-Object 'System.Windows.Forms.Label'
    
    $labelLInv = New-Object 'System.Windows.Forms.Label'

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=SSPI"

    $form1_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select id,device FROM IPAddresses where id= '$selectedCell' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)

        
        $cmd2 = $conn.CreateCommand()
        $cmd2.CommandText = "Select id,MAC FROM IPAddresses where id= '$selectedCell' "
        $script:adapter2 = New-Object System.Data.SqlClient.SqlDataAdapter($cmd2)
        $dt2 = New-Object System.Data.DataTable
        $script:adapter2.Fill($dt2)
        $datagridviewTest.DataSource = $dt2
        $cmdBldr2 = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter2)

        $cmdIP = $conn.CreateCommand()
        $cmdIP.CommandText = "Select id,ipaddress FROM IPAddresses where id= '$selectedCell' "
        $script:adapterIP = New-Object System.Data.SqlClient.SqlDataAdapter($cmdIP)
        $dtIP = New-Object System.Data.DataTable
        $script:adapterIP.Fill($dtIP)
        $datagridviewIP.DataSource = $dtIP
        $cmdBldrIP = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterIP)

        $cmdVers = $conn.CreateCommand()
        $cmdVers.CommandText = "Select id,version FROM IPAddresses where id= '$selectedCell' "
        $script:adapterVers = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVers)
        $dtVers = New-Object System.Data.DataTable
        $script:adapterVers.Fill($dtVers)
        $datagridviewVers.DataSource = $dtVers
        $cmdBldrVers = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVers)

        $cmdDesc = $conn.CreateCommand()
        $cmdDesc.CommandText = "Select id,description FROM IPAddresses where id= '$selectedCell' "
        $script:adapterDesc = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDesc)
        $dtDesc = New-Object System.Data.DataTable
        $script:adapterDesc.Fill($dtDesc)
        $datagridviewDesc.DataSource = $dtDesc
        $cmdBldrDesc = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDesc)

        $cmdNot = $conn.CreateCommand()
        $cmdNot.CommandText = "Select id,notes FROM IPAddresses where id= '$selectedCell' "
        $script:adapterNot = New-Object System.Data.SqlClient.SqlDataAdapter($cmdNot)
        $dtNot = New-Object System.Data.DataTable
        $script:adapterNot.Fill($dtNot)
        $datagridviewNot.DataSource = $dtNot
        $cmdBldrNot = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterNot)

        $cmdMod = $conn.CreateCommand()
        $cmdMod.CommandText = "Select id,model FROM IPAddresses where id= '$selectedCell' "
        $script:adapterMod = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMod)
        $dtMod = New-Object System.Data.DataTable
        $script:adapterMod.Fill($dtMod)
        $datagridviewMod.DataSource = $dtMod
        $cmdBldrMod = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMod)

        $cmdModN = $conn.CreateCommand()
        $cmdModN.CommandText = "Select id,modelNumber FROM IPAddresses where id= '$selectedCell' "
        $script:adapterModN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdModN)
        $dtModN = New-Object System.Data.DataTable
        $script:adapterModN.Fill($dtModN)
        $datagridviewModN.DataSource = $dtModN
        $cmdBldrModN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterModN)

        $cmdSerN = $conn.CreateCommand()
        $cmdSerN.CommandText = "Select id,serialNumber FROM IPAddresses where id= '$selectedCell' "
        $script:adapterSerN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdSerN)
        $dtSerN = New-Object System.Data.DataTable
        $script:adapterSerN.Fill($dtSerN)
        $datagridviewSerN.DataSource = $dtSerN
        $cmdBldrSerN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterSerN)

        $cmdPII = $conn.CreateCommand()
        $cmdPII.CommandText = "Select id,PII FROM IPAddresses where id= '$selectedCell' "
        $script:adapterPII = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPII)
        $dtPII = New-Object System.Data.DataTable
        $script:adapterPII.Fill($dtPII)
        $datagridviewPII.DataSource = $dtPII
        $cmdBldrPII = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPII)

        $cmdBrd = $conn.CreateCommand()
        $cmdBrd.CommandText = "Select id,brand FROM IPAddresses where id= '$selectedCell' "
        $script:adapterBrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdBrd)
        $dtBrd = New-Object System.Data.DataTable
        $script:adapterBrd.Fill($dtBrd)
        $datagridviewBrd.DataSource = $dtBrd
        $cmdBldrBrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterBrd)

        $cmdPrd = $conn.CreateCommand()
        $cmdPrd.CommandText = "Select id,PrdID FROM IPAddresses where id= '$selectedCell' "
        $script:adapterPrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrd)
        $dtPrd = New-Object System.Data.DataTable
        $script:adapterPrd.Fill($dtPrd)
        $datagridviewPrd.DataSource = $dtPrd
        $cmdBldrPrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrd)

        $cmdMT = $conn.CreateCommand()
        $cmdMT.CommandText = "Select id,MT FROM IPAddresses where id= '$selectedCell' "
        $script:adapterMT = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMT)
        $dtMT = New-Object System.Data.DataTable
        $script:adapterMT.Fill($dtMT)
        $datagridviewMT.DataSource = $dtMT
        $cmdBldrMT = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMT)

        

        $cmdVrtl = $conn.CreateCommand()
        $cmdVrtl.CommandText = "Select id,Virtual FROM IPAddresses where id= '$selectedCell' "
        $script:adapterVrtl = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVrtl)
        $dtVrtl = New-Object System.Data.DataTable
        $script:adapterVrtl.Fill($dtVrtl)
        $datagridviewVrtl.DataSource = $dtVrtl
        $cmdBldrVrtl = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVrtl)

        $cmdHtSv = $conn.CreateCommand()
        $cmdHtSv.CommandText = "Select id,HostServer FROM IPAddresses where id= '$selectedCell' "
        $script:adapterHtSv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdHtSv)
        $dtHtSv = New-Object System.Data.DataTable
        $script:adapterHtSv.Fill($dtHtSv)
        $datagridviewHtSv.DataSource = $dtHtSv
        $cmdBldrHtSv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterHtSv)

        $cmdPrch = $conn.CreateCommand()
        $cmdPrch.CommandText = "Select id,PurchDate FROM IPAddresses where id= '$selectedCell' "
        $script:adapterPrch = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrch)
        $dtPrch = New-Object System.Data.DataTable
        $script:adapterPrch.Fill($dtPrch)
        $datagridviewPrch.DataSource = $dtPrch
        $cmdBldrPrch = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrch)

        

        $cmdAst = $conn.CreateCommand()
        $cmdAst.CommandText = "Select id,AssetNum FROM IPAddresses where id= '$selectedCell' "
        $script:adapterAst = New-Object System.Data.SqlClient.SqlDataAdapter($cmdAst)
        $dtAst = New-Object System.Data.DataTable
        $script:adapterAst.Fill($dtAst)
        $datagridviewAst.DataSource = $dtAst
        $cmdBldrAst = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterAst)

        $cmdVnd = $conn.CreateCommand()
        $cmdVnd.CommandText = "Select id,vendor FROM IPAddresses where id= '$selectedCell' "
        $script:adapterVnd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVnd)
        $dtVnd = New-Object System.Data.DataTable
        $script:adapterVnd.Fill($dtVnd)
        $datagridviewVnd.DataSource = $dtVnd
        $cmdBldrVnd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVnd)

        $cmdInv = $conn.CreateCommand()
        $cmdInv.CommandText = "Select id,invnum FROM IPAddresses where id= '$selectedCell' "
        $script:adapterInv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdInv)
        $dtInv = New-Object System.Data.DataTable
        $script:adapterInv.Fill($dtInv)
        $datagridviewInv.DataSource = $dtInv
        $cmdBldrInv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterInv)

        

        $cmdDeco = $conn.CreateCommand()
        $cmdDeco.CommandText = "Select id,decommissioned FROM IPAddresses where id= '$selectedCell' "
        $script:adapterDeco = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDeco)
        $dtDeco = New-Object System.Data.DataTable
        $script:adapterDeco.Fill($dtDeco)
        $datagridviewDeco.DataSource = $dtDeco
        $cmdBldrDeco = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDeco)

        $cmdDeDt = $conn.CreateCommand()
        $cmdDeDt.CommandText = "Select id,datedecommissioned FROM IPAddresses where id= '$selectedCell' "
        $script:adapterDeDt = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDeDt)
        $dtDeDt = New-Object System.Data.DataTable
        $script:adapterDeDt.Fill($dtDeDt)
        $datagridviewDedt.DataSource = $dtDeDt
        $cmdBldrDeDt = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDedt)

        

        $cmdWarV = $conn.CreateCommand()
        $cmdWarV.CommandText = "Select id,warrantyvendor FROM IPAddresses where id= '$selectedCell' "
        $script:adapterWarV = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarV)
        $dtWarV = New-Object System.Data.DataTable
        $script:adapterWarV.Fill($dtWarV)
        $datagridviewWarV.DataSource = $dtWarV
        $cmdBldrWarV = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarV)

        $cmdWarE = $conn.CreateCommand()
        $cmdWarE.CommandText = "Select id,warrantyenddate FROM IPAddresses where id= '$selectedCell' "
        $script:adapterWarE = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarE)
        $dtWarE = New-Object System.Data.DataTable
        $script:adapterWarE.Fill($dtWarE)
        $datagridviewWarE.DataSource = $dtWarE
        $cmdBldrWarE = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarE)

        $cmdWarS = $conn.CreateCommand()
        $cmdWarS.CommandText = "Select id,warrantySLA FROM IPAddresses where id= '$selectedCell' "
        $script:adapterWarS = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarS)
        $dtWarS = New-Object System.Data.DataTable
        $script:adapterWarS.Fill($dtWarS)
        $datagridviewWarS.DataSource = $dtWarS
        $cmdBldrWarS = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarS)

        

        $cmdLInv = $conn.CreateCommand()
        $cmdLInv.CommandText = "Select id,lastinvdate FROM IPAddresses where id= '$selectedCell' "
        $script:adapterLInv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdLInv)
        $dtLInv = New-Object System.Data.DataTable
        $script:adapterLInv.Fill($dtLInv)
        $datagridviewLInv.DataSource = $dtLInv
        $cmdBldrLInv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterLInv)
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
            
            $script:adapterVrtl.Update($dataGridViewVrtl.DataSource)
            $script:adapterHtSv.Update($dataGridViewHtSv.DataSource)
            $script:adapterPrch.Update($dataGridViewPrch.DataSource)
            
            $script:adapterAst.Update($dataGridViewAst.DataSource)
            $script:adapterVnd.Update($dataGridViewVnd.DataSource)
            $script:adapterInv.Update($dataGridViewInv.DataSource)
            
            $script:adapterDeco.Update($dataGridViewDeco.DataSource)
            $script:adapterDeDt.Update($dataGridViewDeDt.DataSource)
            
            $script:adapterWarV.Update($dataGridViewWarV.DataSource)
            $script:adapterWarE.Update($dataGridViewWarE.DataSource)
            $script:adapterWarS.Update($dataGridViewWarS.DataSource)
            
            $script:adapterLInv.Update($dataGridViewLInv.DataSource)
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
    
    $form1.Controls.Add($dataGridViewVrtl)
    $form1.Controls.Add($dataGridViewHtSv)
    $form1.Controls.Add($dataGridViewPrch)
    
    $form1.Controls.Add($dataGridViewAst)
    $form1.Controls.Add($dataGridViewVnd)
    $form1.Controls.Add($dataGridViewInv)
    
    $form1.Controls.Add($dataGridViewDeco)
    $form1.Controls.Add($dataGridViewDeDt)
    
    $form1.Controls.Add($dataGridViewWarV)
    $form1.Controls.Add($dataGridViewWarE)
    $form1.Controls.Add($dataGridViewWarS)
    
    $form1.Controls.Add($dataGridViewLInv)
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
    
    $form1.Controls.Add($labelVrtl)
    $form1.Controls.Add($labelHtSv)
    $form1.Controls.Add($labelPrch)
    
    $form1.Controls.Add($labelAst)
    $form1.Controls.Add($labelVnd)
    $form1.Controls.Add($labelInv)
    
    $form1.Controls.Add($labelDeco)
    $form1.Controls.Add($labelDeDt)
    
    $form1.Controls.Add($labelWarV)
    $form1.Controls.Add($labelWarE)
    $form1.Controls.Add($labelWarS)
    
    $form1.Controls.Add($labelLInv)

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
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMAC.size = $System_Drawing_Size
    $labelMAC.text = "ID   | MAC"
    $labelMAC.Location = '5,53'
    
    $labelVers.Name = "Version"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVers.size = $System_Drawing_Size
    $labelVers.text = "ID   | Version"
    $labelVers.Location = '5,153'

    $labelDesc.Name = "Description"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDesc.size = $System_Drawing_Size
    $labelDesc.text = "ID   | Description"
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
    $labelMod.text = "ID   | Model"
    $labelMod.Location = '5,333'

    $labelModN.Name = "Model #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelModN.size = $System_Drawing_Size
    $labelModN.text = "ID   | Model Number"
    $labelModN.Location = '5,383'

    $labelSerN.Name = "Serial Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelSerN.size = $System_Drawing_Size
    $labelSerN.text = "ID   | Serial #"
    $labelSerN.Location = '5,433'

    $labelPII.Name = "PII"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelPII.size = $System_Drawing_Size
    $labelPII.text = "ID   | PII"
    $labelPII.Location = '250,2'

    $labelBrd.Name = "Brand"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelBrd.size = $System_Drawing_Size
    $labelBrd.text = "ID   | Brand"
    $labelBrd.Location = '250,53'

    $labelPrd.Name = "Product ID"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelPrd.size = $System_Drawing_Size
    $labelPrd.text = "ID   | Product ID"
    $labelPrd.Location = '250,103'

    $labelMT.Name = "MT"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMT.size = $System_Drawing_Size
    $labelMT.text = "ID   | MT"
    $labelMT.Location = '250,153'

    

    $labelVrtl.Name = "Virtual"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVrtl.size = $System_Drawing_Size
    $labelVrtl.text = "ID   | Virtual"
    $labelVrtl.Location = '250,333'

    $labelHtSv.Name = "HostServer"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelHtSv.size = $System_Drawing_Size
    $labelHtSv.text = "ID   | Host Server"
    $labelHtSv.Location = '250,383'

    $labelPrch.Name = "Purchase Date"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelPrch.size = $System_Drawing_Size
    $labelPrch.text = "ID   | Purchase Date"
    $labelPrch.Location = '250,433'

    

    $labelAst.Name = "Asset #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelAst.size = $System_Drawing_Size
    $labelAst.text = "ID   | Asset Number"
    $labelAst.Location = '500,2'

    $labelVnd.Name = "Vendor"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelVnd.size = $System_Drawing_Size
    $labelVnd.text = "ID   | Vendor"
    $labelVnd.Location = '500,53'

    $labelInv.Name = "Inv Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelInv.size = $System_Drawing_Size
    $labelInv.text = "ID   | Invoice Number"
    $labelInv.Location = '500,103'

    

    $labelDeco.Name = "Decommissioned"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelDeco.size = $System_Drawing_Size
    $labelDeco.text = "ID   | Decommissioned"
    $labelDeco.Location = '500,153'

    $labelDeDt.Name = "Date Decommissioned"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelDeDt.size = $System_Drawing_Size
    $labelDeDt.text = "ID   | Date Decommissioned"
    $labelDeDt.Location = '500,203'

    

    $labelWarV.Name = "Warranty Vendor"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarV.size = $System_Drawing_Size
    $labelWarV.text = "ID   | Warranty Vendor"
    $labelWarV.Location = '500,268'

    $labelWarE.Name = "Warranty End Date"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarE.size = $System_Drawing_Size
    $labelWarE.text = "ID   | Warranty End Date"
    $labelWarE.Location = '500,333'

    $labelWarS.Name = "WarSLA"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarS.size = $System_Drawing_Size
    $labelWarS.text = "ID   | Warranty SLA"
    $labelWarS.Location = '500,383'

    

    $labelLInv.Name = "Last Invoice"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelLInv.size = $System_Drawing_Size
    $labelLInv.text = "ID   | Last Invoice Date"
    $labelLInv.Location = '500,433'

    #$form1.AcceptButton = $btnQuit2
    $form1.ClientSize = '725, 515' #900,800
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

    #**Virtual
    $dataGridViewVrtl.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVrtl.DefaultCellStyle.BackColor = "White"
    $dataGridViewVrtl.BackgroundColor = "White"
    $dataGridViewVrtl.Name = 'dataGridViewRkPs'
    $dataGridViewVrtl.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVrtl.ReadOnly = $False
    $dataGridViewVrtl.AllowUserToDeleteRows = $False
    $dataGridViewVrtl.ColumnHeadersVisible = $false
    $dataGridViewVrtl.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVrtl.Size = $System_Drawing_Size
    $dataGridViewVrtl.TabIndex = 8
    $dataGridViewVrtl.Anchor = 15
    $dataGridViewVrtl.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVrtl.AllowUserToAddRows = $false
    $dataGridViewVrtl.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 345
    $dataGridViewVrtl.Location = $System_Drawing_Point
    $dataGridViewVrtl.AllowUserToOrderColumns = $True
    $dataGridViewVrtl.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVrtlAutoSizeColumnsMode.AllCells
    $datagridviewVrtl.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**HostServer
    $dataGridViewHtSv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewHtSv.DefaultCellStyle.BackColor = "White"
    $dataGridViewHtSv.BackgroundColor = "White"
    $dataGridViewHtSv.Name = 'dataGridViewHtSv'
    $dataGridViewHtSv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewHtSv.ReadOnly = $False
    $dataGridViewHtSv.AllowUserToDeleteRows = $False
    $dataGridViewHtSv.ColumnHeadersVisible = $false
    $dataGridViewHtSv.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewHtSv.Size = $System_Drawing_Size
    $dataGridViewHtSv.TabIndex = 8
    $dataGridViewHtSv.Anchor = 15
    $dataGridViewHtSv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewHtSv.AllowUserToAddRows = $false
    $dataGridViewHtSv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 395
    $dataGridViewHtSv.Location = $System_Drawing_Point
    $dataGridViewHtSv.AllowUserToOrderColumns = $True
    $dataGridViewHtSv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewHtSvAutoSizeColumnsMode.AllCells
    $datagridviewHtSv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PurchDate
    $dataGridViewPrch.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPrch.DefaultCellStyle.BackColor = "White"
    $dataGridViewPrch.BackgroundColor = "White"
    $dataGridViewPrch.Name = 'dataGridViewHtSv'
    $dataGridViewPrch.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPrch.ReadOnly = $False
    $dataGridViewPrch.AllowUserToDeleteRows = $False
    $dataGridViewPrch.ColumnHeadersVisible = $false
    $dataGridViewPrch.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPrch.Size = $System_Drawing_Size
    $dataGridViewPrch.TabIndex = 8
    $dataGridViewPrch.Anchor = 15
    $dataGridViewPrch.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPrch.AllowUserToAddRows = $false
    $dataGridViewPrch.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 445
    $dataGridViewPrch.Location = $System_Drawing_Point
    $dataGridViewPrch.AllowUserToOrderColumns = $True
    $dataGridViewPrch.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPrchAutoSizeColumnsMode.AllCells
    $datagridviewPrch.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**AssetNum
    $dataGridViewAst.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewAst.DefaultCellStyle.BackColor = "White"
    $dataGridViewAst.BackgroundColor = "White"
    $dataGridViewAst.Name = 'dataGridViewAst'
    $dataGridViewAst.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewAst.ReadOnly = $False
    $dataGridViewAst.AllowUserToDeleteRows = $False
    $dataGridViewAst.ColumnHeadersVisible = $false
    $dataGridViewAst.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewAst.Size = $System_Drawing_Size
    $dataGridViewAst.TabIndex = 8
    $dataGridViewAst.Anchor = 15
    $dataGridViewAst.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewAst.AllowUserToAddRows = $false
    $dataGridViewAst.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 15
    $dataGridViewAst.Location = $System_Drawing_Point
    $dataGridViewAst.AllowUserToOrderColumns = $True
    $dataGridViewAst.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewAstAutoSizeColumnsMode.AllCells
    $datagridviewAst.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Vendor
    $dataGridViewVnd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVnd.DefaultCellStyle.BackColor = "White"
    $dataGridViewVnd.BackgroundColor = "White"
    $dataGridViewVnd.Name = 'dataGridViewVnd'
    $dataGridViewVnd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVnd.ReadOnly = $False
    $dataGridViewVnd.AllowUserToDeleteRows = $False
    $dataGridViewVnd.ColumnHeadersVisible = $false
    $dataGridViewVnd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVnd.Size = $System_Drawing_Size
    $dataGridViewVnd.TabIndex = 8
    $dataGridViewVnd.Anchor = 15
    $dataGridViewVnd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVnd.AllowUserToAddRows = $false
    $dataGridViewVnd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 65
    $dataGridViewVnd.Location = $System_Drawing_Point
    $dataGridViewVnd.AllowUserToOrderColumns = $True
    $dataGridViewVnd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVndAutoSizeColumnsMode.AllCells
    $datagridviewVnd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**InvNum
    $dataGridViewInv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewInv.DefaultCellStyle.BackColor = "White"
    $dataGridViewInv.BackgroundColor = "White"
    $dataGridViewInv.Name = 'dataGridViewInv'
    $dataGridViewInv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewInv.ReadOnly = $False
    $dataGridViewInv.AllowUserToDeleteRows = $False
    $dataGridViewInv.ColumnHeadersVisible = $false
    $dataGridViewInv.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewInv.Size = $System_Drawing_Size
    $dataGridViewInv.TabIndex = 8
    $dataGridViewInv.Anchor = 15
    $dataGridViewInv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewInv.AllowUserToAddRows = $false
    $dataGridViewInv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 115
    $dataGridViewInv.Location = $System_Drawing_Point
    $dataGridViewInv.AllowUserToOrderColumns = $True
    $dataGridViewInv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewInvAutoSizeColumnsMode.AllCells
    $datagridviewInv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Decommissioned
    $dataGridViewDeco.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDeco.DefaultCellStyle.BackColor = "White"
    $dataGridViewDeco.BackgroundColor = "White"
    $dataGridViewDeco.Name = 'dataGridViewDeco'
    $dataGridViewDeco.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDeco.ReadOnly = $False
    $dataGridViewDeco.AllowUserToDeleteRows = $False
    $dataGridViewDeco.ColumnHeadersVisible = $false
    $dataGridViewDeco.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewDeco.Size = $System_Drawing_Size
    $dataGridViewDeco.TabIndex = 8
    $dataGridViewDeco.Anchor = 15
    $dataGridViewDeco.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDeco.AllowUserToAddRows = $false
    $dataGridViewDeco.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 165
    $dataGridViewDeco.Location = $System_Drawing_Point
    $dataGridViewDeco.AllowUserToOrderColumns = $True
    $dataGridViewDeco.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDecoAutoSizeColumnsMode.AllCells
    $datagridviewDeco.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**DateDecommissioned
    $dataGridViewDeDt.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDeDt.DefaultCellStyle.BackColor = "White"
    $dataGridViewDeDt.BackgroundColor = "White"
    $dataGridViewDeDt.Name = 'dataGridViewDeDt'
    $dataGridViewDeDt.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDeDt.ReadOnly = $False
    $dataGridViewDeDt.AllowUserToDeleteRows = $False
    $dataGridViewDeDt.ColumnHeadersVisible = $false
    $dataGridViewDeDt.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewDeDt.Size = $System_Drawing_Size
    $dataGridViewDeDt.TabIndex = 8
    $dataGridViewDeDt.Anchor = 15
    $dataGridViewDeDt.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDeDt.AllowUserToAddRows = $false
    $dataGridViewDeDt.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 215
    $dataGridViewDeDt.Location = $System_Drawing_Point
    $dataGridViewDeDt.AllowUserToOrderColumns = $True
    $dataGridViewDeDt.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDeDtAutoSizeColumnsMode.AllCells
    $datagridviewDeDt.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantyVendor
    $dataGridViewWarV.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarV.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarV.BackgroundColor = "White"
    $dataGridViewWarV.Name = 'dataGridViewWarV'
    $dataGridViewWarV.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarV.ReadOnly = $False
    $dataGridViewWarV.AllowUserToDeleteRows = $False
    $dataGridViewWarV.ColumnHeadersVisible = $false
    $dataGridViewWarV.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarV.Size = $System_Drawing_Size
    $dataGridViewWarV.TabIndex = 8
    $dataGridViewWarV.Anchor = 15
    $dataGridViewWarV.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarV.AllowUserToAddRows = $false
    $dataGridViewWarV.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 280
    $dataGridViewWarV.Location = $System_Drawing_Point
    $dataGridViewWarV.AllowUserToOrderColumns = $True
    $dataGridViewWarV.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarVAutoSizeColumnsMode.AllCells
    $datagridviewWarV.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantyEndDate
    $dataGridViewWarE.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarE.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarE.BackgroundColor = "White"
    $dataGridViewWarE.Name = 'dataGridViewWarE'
    $dataGridViewWarE.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarE.ReadOnly = $False
    $dataGridViewWarE.AllowUserToDeleteRows = $False
    $dataGridViewWarE.ColumnHeadersVisible = $false
    $dataGridViewWarE.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarE.Size = $System_Drawing_Size
    $dataGridViewWarE.TabIndex = 8
    $dataGridViewWarE.Anchor = 15
    $dataGridViewWarE.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarE.AllowUserToAddRows = $false
    $dataGridViewWarE.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 345
    $dataGridViewWarE.Location = $System_Drawing_Point
    $dataGridViewWarE.AllowUserToOrderColumns = $True
    $dataGridViewWarE.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarEAutoSizeColumnsMode.AllCells
    $datagridviewWarE.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantySLA
    $dataGridViewWarS.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarS.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarS.BackgroundColor = "White"
    $dataGridViewWarS.Name = 'dataGridViewWarS'
    $dataGridViewWarS.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarS.ReadOnly = $False
    $dataGridViewWarS.AllowUserToDeleteRows = $False
    $dataGridViewWarS.RowHeadersVisible = $false
    $dataGridViewWarS.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarS.Size = $System_Drawing_Size
    $dataGridViewWarS.TabIndex = 8
    $dataGridViewWarS.Anchor = 15
    $dataGridViewWarS.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarS.AllowUserToAddRows = $False
    $dataGridViewWarS.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 395
    $dataGridViewWarS.Location = $System_Drawing_Point
    $dataGridViewWarS.AllowUserToOrderColumns = $True
    $dataGridViewWarS.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarSAutoSizeColumnsMode.AllCells
    $datagridviewWarS.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**LastInvDate
    $dataGridViewLInv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewLInv.DefaultCellStyle.BackColor = "White"
    $dataGridViewLInv.BackgroundColor = "White"
    $dataGridViewLInv.Name = 'dataGridViewLInv'
    $dataGridViewLInv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewLInv.ReadOnly = $False
    $dataGridViewLInv.AllowUserToDeleteRows = $False
    $dataGridViewLInv.RowHeadersVisible = $false
    $dataGridViewLInv.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewLInv.Size = $System_Drawing_Size
    $dataGridViewLInv.TabIndex = 8
    $dataGridViewLInv.Anchor = 15
    $dataGridViewLInv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewLInv.AllowUserToAddRows = $False
    $dataGridViewLInv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 445
    $dataGridViewLInv.Location = $System_Drawing_Point
    $dataGridViewLInv.AllowUserToOrderColumns = $True
    $dataGridViewLInv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewLInvAutoSizeColumnsMode.AllCells
    $datagridviewLInv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
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
    
    $dataGridViewVrtl = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewHtSv = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPrch = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewAst = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewVnd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewInv = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewDeco = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewDeDt = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewWarV = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewWarE = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewWarS = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewLInv = New-Object 'System.Windows.Forms.DataGridView'
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
    
    $labelVrtl = New-Object 'System.Windows.Forms.Label'
    $labelHtSv = New-Object 'System.Windows.Forms.Label'
    $labelPrch = New-Object 'System.Windows.Forms.Label'
    
    $labelAst = New-Object 'System.Windows.Forms.Label'
    $labelVnd = New-Object 'System.Windows.Forms.Label'
    $labelInv = New-Object 'System.Windows.Forms.Label'
    
    $labelDeco = New-Object 'System.Windows.Forms.Label'
    $labelDeDt = New-Object 'System.Windows.Forms.Label'
    
    $labelWarV = New-Object 'System.Windows.Forms.Label'
    $labelWarE = New-Object 'System.Windows.Forms.Label'
    $labelWarS = New-Object 'System.Windows.Forms.Label'
   
    $labelLInv = New-Object 'System.Windows.Forms.Label'

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=SSPI"

    $form1_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select id,device FROM IPAddresses where device= '$TEST' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)

        
        $cmd2 = $conn.CreateCommand()
        $cmd2.CommandText = "Select id,MAC FROM IPAddresses where device= '$TEST' "
        $script:adapter2 = New-Object System.Data.SqlClient.SqlDataAdapter($cmd2)
        $dt2 = New-Object System.Data.DataTable
        $script:adapter2.Fill($dt2)
        $datagridviewTest.DataSource = $dt2
        $cmdBldr2 = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter2)

        $cmdIP = $conn.CreateCommand()
        $cmdIP.CommandText = "Select id,ipaddress FROM IPAddresses where device= '$TEST' "
        $script:adapterIP = New-Object System.Data.SqlClient.SqlDataAdapter($cmdIP)
        $dtIP = New-Object System.Data.DataTable
        $script:adapterIP.Fill($dtIP)
        $datagridviewIP.DataSource = $dtIP
        $cmdBldrIP = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterIP)

        $cmdVers = $conn.CreateCommand()
        $cmdVers.CommandText = "Select id,version FROM IPAddresses where device= '$TEST' "
        $script:adapterVers = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVers)
        $dtVers = New-Object System.Data.DataTable
        $script:adapterVers.Fill($dtVers)
        $datagridviewVers.DataSource = $dtVers
        $cmdBldrVers = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVers)

        $cmdDesc = $conn.CreateCommand()
        $cmdDesc.CommandText = "Select id,description FROM IPAddresses where device= '$TEST' "
        $script:adapterDesc = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDesc)
        $dtDesc = New-Object System.Data.DataTable
        $script:adapterDesc.Fill($dtDesc)
        $datagridviewDesc.DataSource = $dtDesc
        $cmdBldrDesc = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDesc)

        $cmdNot = $conn.CreateCommand()
        $cmdNot.CommandText = "Select id,notes FROM IPAddresses where device= '$TEST' "
        $script:adapterNot = New-Object System.Data.SqlClient.SqlDataAdapter($cmdNot)
        $dtNot = New-Object System.Data.DataTable
        $script:adapterNot.Fill($dtNot)
        $datagridviewNot.DataSource = $dtNot
        $cmdBldrNot = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterNot)

        $cmdMod = $conn.CreateCommand()
        $cmdMod.CommandText = "Select id,model FROM IPAddresses where device= '$TEST' "
        $script:adapterMod = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMod)
        $dtMod = New-Object System.Data.DataTable
        $script:adapterMod.Fill($dtMod)
        $datagridviewMod.DataSource = $dtMod
        $cmdBldrMod = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMod)

        $cmdModN = $conn.CreateCommand()
        $cmdModN.CommandText = "Select id,modelNumber FROM IPAddresses where device= '$TEST' "
        $script:adapterModN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdModN)
        $dtModN = New-Object System.Data.DataTable
        $script:adapterModN.Fill($dtModN)
        $datagridviewModN.DataSource = $dtModN
        $cmdBldrModN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterModN)

        $cmdSerN = $conn.CreateCommand()
        $cmdSerN.CommandText = "Select id,serialNumber FROM IPAddresses where device= '$TEST' "
        $script:adapterSerN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdSerN)
        $dtSerN = New-Object System.Data.DataTable
        $script:adapterSerN.Fill($dtSerN)
        $datagridviewSerN.DataSource = $dtSerN
        $cmdBldrSerN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterSerN)

        $cmdPII = $conn.CreateCommand()
        $cmdPII.CommandText = "Select id,PII FROM IPAddresses where device= '$TEST' "
        $script:adapterPII = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPII)
        $dtPII = New-Object System.Data.DataTable
        $script:adapterPII.Fill($dtPII)
        $datagridviewPII.DataSource = $dtPII
        $cmdBldrPII = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPII)

        $cmdBrd = $conn.CreateCommand()
        $cmdBrd.CommandText = "Select id,brand FROM IPAddresses where device= '$TEST' "
        $script:adapterBrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdBrd)
        $dtBrd = New-Object System.Data.DataTable
        $script:adapterBrd.Fill($dtBrd)
        $datagridviewBrd.DataSource = $dtBrd
        $cmdBldrBrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterBrd)

        $cmdPrd = $conn.CreateCommand()
        $cmdPrd.CommandText = "Select id,PrdID FROM IPAddresses where device= '$TEST' "
        $script:adapterPrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrd)
        $dtPrd = New-Object System.Data.DataTable
        $script:adapterPrd.Fill($dtPrd)
        $datagridviewPrd.DataSource = $dtPrd
        $cmdBldrPrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrd)

        $cmdMT = $conn.CreateCommand()
        $cmdMT.CommandText = "Select id,MT FROM IPAddresses where device= '$TEST' "
        $script:adapterMT = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMT)
        $dtMT = New-Object System.Data.DataTable
        $script:adapterMT.Fill($dtMT)
        $datagridviewMT.DataSource = $dtMT
        $cmdBldrMT = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMT)

        

        $cmdVrtl = $conn.CreateCommand()
        $cmdVrtl.CommandText = "Select id,Virtual FROM IPAddresses where device= '$TEST' "
        $script:adapterVrtl = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVrtl)
        $dtVrtl = New-Object System.Data.DataTable
        $script:adapterVrtl.Fill($dtVrtl)
        $datagridviewVrtl.DataSource = $dtVrtl
        $cmdBldrVrtl = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVrtl)

        $cmdHtSv = $conn.CreateCommand()
        $cmdHtSv.CommandText = "Select id,HostServer FROM IPAddresses where device= '$TEST' "
        $script:adapterHtSv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdHtSv)
        $dtHtSv = New-Object System.Data.DataTable
        $script:adapterHtSv.Fill($dtHtSv)
        $datagridviewHtSv.DataSource = $dtHtSv
        $cmdBldrHtSv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterHtSv)

        $cmdPrch = $conn.CreateCommand()
        $cmdPrch.CommandText = "Select id,PurchDate FROM IPAddresses where device= '$TEST' "
        $script:adapterPrch = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrch)
        $dtPrch = New-Object System.Data.DataTable
        $script:adapterPrch.Fill($dtPrch)
        $datagridviewPrch.DataSource = $dtPrch
        $cmdBldrPrch = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrch)

        

        $cmdAst = $conn.CreateCommand()
        $cmdAst.CommandText = "Select id,AssetNum FROM IPAddresses where device= '$TEST' "
        $script:adapterAst = New-Object System.Data.SqlClient.SqlDataAdapter($cmdAst)
        $dtAst = New-Object System.Data.DataTable
        $script:adapterAst.Fill($dtAst)
        $datagridviewAst.DataSource = $dtAst
        $cmdBldrAst = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterAst)

        $cmdVnd = $conn.CreateCommand()
        $cmdVnd.CommandText = "Select id,vendor FROM IPAddresses where device= '$TEST' "
        $script:adapterVnd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVnd)
        $dtVnd = New-Object System.Data.DataTable
        $script:adapterVnd.Fill($dtVnd)
        $datagridviewVnd.DataSource = $dtVnd
        $cmdBldrVnd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVnd)

        $cmdInv = $conn.CreateCommand()
        $cmdInv.CommandText = "Select id,invnum FROM IPAddresses where device= '$TEST' "
        $script:adapterInv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdInv)
        $dtInv = New-Object System.Data.DataTable
        $script:adapterInv.Fill($dtInv)
        $datagridviewInv.DataSource = $dtInv
        $cmdBldrInv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterInv)

        
        

        $cmdDeco = $conn.CreateCommand()
        $cmdDeco.CommandText = "Select id,decommissioned FROM IPAddresses where device= '$TEST' "
        $script:adapterDeco = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDeco)
        $dtDeco = New-Object System.Data.DataTable
        $script:adapterDeco.Fill($dtDeco)
        $datagridviewDeco.DataSource = $dtDeco
        $cmdBldrDeco = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDeco)

        $cmdDeDt = $conn.CreateCommand()
        $cmdDeDt.CommandText = "Select id,datedecommissioned FROM IPAddresses where device= '$TEST' "
        $script:adapterDeDt = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDeDt)
        $dtDeDt = New-Object System.Data.DataTable
        $script:adapterDeDt.Fill($dtDeDt)
        $datagridviewDedt.DataSource = $dtDeDt
        $cmdBldrDeDt = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDedt)

        

        $cmdWarV = $conn.CreateCommand()
        $cmdWarV.CommandText = "Select id,warrantyvendor FROM IPAddresses where device= '$TEST' "
        $script:adapterWarV = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarV)
        $dtWarV = New-Object System.Data.DataTable
        $script:adapterWarV.Fill($dtWarV)
        $datagridviewWarV.DataSource = $dtWarV
        $cmdBldrWarV = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarV)

        $cmdWarE = $conn.CreateCommand()
        $cmdWarE.CommandText = "Select id,warrantyenddate FROM IPAddresses where device= '$TEST' "
        $script:adapterWarE = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarE)
        $dtWarE = New-Object System.Data.DataTable
        $script:adapterWarE.Fill($dtWarE)
        $datagridviewWarE.DataSource = $dtWarE
        $cmdBldrWarE = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarE)

        $cmdWarS = $conn.CreateCommand()
        $cmdWarS.CommandText = "Select id,warrantySLA FROM IPAddresses where device= '$TEST' "
        $script:adapterWarS = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarS)
        $dtWarS = New-Object System.Data.DataTable
        $script:adapterWarS.Fill($dtWarS)
        $datagridviewWarS.DataSource = $dtWarS
        $cmdBldrWarS = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarS)

        

        $cmdLInv = $conn.CreateCommand()
        $cmdLInv.CommandText = "Select id,lastinvdate FROM IPAddresses where device= '$TEST' "
        $script:adapterLInv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdLInv)
        $dtLInv = New-Object System.Data.DataTable
        $script:adapterLInv.Fill($dtLInv)
        $datagridviewLInv.DataSource = $dtLInv
        $cmdBldrLInv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterLInv)
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
            
            $script:adapterVrtl.Update($dataGridViewVrtl.DataSource)
            $script:adapterHtSv.Update($dataGridViewHtSv.DataSource)
            $script:adapterPrch.Update($dataGridViewPrch.DataSource)
            
            $script:adapterAst.Update($dataGridViewAst.DataSource)
            $script:adapterVnd.Update($dataGridViewVnd.DataSource)
            $script:adapterInv.Update($dataGridViewInv.DataSource)
            
            $script:adapterDeco.Update($dataGridViewDeco.DataSource)
            $script:adapterDeDt.Update($dataGridViewDeDt.DataSource)
            
            $script:adapterWarV.Update($dataGridViewWarV.DataSource)
            $script:adapterWarE.Update($dataGridViewWarE.DataSource)
            $script:adapterWarS.Update($dataGridViewWarS.DataSource)
            
            $script:adapterLInv.Update($dataGridViewLInv.DataSource)
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
    
    $form1.Controls.Add($dataGridViewVrtl)
    $form1.Controls.Add($dataGridViewHtSv)
    $form1.Controls.Add($dataGridViewPrch)
    
    $form1.Controls.Add($dataGridViewAst)
    $form1.Controls.Add($dataGridViewVnd)
    $form1.Controls.Add($dataGridViewInv)
    
    $form1.Controls.Add($dataGridViewDeco)
    $form1.Controls.Add($dataGridViewDeDt)
    
    $form1.Controls.Add($dataGridViewWarV)
    $form1.Controls.Add($dataGridViewWarE)
    $form1.Controls.Add($dataGridViewWarS)
    
    $form1.Controls.Add($dataGridViewLInv)
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
    
    $form1.Controls.Add($labelVrtl)
    $form1.Controls.Add($labelHtSv)
    $form1.Controls.Add($labelPrch)
    
    $form1.Controls.Add($labelAst)
    $form1.Controls.Add($labelVnd)
    $form1.Controls.Add($labelInv)
    
    $form1.Controls.Add($labelDeco)
    $form1.Controls.Add($labelDeDt)
    
    $form1.Controls.Add($labelWarV)
    $form1.Controls.Add($labelWarE)
    $form1.Controls.Add($labelWarS)
    
    $form1.Controls.Add($labelLInv)

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
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMAC.size = $System_Drawing_Size
    $labelMAC.text = "ID   | MAC"
    $labelMAC.Location = '5,53'
    
    $labelVers.Name = "Version"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVers.size = $System_Drawing_Size
    $labelVers.text = "ID   | Version"
    $labelVers.Location = '5,153'

    $labelDesc.Name = "Description"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDesc.size = $System_Drawing_Size
    $labelDesc.text = "ID   | Description"
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
    $labelMod.text = "ID   | Model"
    $labelMod.Location = '5,333'

    $labelModN.Name = "Model #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelModN.size = $System_Drawing_Size
    $labelModN.text = "ID   | Model Number"
    $labelModN.Location = '5,383'

    $labelSerN.Name = "Serial Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelSerN.size = $System_Drawing_Size
    $labelSerN.text = "ID   | Serial #"
    $labelSerN.Location = '5,433'

    $labelPII.Name = "PII"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelPII.size = $System_Drawing_Size
    $labelPII.text = "ID   | PII"
    $labelPII.Location = '250,2'

    $labelBrd.Name = "Brand"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelBrd.size = $System_Drawing_Size
    $labelBrd.text = "ID   | Brand"
    $labelBrd.Location = '250,53'

    $labelPrd.Name = "Product ID"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelPrd.size = $System_Drawing_Size
    $labelPrd.text = "ID   | Product ID"
    $labelPrd.Location = '250,103'

    $labelMT.Name = "MT"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMT.size = $System_Drawing_Size
    $labelMT.text = "ID   | MT"
    $labelMT.Location = '250,153'

    

    $labelVrtl.Name = "Virtual"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVrtl.size = $System_Drawing_Size
    $labelVrtl.text = "ID   | Virtual"
    $labelVrtl.Location = '250,333'

    $labelHtSv.Name = "HostServer"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelHtSv.size = $System_Drawing_Size
    $labelHtSv.text = "ID   | Host Server"
    $labelHtSv.Location = '250,383'

    $labelPrch.Name = "Purchase Date"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelPrch.size = $System_Drawing_Size
    $labelPrch.text = "ID   | Purchase Date"
    $labelPrch.Location = '250,433'

    

    $labelAst.Name = "Asset #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelAst.size = $System_Drawing_Size
    $labelAst.text = "ID   | Asset Number"
    $labelAst.Location = '500,2'

    $labelVnd.Name = "Vendor"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelVnd.size = $System_Drawing_Size
    $labelVnd.text = "ID   | Vendor"
    $labelVnd.Location = '500,53'

    $labelInv.Name = "Inv Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelInv.size = $System_Drawing_Size
    $labelInv.text = "ID   | Invoice Number"
    $labelInv.Location = '500,103'

    

    $labelDeco.Name = "Decommissioned"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelDeco.size = $System_Drawing_Size
    $labelDeco.text = "ID   | Decommissioned"
    $labelDeco.Location = '500,153'

    $labelDeDt.Name = "Date Decommissioned"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelDeDt.size = $System_Drawing_Size
    $labelDeDt.text = "ID   | Date Decommissioned"
    $labelDeDt.Location = '500,203'

    

    $labelWarV.Name = "Warranty Vendor"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarV.size = $System_Drawing_Size
    $labelWarV.text = "ID   | Warranty Vendor"
    $labelWarV.Location = '500,268'

    $labelWarE.Name = "Warranty End Date"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarE.size = $System_Drawing_Size
    $labelWarE.text = "ID   | Warranty End Date"
    $labelWarE.Location = '500,333'

    $labelWarS.Name = "WarSLA"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarS.size = $System_Drawing_Size
    $labelWarS.text = "ID   | Warranty SLA"
    $labelWarS.Location = '500,383'

    

    $labelLInv.Name = "Last Invoice"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelLInv.size = $System_Drawing_Size
    $labelLInv.text = "ID   | Last Invoice Date"
    $labelLInv.Location = '500,433'

    #$form1.AcceptButton = $btnQuit2
    $form1.ClientSize = '725, 515' #900,800
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

    #**Virtual
    $dataGridViewVrtl.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVrtl.DefaultCellStyle.BackColor = "White"
    $dataGridViewVrtl.BackgroundColor = "White"
    $dataGridViewVrtl.Name = 'dataGridViewRkPs'
    $dataGridViewVrtl.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVrtl.ReadOnly = $False
    $dataGridViewVrtl.AllowUserToDeleteRows = $False
    $dataGridViewVrtl.ColumnHeadersVisible = $false
    $dataGridViewVrtl.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVrtl.Size = $System_Drawing_Size
    $dataGridViewVrtl.TabIndex = 8
    $dataGridViewVrtl.Anchor = 15
    $dataGridViewVrtl.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVrtl.AllowUserToAddRows = $false
    $dataGridViewVrtl.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 345
    $dataGridViewVrtl.Location = $System_Drawing_Point
    $dataGridViewVrtl.AllowUserToOrderColumns = $True
    $dataGridViewVrtl.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVrtlAutoSizeColumnsMode.AllCells
    $datagridviewVrtl.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**HostServer
    $dataGridViewHtSv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewHtSv.DefaultCellStyle.BackColor = "White"
    $dataGridViewHtSv.BackgroundColor = "White"
    $dataGridViewHtSv.Name = 'dataGridViewHtSv'
    $dataGridViewHtSv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewHtSv.ReadOnly = $False
    $dataGridViewHtSv.AllowUserToDeleteRows = $False
    $dataGridViewHtSv.ColumnHeadersVisible = $false
    $dataGridViewHtSv.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewHtSv.Size = $System_Drawing_Size
    $dataGridViewHtSv.TabIndex = 8
    $dataGridViewHtSv.Anchor = 15
    $dataGridViewHtSv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewHtSv.AllowUserToAddRows = $false
    $dataGridViewHtSv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 395
    $dataGridViewHtSv.Location = $System_Drawing_Point
    $dataGridViewHtSv.AllowUserToOrderColumns = $True
    $dataGridViewHtSv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewHtSvAutoSizeColumnsMode.AllCells
    $datagridviewHtSv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PurchDate
    $dataGridViewPrch.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPrch.DefaultCellStyle.BackColor = "White"
    $dataGridViewPrch.BackgroundColor = "White"
    $dataGridViewPrch.Name = 'dataGridViewHtSv'
    $dataGridViewPrch.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPrch.ReadOnly = $False
    $dataGridViewPrch.AllowUserToDeleteRows = $False
    $dataGridViewPrch.ColumnHeadersVisible = $false
    $dataGridViewPrch.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPrch.Size = $System_Drawing_Size
    $dataGridViewPrch.TabIndex = 8
    $dataGridViewPrch.Anchor = 15
    $dataGridViewPrch.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPrch.AllowUserToAddRows = $false
    $dataGridViewPrch.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 445
    $dataGridViewPrch.Location = $System_Drawing_Point
    $dataGridViewPrch.AllowUserToOrderColumns = $True
    $dataGridViewPrch.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPrchAutoSizeColumnsMode.AllCells
    $datagridviewPrch.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**AssetNum
    $dataGridViewAst.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewAst.DefaultCellStyle.BackColor = "White"
    $dataGridViewAst.BackgroundColor = "White"
    $dataGridViewAst.Name = 'dataGridViewAst'
    $dataGridViewAst.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewAst.ReadOnly = $False
    $dataGridViewAst.AllowUserToDeleteRows = $False
    $dataGridViewAst.ColumnHeadersVisible = $false
    $dataGridViewAst.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewAst.Size = $System_Drawing_Size
    $dataGridViewAst.TabIndex = 8
    $dataGridViewAst.Anchor = 15
    $dataGridViewAst.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewAst.AllowUserToAddRows = $false
    $dataGridViewAst.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 15
    $dataGridViewAst.Location = $System_Drawing_Point
    $dataGridViewAst.AllowUserToOrderColumns = $True
    $dataGridViewAst.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewAstAutoSizeColumnsMode.AllCells
    $datagridviewAst.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Vendor
    $dataGridViewVnd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVnd.DefaultCellStyle.BackColor = "White"
    $dataGridViewVnd.BackgroundColor = "White"
    $dataGridViewVnd.Name = 'dataGridViewVnd'
    $dataGridViewVnd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVnd.ReadOnly = $False
    $dataGridViewVnd.AllowUserToDeleteRows = $False
    $dataGridViewVnd.ColumnHeadersVisible = $false
    $dataGridViewVnd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVnd.Size = $System_Drawing_Size
    $dataGridViewVnd.TabIndex = 8
    $dataGridViewVnd.Anchor = 15
    $dataGridViewVnd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVnd.AllowUserToAddRows = $false
    $dataGridViewVnd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 65
    $dataGridViewVnd.Location = $System_Drawing_Point
    $dataGridViewVnd.AllowUserToOrderColumns = $True
    $dataGridViewVnd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVndAutoSizeColumnsMode.AllCells
    $datagridviewVnd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**InvNum
    $dataGridViewInv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewInv.DefaultCellStyle.BackColor = "White"
    $dataGridViewInv.BackgroundColor = "White"
    $dataGridViewInv.Name = 'dataGridViewInv'
    $dataGridViewInv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewInv.ReadOnly = $False
    $dataGridViewInv.AllowUserToDeleteRows = $False
    $dataGridViewInv.ColumnHeadersVisible = $false
    $dataGridViewInv.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewInv.Size = $System_Drawing_Size
    $dataGridViewInv.TabIndex = 8
    $dataGridViewInv.Anchor = 15
    $dataGridViewInv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewInv.AllowUserToAddRows = $false
    $dataGridViewInv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 115
    $dataGridViewInv.Location = $System_Drawing_Point
    $dataGridViewInv.AllowUserToOrderColumns = $True
    $dataGridViewInv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewInvAutoSizeColumnsMode.AllCells
    $datagridviewInv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Decommissioned
    $dataGridViewDeco.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDeco.DefaultCellStyle.BackColor = "White"
    $dataGridViewDeco.BackgroundColor = "White"
    $dataGridViewDeco.Name = 'dataGridViewDeco'
    $dataGridViewDeco.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDeco.ReadOnly = $False
    $dataGridViewDeco.AllowUserToDeleteRows = $False
    $dataGridViewDeco.ColumnHeadersVisible = $false
    $dataGridViewDeco.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewDeco.Size = $System_Drawing_Size
    $dataGridViewDeco.TabIndex = 8
    $dataGridViewDeco.Anchor = 15
    $dataGridViewDeco.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDeco.AllowUserToAddRows = $false
    $dataGridViewDeco.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 165
    $dataGridViewDeco.Location = $System_Drawing_Point
    $dataGridViewDeco.AllowUserToOrderColumns = $True
    $dataGridViewDeco.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDecoAutoSizeColumnsMode.AllCells
    $datagridviewDeco.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**DateDecommissioned
    $dataGridViewDeDt.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDeDt.DefaultCellStyle.BackColor = "White"
    $dataGridViewDeDt.BackgroundColor = "White"
    $dataGridViewDeDt.Name = 'dataGridViewDeDt'
    $dataGridViewDeDt.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDeDt.ReadOnly = $False
    $dataGridViewDeDt.AllowUserToDeleteRows = $False
    $dataGridViewDeDt.ColumnHeadersVisible = $false
    $dataGridViewDeDt.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewDeDt.Size = $System_Drawing_Size
    $dataGridViewDeDt.TabIndex = 8
    $dataGridViewDeDt.Anchor = 15
    $dataGridViewDeDt.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDeDt.AllowUserToAddRows = $false
    $dataGridViewDeDt.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 215
    $dataGridViewDeDt.Location = $System_Drawing_Point
    $dataGridViewDeDt.AllowUserToOrderColumns = $True
    $dataGridViewDeDt.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDeDtAutoSizeColumnsMode.AllCells
    $datagridviewDeDt.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantyVendor
    $dataGridViewWarV.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarV.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarV.BackgroundColor = "White"
    $dataGridViewWarV.Name = 'dataGridViewWarV'
    $dataGridViewWarV.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarV.ReadOnly = $False
    $dataGridViewWarV.AllowUserToDeleteRows = $False
    $dataGridViewWarV.ColumnHeadersVisible = $false
    $dataGridViewWarV.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarV.Size = $System_Drawing_Size
    $dataGridViewWarV.TabIndex = 8
    $dataGridViewWarV.Anchor = 15
    $dataGridViewWarV.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarV.AllowUserToAddRows = $false
    $dataGridViewWarV.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 280
    $dataGridViewWarV.Location = $System_Drawing_Point
    $dataGridViewWarV.AllowUserToOrderColumns = $True
    $dataGridViewWarV.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarVAutoSizeColumnsMode.AllCells
    $datagridviewWarV.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantyEndDate
    $dataGridViewWarE.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarE.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarE.BackgroundColor = "White"
    $dataGridViewWarE.Name = 'dataGridViewWarE'
    $dataGridViewWarE.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarE.ReadOnly = $False
    $dataGridViewWarE.AllowUserToDeleteRows = $False
    $dataGridViewWarE.ColumnHeadersVisible = $false
    $dataGridViewWarE.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarE.Size = $System_Drawing_Size
    $dataGridViewWarE.TabIndex = 8
    $dataGridViewWarE.Anchor = 15
    $dataGridViewWarE.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarE.AllowUserToAddRows = $false
    $dataGridViewWarE.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 345
    $dataGridViewWarE.Location = $System_Drawing_Point
    $dataGridViewWarE.AllowUserToOrderColumns = $True
    $dataGridViewWarE.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarEAutoSizeColumnsMode.AllCells
    $datagridviewWarE.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantySLA
    $dataGridViewWarS.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarS.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarS.BackgroundColor = "White"
    $dataGridViewWarS.Name = 'dataGridViewWarS'
    $dataGridViewWarS.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarS.ReadOnly = $False
    $dataGridViewWarS.AllowUserToDeleteRows = $False
    $dataGridViewWarS.RowHeadersVisible = $false
    $dataGridViewWarS.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarS.Size = $System_Drawing_Size
    $dataGridViewWarS.TabIndex = 8
    $dataGridViewWarS.Anchor = 15
    $dataGridViewWarS.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarS.AllowUserToAddRows = $False
    $dataGridViewWarS.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 395
    $dataGridViewWarS.Location = $System_Drawing_Point
    $dataGridViewWarS.AllowUserToOrderColumns = $True
    $dataGridViewWarS.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarSAutoSizeColumnsMode.AllCells
    $datagridviewWarS.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**LastInvDate
    $dataGridViewLInv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewLInv.DefaultCellStyle.BackColor = "White"
    $dataGridViewLInv.BackgroundColor = "White"
    $dataGridViewLInv.Name = 'dataGridViewLInv'
    $dataGridViewLInv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewLInv.ReadOnly = $False
    $dataGridViewLInv.AllowUserToDeleteRows = $False
    $dataGridViewLInv.RowHeadersVisible = $false
    $dataGridViewLInv.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewLInv.Size = $System_Drawing_Size
    $dataGridViewLInv.TabIndex = 8
    $dataGridViewLInv.Anchor = 15
    $dataGridViewLInv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewLInv.AllowUserToAddRows = $False
    $dataGridViewLInv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 445
    $dataGridViewLInv.Location = $System_Drawing_Point
    $dataGridViewLInv.AllowUserToOrderColumns = $True
    $dataGridViewLInv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewLInvAutoSizeColumnsMode.AllCells
    $datagridviewLInv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
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
    
    $dataGridViewVrtl = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewHtSv = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewPrch = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewAst = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewVnd = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewInv = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewDeco = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewDeDt = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewWarV = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewWarE = New-Object 'System.Windows.Forms.DataGridView'
    $dataGridViewWarS = New-Object 'System.Windows.Forms.DataGridView'
    
    $dataGridViewLInv = New-Object 'System.Windows.Forms.DataGridView'
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
    
    $labelVrtl = New-Object 'System.Windows.Forms.Label'
    $labelHtSv = New-Object 'System.Windows.Forms.Label'
    $labelPrch = New-Object 'System.Windows.Forms.Label'
    
    $labelAst = New-Object 'System.Windows.Forms.Label'
    $labelVnd = New-Object 'System.Windows.Forms.Label'
    $labelInv = New-Object 'System.Windows.Forms.Label'
    
    $labelDeco = New-Object 'System.Windows.Forms.Label'
    $labelDeDt = New-Object 'System.Windows.Forms.Label'
    
    $labelWarV = New-Object 'System.Windows.Forms.Label'
    $labelWarE = New-Object 'System.Windows.Forms.Label'
    $labelWarS = New-Object 'System.Windows.Forms.Label'
    
    $labelLInv = New-Object 'System.Windows.Forms.Label'

    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    $DVGhasChanged = $false
    $connStr = "Server=$Instance;Database=$Database;Integrated Security=SSPI"

    $form1_Load = {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "Select id,device FROM IPAddresses where ipaddress= '$userText' "
        $script:adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        $script:adapter.Fill($dt)
        $datagridview1.DataSource = $dt
        $cmdBldr = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter)

        
        $cmd2 = $conn.CreateCommand()
        $cmd2.CommandText = "Select id,MAC FROM IPAddresses where ipaddress= '$userText' "
        $script:adapter2 = New-Object System.Data.SqlClient.SqlDataAdapter($cmd2)
        $dt2 = New-Object System.Data.DataTable
        $script:adapter2.Fill($dt2)
        $datagridviewTest.DataSource = $dt2
        $cmdBldr2 = New-Object System.Data.SqlClient.SqlCommandBuilder($adapter2)

        $cmdIP = $conn.CreateCommand()
        $cmdIP.CommandText = "Select id,ipaddress FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterIP = New-Object System.Data.SqlClient.SqlDataAdapter($cmdIP)
        $dtIP = New-Object System.Data.DataTable
        $script:adapterIP.Fill($dtIP)
        $datagridviewIP.DataSource = $dtIP
        $cmdBldrIP = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterIP)

        $cmdVers = $conn.CreateCommand()
        $cmdVers.CommandText = "Select id,version FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterVers = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVers)
        $dtVers = New-Object System.Data.DataTable
        $script:adapterVers.Fill($dtVers)
        $datagridviewVers.DataSource = $dtVers
        $cmdBldrVers = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVers)

        $cmdDesc = $conn.CreateCommand()
        $cmdDesc.CommandText = "Select id,description FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterDesc = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDesc)
        $dtDesc = New-Object System.Data.DataTable
        $script:adapterDesc.Fill($dtDesc)
        $datagridviewDesc.DataSource = $dtDesc
        $cmdBldrDesc = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDesc)

        $cmdNot = $conn.CreateCommand()
        $cmdNot.CommandText = "Select id,notes FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterNot = New-Object System.Data.SqlClient.SqlDataAdapter($cmdNot)
        $dtNot = New-Object System.Data.DataTable
        $script:adapterNot.Fill($dtNot)
        $datagridviewNot.DataSource = $dtNot
        $cmdBldrNot = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterNot)

        $cmdMod = $conn.CreateCommand()
        $cmdMod.CommandText = "Select id,model FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterMod = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMod)
        $dtMod = New-Object System.Data.DataTable
        $script:adapterMod.Fill($dtMod)
        $datagridviewMod.DataSource = $dtMod
        $cmdBldrMod = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMod)

        $cmdModN = $conn.CreateCommand()
        $cmdModN.CommandText = "Select id,modelNumber FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterModN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdModN)
        $dtModN = New-Object System.Data.DataTable
        $script:adapterModN.Fill($dtModN)
        $datagridviewModN.DataSource = $dtModN
        $cmdBldrModN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterModN)

        $cmdSerN = $conn.CreateCommand()
        $cmdSerN.CommandText = "Select id,serialNumber FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterSerN = New-Object System.Data.SqlClient.SqlDataAdapter($cmdSerN)
        $dtSerN = New-Object System.Data.DataTable
        $script:adapterSerN.Fill($dtSerN)
        $datagridviewSerN.DataSource = $dtSerN
        $cmdBldrSerN = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterSerN)

        $cmdPII = $conn.CreateCommand()
        $cmdPII.CommandText = "Select id,PII FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterPII = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPII)
        $dtPII = New-Object System.Data.DataTable
        $script:adapterPII.Fill($dtPII)
        $datagridviewPII.DataSource = $dtPII
        $cmdBldrPII = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPII)

        $cmdBrd = $conn.CreateCommand()
        $cmdBrd.CommandText = "Select id,brand FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterBrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdBrd)
        $dtBrd = New-Object System.Data.DataTable
        $script:adapterBrd.Fill($dtBrd)
        $datagridviewBrd.DataSource = $dtBrd
        $cmdBldrBrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterBrd)

        $cmdPrd = $conn.CreateCommand()
        $cmdPrd.CommandText = "Select id,PrdID FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterPrd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrd)
        $dtPrd = New-Object System.Data.DataTable
        $script:adapterPrd.Fill($dtPrd)
        $datagridviewPrd.DataSource = $dtPrd
        $cmdBldrPrd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrd)

        $cmdMT = $conn.CreateCommand()
        $cmdMT.CommandText = "Select id,MT FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterMT = New-Object System.Data.SqlClient.SqlDataAdapter($cmdMT)
        $dtMT = New-Object System.Data.DataTable
        $script:adapterMT.Fill($dtMT)
        $datagridviewMT.DataSource = $dtMT
        $cmdBldrMT = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterMT)

        

        $cmdVrtl = $conn.CreateCommand()
        $cmdVrtl.CommandText = "Select id,Virtual FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterVrtl = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVrtl)
        $dtVrtl = New-Object System.Data.DataTable
        $script:adapterVrtl.Fill($dtVrtl)
        $datagridviewVrtl.DataSource = $dtVrtl
        $cmdBldrVrtl = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVrtl)

        $cmdHtSv = $conn.CreateCommand()
        $cmdHtSv.CommandText = "Select id,HostServer FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterHtSv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdHtSv)
        $dtHtSv = New-Object System.Data.DataTable
        $script:adapterHtSv.Fill($dtHtSv)
        $datagridviewHtSv.DataSource = $dtHtSv
        $cmdBldrHtSv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterHtSv)

        $cmdPrch = $conn.CreateCommand()
        $cmdPrch.CommandText = "Select id,PurchDate FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterPrch = New-Object System.Data.SqlClient.SqlDataAdapter($cmdPrch)
        $dtPrch = New-Object System.Data.DataTable
        $script:adapterPrch.Fill($dtPrch)
        $datagridviewPrch.DataSource = $dtPrch
        $cmdBldrPrch = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterPrch)

        

        $cmdAst = $conn.CreateCommand()
        $cmdAst.CommandText = "Select id,AssetNum FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterAst = New-Object System.Data.SqlClient.SqlDataAdapter($cmdAst)
        $dtAst = New-Object System.Data.DataTable
        $script:adapterAst.Fill($dtAst)
        $datagridviewAst.DataSource = $dtAst
        $cmdBldrAst = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterAst)

        $cmdVnd = $conn.CreateCommand()
        $cmdVnd.CommandText = "Select id,vendor FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterVnd = New-Object System.Data.SqlClient.SqlDataAdapter($cmdVnd)
        $dtVnd = New-Object System.Data.DataTable
        $script:adapterVnd.Fill($dtVnd)
        $datagridviewVnd.DataSource = $dtVnd
        $cmdBldrVnd = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterVnd)

        $cmdInv = $conn.CreateCommand()
        $cmdInv.CommandText = "Select id,invnum FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterInv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdInv)
        $dtInv = New-Object System.Data.DataTable
        $script:adapterInv.Fill($dtInv)
        $datagridviewInv.DataSource = $dtInv
        $cmdBldrInv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterInv)

        

        $cmdDeco = $conn.CreateCommand()
        $cmdDeco.CommandText = "Select id,decommissioned FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterDeco = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDeco)
        $dtDeco = New-Object System.Data.DataTable
        $script:adapterDeco.Fill($dtDeco)
        $datagridviewDeco.DataSource = $dtDeco
        $cmdBldrDeco = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDeco)

        $cmdDeDt = $conn.CreateCommand()
        $cmdDeDt.CommandText = "Select id,datedecommissioned FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterDeDt = New-Object System.Data.SqlClient.SqlDataAdapter($cmdDeDt)
        $dtDeDt = New-Object System.Data.DataTable
        $script:adapterDeDt.Fill($dtDeDt)
        $datagridviewDedt.DataSource = $dtDeDt
        $cmdBldrDeDt = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterDedt)

        

        

        $cmdWarV = $conn.CreateCommand()
        $cmdWarV.CommandText = "Select id,warrantyvendor FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterWarV = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarV)
        $dtWarV = New-Object System.Data.DataTable
        $script:adapterWarV.Fill($dtWarV)
        $datagridviewWarV.DataSource = $dtWarV
        $cmdBldrWarV = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarV)

        $cmdWarE = $conn.CreateCommand()
        $cmdWarE.CommandText = "Select id,warrantyenddate FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterWarE = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarE)
        $dtWarE = New-Object System.Data.DataTable
        $script:adapterWarE.Fill($dtWarE)
        $datagridviewWarE.DataSource = $dtWarE
        $cmdBldrWarE = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarE)

        $cmdWarS = $conn.CreateCommand()
        $cmdWarS.CommandText = "Select id,warrantySLA FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterWarS = New-Object System.Data.SqlClient.SqlDataAdapter($cmdWarS)
        $dtWarS = New-Object System.Data.DataTable
        $script:adapterWarS.Fill($dtWarS)
        $datagridviewWarS.DataSource = $dtWarS
        $cmdBldrWarS = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterWarS)

        

        $cmdLInv = $conn.CreateCommand()
        $cmdLInv.CommandText = "Select id,lastinvdate FROM IPAddresses where ipaddress= '$userText' "
        $script:adapterLInv = New-Object System.Data.SqlClient.SqlDataAdapter($cmdLInv)
        $dtLInv = New-Object System.Data.DataTable
        $script:adapterLInv.Fill($dtLInv)
        $datagridviewLInv.DataSource = $dtLInv
        $cmdBldrLInv = New-Object System.Data.SqlClient.SqlCommandBuilder($adapterLInv)
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
            
            $script:adapterVrtl.Update($dataGridViewVrtl.DataSource)
            $script:adapterHtSv.Update($dataGridViewHtSv.DataSource)
            $script:adapterPrch.Update($dataGridViewPrch.DataSource)
            
            $script:adapterAst.Update($dataGridViewAst.DataSource)
            $script:adapterVnd.Update($dataGridViewVnd.DataSource)
            $script:adapterInv.Update($dataGridViewInv.DataSource)
            
            $script:adapterDeco.Update($dataGridViewDeco.DataSource)
            $script:adapterDeDt.Update($dataGridViewDeDt.DataSource)
            
            $script:adapterWarV.Update($dataGridViewWarV.DataSource)
            $script:adapterWarE.Update($dataGridViewWarE.DataSource)
            $script:adapterWarS.Update($dataGridViewWarS.DataSource)
            
            $script:adapterLInv.Update($dataGridViewLInv.DataSource)
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
    
    $form1.Controls.Add($dataGridViewVrtl)
    $form1.Controls.Add($dataGridViewHtSv)
    $form1.Controls.Add($dataGridViewPrch)
    
    $form1.Controls.Add($dataGridViewAst)
    $form1.Controls.Add($dataGridViewVnd)
    $form1.Controls.Add($dataGridViewInv)
    
    $form1.Controls.Add($dataGridViewDeco)
    $form1.Controls.Add($dataGridViewDeDt)
    
    $form1.Controls.Add($dataGridViewWarV)
    $form1.Controls.Add($dataGridViewWarE)
    $form1.Controls.Add($dataGridViewWarS)
   
    $form1.Controls.Add($dataGridViewLInv)
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
    
    $form1.Controls.Add($labelVrtl)
    $form1.Controls.Add($labelHtSv)
    $form1.Controls.Add($labelPrch)
    
    $form1.Controls.Add($labelAst)
    $form1.Controls.Add($labelVnd)
    $form1.Controls.Add($labelInv)
    
    $form1.Controls.Add($labelDeco)
    $form1.Controls.Add($labelDeDt)
    
    $form1.Controls.Add($labelWarV)
    $form1.Controls.Add($labelWarE)
    $form1.Controls.Add($labelWarS)
    
    $form1.Controls.Add($labelLInv)

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
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMAC.size = $System_Drawing_Size
    $labelMAC.text = "ID   | MAC"
    $labelMAC.Location = '5,53'
    
    $labelVers.Name = "Version"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVers.size = $System_Drawing_Size
    $labelVers.text = "ID   | Version"
    $labelVers.Location = '5,153'

    $labelDesc.Name = "Description"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelDesc.size = $System_Drawing_Size
    $labelDesc.text = "ID   | Description"
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
    $labelMod.text = "ID   | Model"
    $labelMod.Location = '5,333'

    $labelModN.Name = "Model #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelModN.size = $System_Drawing_Size
    $labelModN.text = "ID   | Model Number"
    $labelModN.Location = '5,383'

    $labelSerN.Name = "Serial Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelSerN.size = $System_Drawing_Size
    $labelSerN.text = "ID   | Serial #"
    $labelSerN.Location = '5,433'

    $labelPII.Name = "PII"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelPII.size = $System_Drawing_Size
    $labelPII.text = "ID   | PII"
    $labelPII.Location = '250,2'

    $labelBrd.Name = "Brand"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelBrd.size = $System_Drawing_Size
    $labelBrd.text = "ID   | Brand"
    $labelBrd.Location = '250,53'

    $labelPrd.Name = "Product ID"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelPrd.size = $System_Drawing_Size
    $labelPrd.text = "ID   | Product ID"
    $labelPrd.Location = '250,103'

    $labelMT.Name = "MT"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelMT.size = $System_Drawing_Size
    $labelMT.text = "ID   | MT"
    $labelMT.Location = '250,153'

    

    $labelVrtl.Name = "Virtual"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelVrtl.size = $System_Drawing_Size
    $labelVrtl.text = "ID   | Virtual"
    $labelVrtl.Location = '250,333'

    $labelHtSv.Name = "HostServer"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 100
    $System_Drawing_Size.Height = 15
    $labelHtSv.size = $System_Drawing_Size
    $labelHtSv.text = "ID   | Host Server"
    $labelHtSv.Location = '250,383'

    $labelPrch.Name = "Purchase Date"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelPrch.size = $System_Drawing_Size
    $labelPrch.text = "ID   | Purchase Date"
    $labelPrch.Location = '250,433'

    

    $labelAst.Name = "Asset #"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelAst.size = $System_Drawing_Size
    $labelAst.text = "ID   | Asset Number"
    $labelAst.Location = '500,2'

    $labelVnd.Name = "Vendor"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelVnd.size = $System_Drawing_Size
    $labelVnd.text = "ID   | Vendor"
    $labelVnd.Location = '500,53'

    $labelInv.Name = "Inv Number"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 115
    $System_Drawing_Size.Height = 15
    $labelInv.size = $System_Drawing_Size
    $labelInv.text = "ID   | Invoice Number"
    $labelInv.Location = '500,103'

    

    $labelDeco.Name = "Decommissioned"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelDeco.size = $System_Drawing_Size
    $labelDeco.text = "ID   | Decommissioned"
    $labelDeco.Location = '500,153'

    $labelDeDt.Name = "Date Decommissioned"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelDeDt.size = $System_Drawing_Size
    $labelDeDt.text = "ID   | Date Decommissioned"
    $labelDeDt.Location = '500,203'

    

    $labelWarV.Name = "Warranty Vendor"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarV.size = $System_Drawing_Size
    $labelWarV.text = "ID   | Warranty Vendor"
    $labelWarV.Location = '500,268'

    $labelWarE.Name = "Warranty End Date"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarE.size = $System_Drawing_Size
    $labelWarE.text = "ID   | Warranty End Date"
    $labelWarE.Location = '500,333'

    $labelWarS.Name = "WarSLA"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelWarS.size = $System_Drawing_Size
    $labelWarS.text = "ID   | Warranty SLA"
    $labelWarS.Location = '500,383'

    

    $labelLInv.Name = "Last Invoice"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 15
    $labelLInv.size = $System_Drawing_Size
    $labelLInv.text = "ID   | Last Invoice Date"
    $labelLInv.Location = '500,433'

    #$form1.AcceptButton = $btnQuit2
    $form1.ClientSize = '725, 515' #900,800
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

    #**Virtual
    $dataGridViewVrtl.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVrtl.DefaultCellStyle.BackColor = "White"
    $dataGridViewVrtl.BackgroundColor = "White"
    $dataGridViewVrtl.Name = 'dataGridViewRkPs'
    $dataGridViewVrtl.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVrtl.ReadOnly = $False
    $dataGridViewVrtl.AllowUserToDeleteRows = $False
    $dataGridViewVrtl.ColumnHeadersVisible = $false
    $dataGridViewVrtl.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVrtl.Size = $System_Drawing_Size
    $dataGridViewVrtl.TabIndex = 8
    $dataGridViewVrtl.Anchor = 15
    $dataGridViewVrtl.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVrtl.AllowUserToAddRows = $false
    $dataGridViewVrtl.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 345
    $dataGridViewVrtl.Location = $System_Drawing_Point
    $dataGridViewVrtl.AllowUserToOrderColumns = $True
    $dataGridViewVrtl.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVrtlAutoSizeColumnsMode.AllCells
    $datagridviewVrtl.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**HostServer
    $dataGridViewHtSv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewHtSv.DefaultCellStyle.BackColor = "White"
    $dataGridViewHtSv.BackgroundColor = "White"
    $dataGridViewHtSv.Name = 'dataGridViewHtSv'
    $dataGridViewHtSv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewHtSv.ReadOnly = $False
    $dataGridViewHtSv.AllowUserToDeleteRows = $False
    $dataGridViewHtSv.ColumnHeadersVisible = $false
    $dataGridViewHtSv.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewHtSv.Size = $System_Drawing_Size
    $dataGridViewHtSv.TabIndex = 8
    $dataGridViewHtSv.Anchor = 15
    $dataGridViewHtSv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewHtSv.AllowUserToAddRows = $false
    $dataGridViewHtSv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 395
    $dataGridViewHtSv.Location = $System_Drawing_Point
    $dataGridViewHtSv.AllowUserToOrderColumns = $True
    $dataGridViewHtSv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewHtSvAutoSizeColumnsMode.AllCells
    $datagridviewHtSv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**PurchDate
    $dataGridViewPrch.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewPrch.DefaultCellStyle.BackColor = "White"
    $dataGridViewPrch.BackgroundColor = "White"
    $dataGridViewPrch.Name = 'dataGridViewHtSv'
    $dataGridViewPrch.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewPrch.ReadOnly = $False
    $dataGridViewPrch.AllowUserToDeleteRows = $False
    $dataGridViewPrch.ColumnHeadersVisible = $false
    $dataGridViewPrch.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewPrch.Size = $System_Drawing_Size
    $dataGridViewPrch.TabIndex = 8
    $dataGridViewPrch.Anchor = 15
    $dataGridViewPrch.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewPrch.AllowUserToAddRows = $false
    $dataGridViewPrch.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 250
    $System_Drawing_Point.Y = 445
    $dataGridViewPrch.Location = $System_Drawing_Point
    $dataGridViewPrch.AllowUserToOrderColumns = $True
    $dataGridViewPrch.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewPrchAutoSizeColumnsMode.AllCells
    $datagridviewPrch.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**AssetNum
    $dataGridViewAst.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewAst.DefaultCellStyle.BackColor = "White"
    $dataGridViewAst.BackgroundColor = "White"
    $dataGridViewAst.Name = 'dataGridViewAst'
    $dataGridViewAst.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewAst.ReadOnly = $False
    $dataGridViewAst.AllowUserToDeleteRows = $False
    $dataGridViewAst.ColumnHeadersVisible = $false
    $dataGridViewAst.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewAst.Size = $System_Drawing_Size
    $dataGridViewAst.TabIndex = 8
    $dataGridViewAst.Anchor = 15
    $dataGridViewAst.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewAst.AllowUserToAddRows = $false
    $dataGridViewAst.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 15
    $dataGridViewAst.Location = $System_Drawing_Point
    $dataGridViewAst.AllowUserToOrderColumns = $True
    $dataGridViewAst.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewAstAutoSizeColumnsMode.AllCells
    $datagridviewAst.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Vendor
    $dataGridViewVnd.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewVnd.DefaultCellStyle.BackColor = "White"
    $dataGridViewVnd.BackgroundColor = "White"
    $dataGridViewVnd.Name = 'dataGridViewVnd'
    $dataGridViewVnd.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewVnd.ReadOnly = $False
    $dataGridViewVnd.AllowUserToDeleteRows = $False
    $dataGridViewVnd.ColumnHeadersVisible = $false
    $dataGridViewVnd.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewVnd.Size = $System_Drawing_Size
    $dataGridViewVnd.TabIndex = 8
    $dataGridViewVnd.Anchor = 15
    $dataGridViewVnd.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewVnd.AllowUserToAddRows = $false
    $dataGridViewVnd.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 65
    $dataGridViewVnd.Location = $System_Drawing_Point
    $dataGridViewVnd.AllowUserToOrderColumns = $True
    $dataGridViewVnd.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewVndAutoSizeColumnsMode.AllCells
    $datagridviewVnd.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**InvNum
    $dataGridViewInv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewInv.DefaultCellStyle.BackColor = "White"
    $dataGridViewInv.BackgroundColor = "White"
    $dataGridViewInv.Name = 'dataGridViewInv'
    $dataGridViewInv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewInv.ReadOnly = $False
    $dataGridViewInv.AllowUserToDeleteRows = $False
    $dataGridViewInv.ColumnHeadersVisible = $false
    $dataGridViewInv.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewInv.Size = $System_Drawing_Size
    $dataGridViewInv.TabIndex = 8
    $dataGridViewInv.Anchor = 15
    $dataGridViewInv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewInv.AllowUserToAddRows = $false
    $dataGridViewInv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 115
    $dataGridViewInv.Location = $System_Drawing_Point
    $dataGridViewInv.AllowUserToOrderColumns = $True
    $dataGridViewInv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewInvAutoSizeColumnsMode.AllCells
    $datagridviewInv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**Decommissioned
    $dataGridViewDeco.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDeco.DefaultCellStyle.BackColor = "White"
    $dataGridViewDeco.BackgroundColor = "White"
    $dataGridViewDeco.Name = 'dataGridViewDeco'
    $dataGridViewDeco.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDeco.ReadOnly = $False
    $dataGridViewDeco.AllowUserToDeleteRows = $False
    $dataGridViewDeco.ColumnHeadersVisible = $false
    $dataGridViewDeco.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewDeco.Size = $System_Drawing_Size
    $dataGridViewDeco.TabIndex = 8
    $dataGridViewDeco.Anchor = 15
    $dataGridViewDeco.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDeco.AllowUserToAddRows = $false
    $dataGridViewDeco.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 165
    $dataGridViewDeco.Location = $System_Drawing_Point
    $dataGridViewDeco.AllowUserToOrderColumns = $True
    $dataGridViewDeco.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDecoAutoSizeColumnsMode.AllCells
    $datagridviewDeco.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**DateDecommissioned
    $dataGridViewDeDt.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewDeDt.DefaultCellStyle.BackColor = "White"
    $dataGridViewDeDt.BackgroundColor = "White"
    $dataGridViewDeDt.Name = 'dataGridViewDeDt'
    $dataGridViewDeDt.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewDeDt.ReadOnly = $False
    $dataGridViewDeDt.AllowUserToDeleteRows = $False
    $dataGridViewDeDt.ColumnHeadersVisible = $false
    $dataGridViewDeDt.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewDeDt.Size = $System_Drawing_Size
    $dataGridViewDeDt.TabIndex = 8
    $dataGridViewDeDt.Anchor = 15
    $dataGridViewDeDt.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewDeDt.AllowUserToAddRows = $false
    $dataGridViewDeDt.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 215
    $dataGridViewDeDt.Location = $System_Drawing_Point
    $dataGridViewDeDt.AllowUserToOrderColumns = $True
    $dataGridViewDeDt.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewDeDtAutoSizeColumnsMode.AllCells
    $datagridviewDeDt.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantyVendor
    $dataGridViewWarV.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarV.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarV.BackgroundColor = "White"
    $dataGridViewWarV.Name = 'dataGridViewWarV'
    $dataGridViewWarV.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarV.ReadOnly = $False
    $dataGridViewWarV.AllowUserToDeleteRows = $False
    $dataGridViewWarV.ColumnHeadersVisible = $false
    $dataGridViewWarV.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarV.Size = $System_Drawing_Size
    $dataGridViewWarV.TabIndex = 8
    $dataGridViewWarV.Anchor = 15
    $dataGridViewWarV.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarV.AllowUserToAddRows = $false
    $dataGridViewWarV.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 280
    $dataGridViewWarV.Location = $System_Drawing_Point
    $dataGridViewWarV.AllowUserToOrderColumns = $True
    $dataGridViewWarV.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarVAutoSizeColumnsMode.AllCells
    $datagridviewWarV.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantyEndDate
    $dataGridViewWarE.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarE.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarE.BackgroundColor = "White"
    $dataGridViewWarE.Name = 'dataGridViewWarE'
    $dataGridViewWarE.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarE.ReadOnly = $False
    $dataGridViewWarE.AllowUserToDeleteRows = $False
    $dataGridViewWarE.ColumnHeadersVisible = $false
    $dataGridViewWarE.RowHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarE.Size = $System_Drawing_Size
    $dataGridViewWarE.TabIndex = 8
    $dataGridViewWarE.Anchor = 15
    $dataGridViewWarE.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarE.AllowUserToAddRows = $false
    $dataGridViewWarE.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 345
    $dataGridViewWarE.Location = $System_Drawing_Point
    $dataGridViewWarE.AllowUserToOrderColumns = $True
    $dataGridViewWarE.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarEAutoSizeColumnsMode.AllCells
    $datagridviewWarE.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**WarrantySLA
    $dataGridViewWarS.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewWarS.DefaultCellStyle.BackColor = "White"
    $dataGridViewWarS.BackgroundColor = "White"
    $dataGridViewWarS.Name = 'dataGridViewWarS'
    $dataGridViewWarS.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewWarS.ReadOnly = $False
    $dataGridViewWarS.AllowUserToDeleteRows = $False
    $dataGridViewWarS.RowHeadersVisible = $false
    $dataGridViewWarS.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewWarS.Size = $System_Drawing_Size
    $dataGridViewWarS.TabIndex = 8
    $dataGridViewWarS.Anchor = 15
    $dataGridViewWarS.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewWarS.AllowUserToAddRows = $False
    $dataGridViewWarS.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 395
    $dataGridViewWarS.Location = $System_Drawing_Point
    $dataGridViewWarS.AllowUserToOrderColumns = $True
    $dataGridViewWarS.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewWarSAutoSizeColumnsMode.AllCells
    $datagridviewWarS.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {
            $buttonOK.PerformClick()
        }
    })

    #**LastInvDate
    $dataGridViewLInv.RowTemplate.DefaultCellStyle.ForeColor = "Black" #[System.Drawing.Color]::FromArgb(255,0,128,0)
    $dataGridViewLInv.DefaultCellStyle.BackColor = "White"
    $dataGridViewLInv.BackgroundColor = "White"
    $dataGridViewLInv.Name = 'dataGridViewLInv'
    $dataGridViewLInv.DataBindings.DefaultDataSourceUpdateMode = 0
    $dataGridViewLInv.ReadOnly = $False
    $dataGridViewLInv.AllowUserToDeleteRows = $False
    $dataGridViewLInv.RowHeadersVisible = $false
    $dataGridViewLInv.ColumnHeadersVisible = $false
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 25
    $dataGridViewLInv.Size = $System_Drawing_Size
    $dataGridViewLInv.TabIndex = 8
    $dataGridViewLInv.Anchor = 15
    $dataGridViewLInv.AutoSizeColumnsMode = 'AllCells'
    $dataGridViewLInv.AllowUserToAddRows = $False
    $dataGridViewLInv.ColumnHeadersHeightSizeMode = 2
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 500
    $System_Drawing_Point.Y = 445
    $dataGridViewLInv.Location = $System_Drawing_Point
    $dataGridViewLInv.AllowUserToOrderColumns = $True
    $dataGridViewLInv.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
    $DataGridViewLInvAutoSizeColumnsMode.AllCells
    $datagridviewLInv.Add_KeyDown({if ($_.KeyCode -eq "Enter")
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
    $ip2TxtSQLQuery.Text = "192.168.1.1"
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
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btn35.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 110
    $System_Drawing_Point.Y = 45
    $btn35.Location = $System_Drawing_Point
    $btn35.add_Click($SelectIP_Click2)

    $form1.Controls.Add($btn35)
    #******************************************************************************************
    $ip2TxtSQLQuery353.Text = "Enter Device Name"
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
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btnRefi.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 15
    $btnRefi.Location = $System_Drawing_Point
    $btnRefi.add_Click({Refresh})

    $btnEdit.UseVisualStyleBackColor = $True
    $btnEdit.Text = 'Edit'
    $btnEdit.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnEdit.TabIndex = 1
    $btnEdit.Name = 'btnEdit'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 23
    $btnEdit.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 45
    $btnEdit.Location = $System_Drawing_Point
    $btnEdit.add_Click($Refi_Click)

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




