############################################################ Import PowerShell modules and Load Libraries ##############################################################################

Import-Module AdminToolbox.Exchange;
Import-Module ActiveDirectory;
[reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null;
[reflection.assembly]::LoadWithPartialName("System.Windows.Drawing") | Out-Null;
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
Add-Type -AssemblyName "PresentationFramework" | Out-Null;

################################################################# Load Config ##########################################################################################################

$config = Get-Content -Path C:\json.cfg | ConvertFrom-Json;

################################################################ Define Classes ########################################################################################################


Class newMailbox {

    $name;
    $aliases;
    $permissions;
    $UserPrincipalName;
    $primarySMTPAddress;
    $remoteRoutingAddress;
    $shared;
    $givenName;
    $surName;

    setName($name) {
        $this.name = $name;
    }

    setAliases($aliases) {
        $this.aliases = $aliases;
    }

    setPermissions($permissions) {
        $this.permissions = $permissions;
    }

    setPrimarySMTP($primarySMTPAddress) {
        $this.primarySMTPAddress = $primarySMTPAddress;
    }

    setRemoteRoudingAddress($remoteRoutingAddress) {
        $this.remoteRoutingAddress = $remoteRoutingAddress;
    }

    setUPN($UserPrincipalName) {
        $this.UserPrincipalName = $UserPrincipalName;
    }

    setGivenName($givenName) {
        $this.givenName = $givenName;
    }

    setSurName($surName) {
        $this.surName = $surName;
    }

    newMailbox($name, $UserPrincipalName, $permissions, $aliases, $primarySMTPAddress, $remoteRoutingAddress, $givenName, $surName) {
        $this.name = $name;
        $this.UserPrincipalName = $UserPrincipalName;
        $this.permissions = $permissions;
        $this.aliases = $aliases;
        $this.primarySMTPAddress = $primarySMTPAddress;
        $this.remoteRoutingAddress = $remoteRoutingAddress;
        $this.givenName = $givenName;
        $this.surName = $surName;
    }

}


Class newGroup {

    $name;
    $alias;
    $samid;
    $category;
    $scope;
    $members;
    $displayName;
    $SID;
    $mail;
    $mailNickname;
    $addresses


    setName($name) {
        $this.name = $name;
    }

    setAddresses($addresses) {
        $this.addresses = $addresses;
    }

    setPermissions($permissions) {
        $this.permissions = $permissions;
    }

    setPrimarySMTP($primarySMTPAddress) {
        $this.primarySMTPAddress = $primarySMTPAddress;
    }

    setRemoteRoudingAddress($remoteRoutingAddress) {
        $this.remoteRoutingAddress = $remoteRoutingAddress;
    }

    setUserPrincipalname($UserPrincipalName) {
        $this.UserPrincipalName = $UserPrincipalName;
    }

    setDisplayName($displayName) {
        $this.displayName = $displayName;
    }

    setSID($SID) {
        $this.SID = $SID;
    }

    setMail($mail) {
        $this.mail = $mail;
    }

    setAlias($alias) {
        $this.alias = $alias;
    }

    newGroup($name, $displayName, $samid, $alias, $addresses, $category, $scope, $members, $SID, $mail) {
        $this.name = $name;
        $this.samid = $samid;
        $this.alias = $alias;
        $this.addresses = $addresses;
        $this.category = $category;
        $this.scope = $scope;
        $this.members = $members;
        $this.displayName = $displayName;
        $this.SID = $SID;
        $this.mail = $mail;
    }

}

############################################################ Set global variables  #####################################################################################################

$global:ADC = $null;
$global:Source = $null;
$global:hybrid = $null;
$global:DSSO365 = $null;
$global:dataImport = @();
$global:mailboxes = @();
$global:newUsers = @();
$global:newMailboxes = @();
$global:failed = @();
$global:conflict = @();
$global:queue = @();
$global:inputSet = $false;
$global:skipped = @();
$global:newGroups = @();
$global:dataSource = $null;
$global:created = @();

############################################################################### Set up GUI ###############################################################################################

#Form Objects
$mainWindow = New-Object System.Windows.Forms.Form;
$objectsLabel = New-Object System.Windows.Forms.Label;
$objectsList = New-Object System.Windows.Forms.ListView;
$loadedLabel = New-Object System.Windows.Forms.Label;
$destinationLabel = New-Object System.Windows.Forms.Label;
$sourceLabel = New-Object System.Windows.Forms.Label;
$keyFieldLabel = New-Object System.Windows.Forms.Label;
$loadedPath = New-Object System.Windows.Forms.TextBox;
$loadDataButton = New-Object System.Windows.Forms.Button;
$startJobButton = New-Object System.Windows.Forms.Button;
$hybridCheck = New-Object System.Windows.Forms.CheckBox;
$cloudCheck = New-Object System.Windows.Forms.CheckBox;
$enableMB = New-Object System.Windows.Forms.CheckBox;
$userCheck = New-Object System.Windows.Forms.CheckBox;
$groupCheck = New-Object System.Windows.Forms.CheckBox;
$modeLabel = New-Object System.Windows.Forms.Label;
$CSVCheck = New-Object System.Windows.Forms.CheckBox;
$ADCheck = New-Object System.Windows.Forms.CheckBox;
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState;
$keyFieldComboBox = New-Object System.Windows.Forms.ComboBox;
$conflictsButton = New-Object System.Windows.Forms.Button;
$reloadButton = New-Object System.Windows.Forms.Button;
if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
    $sharedLabel = New-Object System.Windows.Forms.Label;
}

$startJobButtonClicked= 
{
        $user = $false;
        $group = $false;
        if (($userCheck.Checked) -and !($groupCheck.Checked)) {
            $user = $true;
            $group = $false;
            $enableOnly = $false;
            if ($enableMB.Checked -eq $true) {
                $enableOnly = $true;
            }
        } elseif (($groupCheck.Checked) -and !($userCheck.Checked)) {
            $group = $true;
            $user = $false;
        } else {
            [System.Windows.Messagebox]::Show("Please select a mode.");
        }
        if ($group -and !($user)) {
            Write-Host "Starting Group Mode.";
            Groups-Phase1;
        } elseif ($user -and !($group)) {
            Write-host "Starting User Mode.";
            if ($enableOnly) {
                Write-Host "Running in 'Enable Mailbox Only' mode.";
                Enable-Mailboxes;
            } else {
                Start-Phase1;
            }
        } else {
            [System.Windows.Messagebox]::Show("Please select a mode.");
        }
}

$testHandler= 
{
[System.Windows.Messagebox]::Show("Please choose export CSV")
    if ($hybridCheck.Checked -eq $true) {
            $import = Import-CSV -Path $global:dataImport;
            if ($global:csvSet -eq $false) { 
                [System.Windows.Messagebox]::Show("Please choose export CSV");
            } else {} 
    }

}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$mainWindow.WindowState = $InitialFormWindowState;
}


$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 400;
$System_Drawing_Size.Width = 579;
$mainWindow.ClientSize = $System_Drawing_Size;
$mainWindow.DataBindings.DefaultDataSourceUpdateMode = 0;
$mainWindow.Name = "mainWindow";
$mainWindow.Text = "Object Migration Utility";

$objectsLabel.DataBindings.DefaultDataSourceUpdateMode = 0;

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 356;
$System_Drawing_Point.Y = 57;
$objectsLabel.Location = $System_Drawing_Point;
$objectsLabel.Name = "objectsLabel";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 100;
$objectsLabel.Size = $System_Drawing_Size;
$objectsLabel.TabIndex = 4;
$objectsLabel.Text = "Objects";

$mainWindow.Controls.Add($objectsLabel);


$objectsList.DataBindings.DefaultDataSourceUpdateMode = 0;
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 221;
$System_Drawing_Point.Y = 83;
$objectsList.Location = $System_Drawing_Point;
$objectsList.Name = "objectsList";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 244;
$System_Drawing_Size.Width = 323;
$objectsList.Size = $System_Drawing_Size;
$objectsList.TabIndex = 3;
$objectsList.UseCompatibleStateImageBehavior = $False;
$objectsList.View = "Details";
$objectsList.GridLines = $True;
$objectsList.Columns.Add("Import", 120);

$mainWindow.Controls.Add($objectsList);

$loadedLabel.DataBindings.DefaultDataSourceUpdateMode = 0;

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 170;
$System_Drawing_Point.Y = 17;
$loadedLabel.Location = $System_Drawing_Point;
$loadedLabel.Name = "loadedLabel";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 45;
$loadedLabel.Size = $System_Drawing_Size;
$loadedLabel.TabIndex = 2;
$loadedLabel.Text = "loaded:";

$mainWindow.Controls.Add($loadedLabel);

$loadedPath.DataBindings.DefaultDataSourceUpdateMode = 0;
$loadedPath.Enabled = $False;
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 221;
$System_Drawing_Point.Y = 14;
$loadedPath.Location = $System_Drawing_Point;
$loadedPath.Name = "loadedPath";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 20;
$System_Drawing_Size.Width = 334;
$loadedPath.Size = $System_Drawing_Size;
$loadedPath.TabIndex = 1;

$mainWindow.Controls.Add($loadedPath);


$loadDataButton.DataBindings.DefaultDataSourceUpdateMode = 0;

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 57;
$System_Drawing_Point.Y = 11;
$loadDataButton.Location = $System_Drawing_Point;
$loadDataButton.Name = "loadCSVButton";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 75;
$loadDataButton.Size = $System_Drawing_Size;
$loadDataButton.TabIndex = 0;
$loadDataButton.Text = "Load Data";
$loadDataButton.UseVisualStyleBackColor = $True;
$loadDataButton.add_Click({Load-Data});

$mainWindow.Controls.Add($loadDataButton);

$startJobButton.DataBindings.DefaultDataSourceUpdateMode = 0;

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 57;
$System_Drawing_Point.Y = 41;
$startJobButton.Location = $System_Drawing_Point;
$startJobButton.Name = "startJobButton";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 75;
$startJobButton.Size = $System_Drawing_Size;
$startJobButton.TabIndex = 0;
$startJobButton.Text = "Execute";
$startJobButton.UseVisualStyleBackColor = $True;
$startJobButton.add_Click($startJobButtonClicked);
$startJobButton.Enabled = $false;

$mainWindow.Controls.Add($startJobButton);

$modeLabel.DataBindings.DefaultDataSourceUpdateMode = 0;

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 75;
$System_Drawing_Point.Y = 74;
$modeLabel.Location = $System_Drawing_Point;
$modeLabel.Name = "modeLabel";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 100;
$modeLabel.Size = $System_Drawing_Size;
$modeLabel.TabIndex = 4;
$modeLabel.Text = "Mode";

$mainWindow.Controls.Add($modeLabel);

$userCheck.Name = "User";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 34;
$System_Drawing_Size.Width = 70;
$userCheck.Size = $System_Drawing_Size;
$userCheck.Text = "User";
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 40;
$System_Drawing_Point.Y = 87;
$userCheck.Location = $System_Drawing_Point;
$userCheck.Add_Click({Switch-User});

$mainWindow.Controls.Add($userCheck);

$groupCheck.Name = "Group";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 34;
$System_Drawing_Size.Width = 70;
$groupCheck.Size = $System_Drawing_Size;
$groupCheck.Text = "Group";
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 110;
$System_Drawing_Point.Y = 87;
$groupCheck.Location = $System_Drawing_Point;
$groupCheck.Add_Click({Switch-Group});

$mainWindow.Controls.Add($groupCheck);

$sourceLabel.DataBindings.DefaultDataSourceUpdateMode = 0;

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 75;
$System_Drawing_Point.Y = 120;
$sourceLabel.Location = $System_Drawing_Point;
$sourceLabel.Name = "sourceLabel";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 100;
$sourceLabel.Size = $System_Drawing_Size;
$sourceLabel.TabIndex = 4;
$sourceLabel.Text = "Source";

$mainWindow.Controls.Add($sourceLabel);

$CSVCheck.Name = "CSV";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 34;
$System_Drawing_Size.Width = 70;
$CSVCheck.Size = $System_Drawing_Size;
$CSVCheck.Text = "CSV";
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 40;
$System_Drawing_Point.Y = 140;
$CSVCheck.Location = $System_Drawing_Point;
$CSVCheck.Add_Click({Switch-CSV});

$mainWindow.Controls.Add($CSVCheck);

$ADCheck.Name = "AD";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 34;
$System_Drawing_Size.Width = 70;
$ADCheck.Size = $System_Drawing_Size;
$ADCheck.Text = "AD";
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 110;
$System_Drawing_Point.Y = 140;
$ADCheck.Location = $System_Drawing_Point;
$ADCheck.Add_Click({Switch-AD});

$mainWindow.Controls.Add($ADCheck);

$destinationLabel.DataBindings.DefaultDataSourceUpdateMode = 0;

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 70;
$System_Drawing_Point.Y = 180;
$destinationLabel.Location = $System_Drawing_Point;
$destinationLabel.Name = "destinationLabel";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 100;
$destinationLabel.Size = $System_Drawing_Size;
$destinationLabel.TabIndex = 4;
$destinationLabel.Text = "Destination";

$mainWindow.Controls.Add($destinationLabel);


$hybridCheck.Name = "Hybrid";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 34;
$System_Drawing_Size.Width = 70;
$hybridCheck.Size = $System_Drawing_Size;
$hybridCheck.Text = "Local";
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 40;
$System_Drawing_Point.Y = 200;
$hybridCheck.Location = $System_Drawing_Point;
$hybridCheck.Add_Click({Switch-Hybrid});
$hybridCheck.Enabled = $false;

$mainWindow.Controls.Add($hybridCheck);


$cloudCheck.Name = "Cloud";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 34;
$System_Drawing_Size.Width = 70;
$cloudCheck.Size = $System_Drawing_Size;
$cloudCheck.Text = "Cloud";
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 110;
$System_Drawing_Point.Y = 200;
$cloudCheck.Location = $System_Drawing_Point;
$cloudCheck.Add_Click({Switch-Cloud});
$cloudCheck.Enabled = $false;

$mainWindow.Controls.Add($cloudCheck);

$enableMB.Name = "Enable Mailboxes Only";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 34;
$System_Drawing_Size.Width = 150;
$enableMB.Size = $System_Drawing_Size;
$enableMB.Text = "Enable Mailboxes Only";
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 40;
$System_Drawing_Point.Y = 235;
$enableMB.Location = $System_Drawing_Point;
$enableMB.Add_Click({Show-Warning});
$enableMB.Enabled = $false;

$mainWindow.Controls.Add($enableMB);

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 72;
$System_Drawing_Point.Y = 274;
$keyFieldLabel.Location = $System_Drawing_Point;
$keyFieldLabel.Name = "keyFieldLabel";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 100;
$keyFieldLabel.Size = $System_Drawing_Size;
$keyFieldLabel.TabIndex = 4;
$keyFieldLabel.Text = "Key Field";

$mainWindow.Controls.Add($keyFieldLabel);

$keyFieldComboBox.DataBindings.DefaultDataSourceUpdateMode = 0;
$keyFieldComboBox.FormattingEnabled = $True;
$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 40;
$System_Drawing_Point.Y = 300;
$keyFieldComboBox.Location = $System_Drawing_Point;
$keyFieldComboBox.Name = "keyFieldComboBox";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 21;
$System_Drawing_Size.Width = 121;
$keyFieldComboBox.Size = $System_Drawing_Size;
$keyFieldComboBox.TabIndex = 1;
$keyFieldComboBox.Text = "Please load data";
$keyFieldComboBox.AllowDrop = $true;
$keyFieldComboBox.DropDownStyle = "DropDownList"

$mainWindow.Controls.Add($keyFieldComboBox);

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 340;
$conflictsButton.Location = $System_Drawing_Point;
$conflictsButton.Name = "conflictsButton";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 200;
$conflictsButton.Size = $System_Drawing_Size;
$conflictsButton.TabIndex = 0;
$conflictsButton.Text = "Find Conflics";
$conflictsButton.UseVisualStyleBackColor = $True;
$conflictsButton.add_Click({Find-Conflicts});
$conflictsButton.Enabled = $false;

$mainWindow.Controls.Add($conflictsButton);

$System_Drawing_Point = New-Object System.Drawing.Point;
$System_Drawing_Point.X = 310
$System_Drawing_Point.Y = 340;
$reloadButton.Location = $System_Drawing_Point;
$reloadButton.Name = "reloadButton";
$System_Drawing_Size = New-Object System.Drawing.Size;
$System_Drawing_Size.Height = 23;
$System_Drawing_Size.Width = 200;
$reloadButton.Size = $System_Drawing_Size;
$reloadButton.TabIndex = 0;
$reloadButton.Text = "Reload w/o Conflicts";
$reloadButton.UseVisualStyleBackColor = $True;
$reloadButton.add_Click({Reload-Data});
$reloadButton.Enabled = $false;

$mainWindow.Controls.Add($reloadButton);

#Save the initial state of the form
$InitialFormWindowState = $mainWindow.WindowState;
#Init the OnLoad event to correct the initial state of the form
$mainWindow.add_Load($OnLoadForm_StateCorrection);

if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
    $sharedLabel.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point;
    $System_Drawing_Point.X = 88;
    $System_Drawing_Point.Y = 370;
    $sharedLabel.Location = $System_Drawing_Point;
    $sharedLabel.Name = "sharedLabel";
    $System_Drawing_Size = New-Object System.Drawing.Size;
    $System_Drawing_Size.Height = 23;
    $System_Drawing_Size.Width = 400;
    $sharedLabel.Size = $System_Drawing_Size;
    $sharedLabel.TabIndex = 4;
    $sharedLabel.ForeColor = [System.Drawing.Color]::FromArgb(255,255,0,0);
    $sharedLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",12,1,3,0)
    $sharedLabel.Text = "CURRENTLY RUNNING IN SHARED MODE!!!";
    $mainWindow.Controls.Add($sharedLabel);
}

function Load-Data {

    $CSV = $false;
    $AD = $false;
    if (($CSVCheck.Checked -eq $true) -and ($ADCheck.Checked -eq $false)) {
        $CSV = $true;
        $AD = $false;
        $global:dataSource = "CSV";
    } elseif (($ADCheck.Checked -eq $true) -and ($CSVCheck.Checked -eq $false)) {
        $CSV = $false;
        $AD = $true;
        $global:dataSource = "AD";
    } else {
        [System.Windows.Messagebox]::Show("Please choose a source");
    }
    if ($CSV -and !($AD)) {
        $fileIn = New-Object system.windows.forms.openfiledialog;
        $fileIn.ShowDialog();
        if ($fileIn.FileName -like "*.csv") {
            $global:dataImport = Import-CSV -Path $fileIn.FileName;
        } elseif ($fileIn.FileName -like "*.xml") {
            $global:dataImport = Import-Clixml -Path $fileIn.FileName;
        } else {
        
        }
        if ($global:dataImport -ne $null) {
            $loadedPath.Text = $fileIn.FileName;
            $import = $global:dataImport;
            $global:inputSet = $true;
            $buildObject = Get-Member -InputObject $import[0];
            $listViewBuild = @();
            $listViewItem = $null;
            foreach ($column in $buildObject) {
                if ($column.MemberType -eq "NoteProperty") {
                    $listViewBuild += $column.Name;
                    $objectsList.Columns.Add($column.Name, 120);
                    $keyFieldComboBox.Items.Add($column.Name);
                    $keyFieldComboBox.Text = $column.Name;
                }
            }
            foreach ($item in $import) {
                $listViewItem = New-Object System.Windows.Forms.ListViewItem;
                $listViewItem.Name = "Object";
                $listViewItem.Text = "Object";
                foreach ($listViewCol in $listViewBuild) {            
                    $listViewItem.SubItems.Add($item.($listViewCol));
                }
                $objectsList.Items.Add($listViewItem);
            }
            $cloudCheck.Enabled = $true;
            $hybridCheck.Enabled = $true;
            $conflictsButton.Enabled = $true;
            $CSVCheck.Enabled = $false;
            $ADCheck.Enabled = $false;
        }
    } elseif ($AD -and !($CSV)) {
        $user = $false;
        $group = $false;
        if (($userCheck.Checked) -and !($groupCheck.Checked)) {
            $user = $true;
            $group = $false;
        } elseif (($groupCheck.Checked) -and !($userCheck.Checked)) {
            $group = $true;
            $user = $false;
        } else {
            [System.Windows.Messagebox]::Show("Please select a mode.");
            break;
        }
        $baseDN = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the base DN you wish to use or leave empty for root:", "Enter Base DN");
        try { if (Get-ChildItem -Path C:\tempaddata.csv) { Remove-Item -Path C:\tempaddata.csv; }} catch {  }
        if ($user) {
            if ($baseDN) {
                $ADData = Get-ADUser -Filter * -SearchBase $baseDN -Server $config.sourceDC -Properties Name, SID, SamAccountName, ObjectClass, mail, mailNickName, GivenName, SurName, PrimaryGroup, primaryGroupID;
            } else {
                $ADData = Get-ADUser -Filter * -Server $config.sourceDC -Properties Name, SID, SamAccountName, ObjectClass, mail, mailNickName, GivenName, SurName, PrimaryGroup, primaryGroupID;
            }
        } elseif ($group) {
            if ($config.groupType -like "Distribution") {
                if ($baseDN) {
                    $ADData = Get-ADGroup -Filter { GroupCategory -eq "Distribution" } -SearchBase $baseDN -Server $config.sourceDC -Properties Name, SamAccountName, GroupCategory, GroupScope, mail, mailNickname, proxyAddresses, SID;
                } else {
                    $ADData = Get-ADGroup -Filter { GroupCategory -eq "Distribution" } -Server $config.sourceDC -Properties Name, SamAccountName, GroupCategory, GroupScope, mail, mailNickname, proxyAddresses, SID;
                }
            } elseif ($config.groupType -like "Security") {
                if ($baseDN) {
                    $ADData = Get-ADGroup -Filter { GroupCategory -eq "Security" } -SearchBase $baseDN -Server $config.sourceDC -Properties Name, SamAccountName, GroupCategory, GroupScope, proxyAddresses, SID;
                } else {
                    $ADData = Get-ADGroup -Filter { GroupCategory -eq "Security" } -Server $config.sourceDC -Properties Name, SamAccountName, GroupCategory, GroupScope, proxyAddresses, SID;
                }
            } else {
                [System.Windows.Messagebox]::Show("Invalid option in cfg file (groupType)");
            }
        }
        $ADData | Select * | Export-CSV -Path C:\tempaddata.csv -NoTypeInformation;
        $global:dataImport = Import-CSV -Path C:\tempaddata.csv;
        if ($global:dataImport -ne $null) {
            $loadedPath.Text = "Data Loaded from Active Directory.";
            $import = $global:dataImport;
            $global:inputSet = $true;
            $buildObject = Get-Member -InputObject $import[0];
            $listViewBuild = @();
            foreach ($column in $buildObject) {
                if ($column.MemberType -eq "NoteProperty") {
                    $listViewBuild += $column.Name;
                    $objectsList.Columns.Add($column.Name, 120);
                    $keyFieldComboBox.Items.Add($column.Name);
                    $keyFieldComboBox.Text = $column.Name;
                }
            }
            foreach ($item in $import) {
                $listViewItem = New-Object System.Windows.Forms.ListViewItem;
                $listViewItem.Name = "Object";
                $listViewItem.Text = "Object";
                foreach ($listViewCol in $listViewBuild) {            
                    $listViewItem.SubItems.Add($item.($listViewCol));
                }
                $objectsList.Items.Add($listViewItem);
            }
            $cloudCheck.Enabled = $true;
            $hybridCheck.Enabled = $true;
            $conflictsButton.Enabled = $true;
            $CSVCheck.Enabled = $false;
            $ADCheck.Enabled = $false;
        }
        try { if ($check = Get-ChildItem -Path C:\tempaddata.csv) { Remove-Item -Path C:\tempaddata.csv; }} catch { Write-Host "Could not delete file at C:\tempaddata.csv"; }
    }
    
}

Function Reload-Data {

    $CSV = $false;
    $AD = $false;
    $global:dataImport = @();
    $objectsList.Clear();
    $objectsList.Columns.Add("Import", 120);
    if (($CSVCheck.Checked -eq $true) -and ($ADCheck.Checked -eq $false)) {
        $CSV = $true;
        $AD = $false;
    } elseif (($ADCheck.Checked -eq $true) -and ($CSVCheck.Checked -eq $false)) {
        $CSV = $false;
        $AD = $true;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a source");
    }
    if ($CSV -and !($AD)) {
        $global:dataImport = Import-CSV -Path C:\noConflicts.csv;
        if ($global:dataImport -ne $null) {
            $loadedPath.Text = "C:\noConflicts.csv";
            $import = $global:dataImport;
            $global:inputSet = $true;
            $buildObject = Get-Member -InputObject $import[0];
            $listViewBuild = @();
            $listViewItem = $null;
            foreach ($column in $buildObject) {
                if ($column.MemberType -eq "NoteProperty") {
                    $listViewBuild += $column.Name;
                    $objectsList.Columns.Add($column.Name, 120);
                }
            }
            foreach ($item in $import) {
                $listViewItem = New-Object System.Windows.Forms.ListViewItem;
                $listViewItem.Name = "Object";
                $listViewItem.Text = "Object";
                foreach ($listViewCol in $listViewBuild) {            
                    $listViewItem.SubItems.Add($item.($listViewCol));
                }
                $objectsList.Items.Add($listViewItem);
            }
            $cloudCheck.Enabled = $true;
            $hybridCheck.Enabled = $true;
            $conflictsButton.Enabled = $true;
        }
    } elseif ($AD -and !($CSV)) { 
        $global:dataImport = Import-CSV -Path C:\noConflicts.csv;
        if ($global:dataImport -ne $null) {
            $loadedPath.Text = "Data Loaded from Active Directory.";
            $import = $global:dataImport;
            $global:inputSet = $true;
            $buildObject = Get-Member -InputObject $import[0];
            $listViewBuild = @();
            foreach ($column in $buildObject) {
                if ($column.MemberType -eq "NoteProperty") {
                    $listViewBuild += $column.Name;
                    $objectsList.Columns.Add($column.Name, 120);
                }
            }
            foreach ($item in $import) {
                $listViewItem = New-Object System.Windows.Forms.ListViewItem;
                $listViewItem.Name = "Object";
                $listViewItem.Text = "Object";
                foreach ($listViewCol in $listViewBuild) {            
                    $listViewItem.SubItems.Add($item.($listViewCol));
                }
                $objectsList.Items.Add($listViewItem);
            }
            $cloudCheck.Enabled = $true;
            $hybridCheck.Enabled = $true;
            $conflictsButton.Enabled = $true;
        }
    }
    $startJobButton.Enabled = $true;
}

function Find-Conflicts {
    Login-Hybrid;

    $user = $false;
    $group = $false;
    $conflicts = @();
    $noConflicts = @();
    if (($userCheck.Checked) -and !($groupCheck.Checked)) {
        $user = $true;
        $group = $false;
    } elseif (($groupCheck.Checked) -and !($userCheck.Checked)) {
        $group = $true;
        $user = $false;
    } else {

    }
    if ($group -and !($user)) {
        Write-Host "Checking for conflicting Groups.";
        foreach ($group in $global:dataImport) {
            if ($group.($keyFieldComboBox.Text) -like "*@*") {
                $nameArray = ($group.($keyFieldComboBox.Text)).Split("@");
                $groupName = $nameArray[0];
            } else {
                $groupName = $group.($keyFieldComboBox.Text);
            }
            $search = $null;
            $search = Get-ADGroup -Identity $groupName -Server $config.destinationDC -Properties *;
            if ($search) {
                $conflicts += $search;
            } elseif (!$search) {
                $noConflicts += $group;
            }
        }
        try { if ($check = Get-ChildItem -Path C:\conflicts.csv) { Remove-Item -Path C:\conflicts.csv; }} catch { Write-Host "Could not delete file at C:\conlicts.csv"; }
        try { if ($check = Get-ChildItem -Path C:\noConflicts.csv) { Remove-Item -Path C:\noConflicts.csv; }} catch { Write-Host "Could not delete file at C:\noConflicts.csv"; }
        $conflicts | Select * | Export-CSV -Path C:\conflicts.csv -NoTypeInformation;
        $noConflicts | Select * | Export-CSV -Path C:\noConflicts.csv -NoTypeInformation;
        if ($conflicts.Count -gt 0) {
            [System.Windows.Messagebox]::Show("Conflicts Found: `nConflicts are located at C:\conflicts.csv.");
            $reloadButton.Enabled = $true;
        } else {
            [System.Windows.Messagebox]::Show("No conflicts found");
            $startJobButton.Enabled = $true;
        }
    } elseif ($user -and !($group)) {
        Write-host "Checking for conflicting Users.";
        foreach ($account in $global:dataImport) {
            if ($account.($keyFieldComboBox.Text) -like "*@*") {
                $nameArray = ($account.($keyFieldComboBox.Text)).Split("@");
                $accountName = $nameArray[0];
            } else {
                $accountName = $account.($keyFieldComboBox.Text);
            }
            $oldAccount = Get-ADUser -Identity $accountName -Properties Name, SID, SamAccountName, ObjectClass, mail, mailNickName, GivenName, SurName, PrimaryGroup, primaryGroupID -Server $config.sourceDC;
            if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
                $newName = $accountName;
            } else {
                $newName = (($oldAccount.GivenName.SubString(0,1))+$oldAccount.SurName).toLower();
            }
            $search = $null;
            $search = Get-ADUser -Identity $newName -Properties Name, SID, SamAccountName, ObjectClass, mail, mailNickName, GivenName, SurName, PrimaryGroup, primaryGroupID -Server $config.destinationDC;
            $search2 = Get-RemoteMailbox -Identity $newName"@primowater.com";
            if (($search) -or ($search2)) {
                $conflicts += $search;
            } elseif ((!$search) -or (!$search2)) {
                $noConflicts += $account;
            }

        }
        try { if ($check = Get-ChildItem -Path C:\conflicts.csv) { Remove-Item -Path C:\conflicts.csv; }} catch { Write-Host "Could not delete file at C:\conlicts.csv"; }
        try { if ($check = Get-ChildItem -Path C:\noConflicts.csv) { Remove-Item -Path C:\noConflicts.csv; }} catch { Write-Host "Could not delete file at C:\noConflicts.csv"; }
        $conflicts | Select * | Export-CSV -Path C:\conflicts.csv -NoTypeInformation;
        $noConflicts | Select * | Export-CSV -Path C:\noConflicts.csv -NoTypeInformation;
        if ($conflicts.Count -gt 0) {
            [System.Windows.Messagebox]::Show("Conflicts Found: `nConflicts are located at C:\conflicts.csv.");
            $reloadButton.Enabled = $true;
        } else {
            [System.Windows.Messagebox]::Show("No conflicts found");
            $startJobButton.Enabled = $true;
        }
    } else {
        [System.Windows.Messagebox]::Show("Please select a mode.");
    }
    Logout-Hybrid;

}

function Switch-Cloud {

    if ($global:inputSet -eq $false) {
        [System.Windows.Messagebox]::Show("Please load data");
        $cloudCheck.Checked = $false;
    } else {
        $hybridCheck.Checked = $false;
        if ($groupCheck.Checked -eq $false) {
            $enableMB.Enabled = $true;
        }
    }
    $cloudCheck.Enabled = $false;
    $hybridCheck.Enabled = $true;

}

function Switch-Hybrid {

    if ($global:inputSet -eq $false) {
        [System.Windows.Messagebox]::Show("Please load data");
        $hybridCheck.Checked = $false;
    } else {
        $cloudCheck.Checked = $false;
        $enableMB.Checked = $false;
        $enableMB.Enabled = $false;
    }
    $hybridCheck.Enabled = $false;
    $cloudCheck.Enabled = $true;

}

function Show-Warning {

    if ($global:inputSet -eq $false) {
        [System.Windows.Messagebox]::Show("Please load data");
    } else {
        if ($groupCheck.Checked -eq $true) {
            [System.Windows.Messagebox]::Show("This option does not work with group mode!");
            $enableMB.Checked = $false;
        } else {
            [System.Windows.Messagebox]::Show("Warning: This option can only be used with accounts which already exist in the destination");
        }
    }

}

function Switch-AD {

	$CSVCheck.Checked = $false;
    $ADCheck.Enabled = $false;
    $CSVCheck.Enabled = $true;

}

function Switch-CSV {

	$ADCheck.Checked = $false;
    $CSVCheck.Enabled = $false;
    $ADCheck.Enabled = $true;

}

function Switch-Group {

	$userCheck.Checked = $false;
    if ($enableMB.Checked -eq $true) {
        [System.Windows.Messagebox]::Show("The 'Enable Mailboxes Only' option does not work with group mode!");
        $enableMB.Checked = $false;
    }
    $enableMB.Enabled = $false;
    $groupCheck.Enabled = $false;
    $userCheck.Enabled = $true;

}

function Switch-User {

	$groupCheck.Checked = $false;
    if ($enableMB.Checked -eq $true) {
        [System.Windows.Messagebox]::Show("Note: Enable Mailboxes Only is currently checked.");
    }
    if ($cloudCheck.Checked) {
        $enableMB.Enabled = $true;
    }
    $userCheck.Enabled = $false;
    $groupCheck.Enabled = $true;

}

################################################################## Start-Phase1 (Phase One) #################################################################################################

Function Start-Phase1 {

    $timestamp = Get-Date;
    $timestamp | Out-File -FilePath C:\users.log -Append -NoClobber;
    
    #### Prompt user for credentials for legacy Primo, DSW on-prem, and DSS Office 365 ####
    if (!$global:sourceCreds){
        $global:sourceCreds = Get-Credential -Credential $config.sourcePrefix;
    }
    if (!$global:hybridCreds){
        $global:hybridCreds = Get-Credential -Credential "DSW\$env:USERNAME";
    }
    if (!$global:DSSCreds) {
        $global:DSSCreds = Get-Credential -Credential "@dsservices.onmicrosoft.com";
    }
    
    #### create new variables to hold the mailbox objects ####

    $mailbox = $null;
    $newUser = $null;
    $newUPN = $null;
    $sourceADAccount = $null;
    $oldUPN = $null;
    $oldUPNSuffix = $null;
    $dswAccount = $null;
    $account = $null;
    $hybrid = $false;
    $cloud = $false;
    if (($hybridCheck.Checked -eq $true) -and ($cloudCheck.Checked -eq $false)) {
        $hybrid = $true;
        $cloud = $false;
    } elseif (($cloudCheck.Checked -eq $true) -and ($hybridCheck.Checked -eq $false)) {
        $hybrid = $false;
        $cloud = $true;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }

    #### loop through each mailbox and get the mailbox from the Source Exchange server ####

    if ($hybrid -and !($cloud)) {

        Login-Source;
        foreach ($account in $global:dataImport) {
            $newName = $null;
            $oldUPN = $null;
            $oldUPNSuffix = $null;
            $newUPN = $null;
            $aliases = @();
            $primarySMTPAddress = $null;
            $remoteRoutingAddress = $null;
            $givenName = $null;
            $surName = $null;
            $newMailbox = $null;
            $permissions = @();
            if ($account.($keyFieldComboBox.Text) -like "*@*") {
                $nameArray = ($account.($keyFieldComboBox.Text)).Split("@");
                $mbName = $nameArray[0];
            } else {
                $mbName = $account.($keyFieldComboBox.Text);
            }
            $mailbox = Get-Mailbox -Identity $mbName;
            if ($mailbox.Count -gt 1) {
                $global:failed += $account.($keyFieldComboBox.Text);
                Write-Host "Something went wrong with: ($keyFieldComboBox.Text)."
                break;
            } else {
                $primarySMTPAddress = $mailbox.primarySMTPAddress;
                $permissions = Get-MailboxPermission -Identity $primarySmtpAddress;
                $permissions += Get-Mailbox -Identity $primarySmtpAddress | Get-ADPermission;
                $sourceADAccount = Get-ADUser -Identity $mbName -Properties Name, SID, SamAccountName, ObjectClass, mail, mailNickName, GivenName, SurName, PrimaryGroup, primaryGroupID -Server $config.sourceDC;
                if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
                    $newName = $mbName;
                    $oldUPN = $sourceADAccount.UserPrincipalName.Split("@");
                    $oldUPNSuffix = "@"+$oldUPN[1];
                    $newUPN = $sourceADAccount.UserPrincipalName.Replace($oldUPNSuffix, "@primowater.com");
                } else {
                    $newName = (($sourceADAccount.GivenName.SubString(0,1))+$sourceADAccount.SurName).toLower();
                    $newUPN = $newName+"@primowater.com";
                }
                $aliases = $mailbox.EmailAddresses;
                if (!(@($aliases) -like "*"+$primarySMTPAddress)) {
                    $aliases.Add("smtp:"+$primarySMTPAddress);
                }
                if (!(@($aliases) -like "*"+$newUPN)) {
                    $aliases.Add("SMTP:"+$newUPN);
                }
                $remoteRoutingAddress = $newUPN.Replace("@primowater.com", "@dsservices.onmicrosoft.com");
                $givenName = $sourceADAccount.GivenName;
                $surName = $sourceADAccount.Surname;
                $newMailbox = [newMailbox]::new($newName, $newUPN, $permissions, $aliases, $primarySMTPAddress, $remoteRoutingAddress, $givenName, $surName);
                $global:newMailboxes += $newMailbox;
            }
        }
        Logout-Source;
        Start-Phase2;

    } elseif ($cloud -and !($hybrid)) {

        Login-Source;
        Connect-AzureAD;
        foreach ($account in $global:dataImport) {
            $newName = $null;
            $oldUPN = $null;
            $oldUPNSuffix = $null;
            $newUPN = $null;
            $aliases = @();
            $primarySMTPAddress = $null;
            $remoteRoutingAddress = $null;
            $givenName = $null;
            $surName = $null;
            $newMailbox = $null;
            $permissions = @();
            if ($account.($keyFieldComboBox.Text) -like "*@*") {
                $nameArray = ($account.($keyFieldComboBox.Text)).Split("@");
                $mbName = $nameArray[0];
            } else {
                $mbName = $account.($keyFieldComboBox.Text);
            }
            $mailbox = Get-Mailbox -Identity $mbName;
            if ($mailbox.Count -gt 1) {
                $global:failed += $account.($keyFieldComboBox.Text);
                Write-Host "Something went wrong with: ($keyFieldComboBox.Text)."
                break;
            } else {
                $primarySMTPAddress = $mailbox.primarySMTPAddress;
                $permissions = Get-MailboxPermission -Identity $primarySmtpAddress;
                $permissions += Get-Mailbox -Identity $primarySmtpAddress | Get-ADPermission;
                $sourceADAccount = Get-ADUser -Identity $mbName -Properties Name, SID, SamAccountName, ObjectClass, mail, mailNickName, GivenName, SurName, PrimaryGroup, primaryGroupID -Server $config.sourceDC;
                if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
                    $newName = $mbName;
                    $oldUPN = $sourceADAccount.UserPrincipalName.Split("@");
                    $oldUPNSuffix = "@"+$oldUPN[1];
                    $newUPN = $sourceADAccount.UserPrincipalName.Replace($oldUPNSuffix, "@primowater.com");
                } else {
                    $newName = (($sourceADAccount.GivenName.SubString(0,1))+$sourceADAccount.SurName).toLower();
                    $newUPN = $newName+"@primowater.com";
                }
                $aliases = $mailbox.EmailAddresses;
                if (!(@($aliases) -like "*"+$primarySMTPAddress)) {
                    $aliases.Add("smtp:"+$primarySMTPAddress);
                }
                if (!(@($aliases) -like "*"+$newUPN)) {
                    $aliases.Add("SMTP:"+$newUPN);
                }
                $remoteRoutingAddress = $newUPN.Replace("@primowater.com", "@dsservices.onmicrosoft.com");
                $givenName = $sourceADAccount.GivenName;
                $surName = $sourceADAccount.Surname;
                $newMailbox = [newMailbox]::new($newName, $newUPN, $permissions, $aliases, $primarySMTPAddress, $remoteRoutingAddress, $givenName, $surName);
                $global:newMailboxes += $newMailbox;;
            } 
        }
        Logout-Source;
        Start-Phase2;

    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }
}

############################################################################### End Phase One ############################################################################################

################################################################################# Phase Two ##############################################################################################

Function Start-Phase2 {
    
    Get-Date | Out-File -FilePath C:\users.log -Append -NoClobber;
    $global:newMailboxes | Select * | Export-Clixml -Path C:\newUsers.xml;
    $message = "New mailboxes to be created exported to: C:\newUsers.xml";
    Write-Host $message;
    $message | Out-File -FilePath C:\users.log -Append -NoClobber;

    $hybrid = $false;
    $cloud = $false;
    if (($hybridCheck.Checked -eq $true) -and ($cloudCheck.Checked -eq $false)) {
        $hybrid = $true;
        $cloud = $false;
    } elseif (($cloudCheck.Checked -eq $true) -and ($hybridCheck.Checked -eq $false)) {
        $hybrid = $false;
        $cloud = $true;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }

    if ($hybrid -and !($cloud)) {

        Login-Hybrid;
        $newMailbox = $null;
        foreach ($newMailbox in $global:newMailboxes) {
            $dswAccount = $null;
            $mailbox = $null;
            $aliases = @();
            $mailbox = $newMailbox;
            $dswAccount = Get-ADUser -Identity $mailbox.UserPrincipalname.Replace("@primowater.com", "") -Server $config.destinationDC;
                if (!$dswAccount) {
                    if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
                        New-RemoteMailbox -Shared -Name $mailbox.Name -UserPrincipalName $mailbox.UserPrincipalName -PrimarySmtpAddress $mailbox.UserPrincipalName -RemoteRoutingAddress $mailbox.remoteRoutingAddress -OnPremisesOrganizationalUnit $config.OU;
                        $message = "Creating new remote mailbox for: " + $mailbox.UserPrincipalName;
                        Write-Host $message;
                        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                    } elseif (($config.shared -eq "false") -or ($config.shared -eq "no")) {
                        New-RemoteMailbox -Name $mailbox.Name -UserPrincipalName $mailbox.UserPrincipalName -PrimarySmtpAddress $mailbox.UserPrincipalName -RemoteRoutingAddress $mailbox.remoteRoutingAddress -OnPremisesOrganizationalUnit $config.OU;
                        $message = "Creating new remote mailbox for: " + $mailbox.UserPrincipalName;
                        Write-Host $message;
                        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                    } else {
                        [System.Windows.Messagebox]::Show("Invalid option in cfg file (shared)");
                    }
                    Start-Sleep -Seconds 60;
                    try {
                        $global:queue += Get-RemoteMailbox -Identity $mailbox.UserPrincipalName;
                    } catch {
                        $global:failed += $mailbox;
                    }
                    Start-Sleep -Seconds 60;
                    try {
                        $dswAccount = Get-ADUser -Identity $mailbox.UserPrincipalName.Replace("@primowater.com", "") -Server $config.destinationDC;
                    } catch {
                        $global:failed += $mailbox;
                    }
                } else {
                    Write-Host "Account already exists";
                    $global:conflict += $mailbox;
                }
            $aliases = $mailbox.aliases;
            foreach ($address in $aliases) {
                $contact = $null;
                $contact2 = $null;
                $contact = Get-Contact -Identity $address;
                $contact2 = Get-MailContact -Identity $address;
                if ($contact) {
                    $backupContacts += $contact;
                    Remove-MailContact -Identity $contact.Identity;
                    $message = "Removing contact: " + $contact.Name;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                }
                if ($contact2) {
                    $backupContacts += $contact2;
                    Remove-MailContact -Identity $contact2.Identity;
                    $message = "Removing contact: " + $contact2.Name;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                }
                if (($address.contains("smtp:")) -or ($address.contains("SMTP:"))) {
                    Set-RemoteMailbox -Identity $mailbox.UserPrincipalName -EmailAddresses @{add=$address};
                    $message = "Adding address: " + $address + " to mailbox: " + $mailbox.UserPrincipalName;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                } else {
                    Set-ADUser -Identity $dswAccount -add @{proxyAddresses=$address} -Server $config.destinationDC -Credential $global:hybridCreds;
                    $message = "Adding address: " + $address + " to mailbox: " + $mailbox.UserPrincipalName;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                }
            }
            if ($backupContacts) {
                $message = "Contacts backed up to C:\backupContacts.xml";
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                $backupContacts | Select * | Export-Clixml -Path C:\bacupContacts.xml;
            }
        }
        Logout-Hybrid;
        Sync-ADC;
        $message = "Syncing AADConnect";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
        Start-Sleep -Seconds 60;
        Login-DSSO365;
        try {
            while (!(Get-Mailbox -Identity $mailbox.UserPrincipalName.Replace("@primowater.com", ""))) {
                Write-Host "Awaiting creation in Exchange Online.";
                Start-Sleep -Seconds 60;
            }
        } catch { Write-Host "Awaiting creation in Exchange Online."; }
        Start-Phase3;

    } elseif ($cloud -and !($hybrid)) {\

        $newMailbox = $null;
        $mailbox = $null;
        Login-DSSO365;
        $defaultPW = ConvertTo-SecureString -String $config.defaultPW -AsPlainText -Force
        foreach ($newMailbox in $global:newMailboxes) {
            $newADUser = $null;
            $mailbox = $null;
            $mailbox = $newMailbox;
            if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
                $newADUser = New-AzureADUser -DisplayName $mailbox.name -UserPrincipalName $mailbox.UserPrincipalName -GivenName $mailbox.givenName -Surname $mailbox.surName -AccountEnabled $false -MailNickName $mailbox.name;
                $message = "Creating on-prem AD account for: " + $mailbox.UserPrincipalName;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            } elseif (($config.shared -eq "false") -or ($config.shared -eq "no")) {
                $newADUser = New-AzureADUser -DisplayName $mailbox.name -UserPrincipalName $mailbox.UserPrincipalName -GivenName $mailbox.givenName -Surname $mailbox.surName -AccountEnabled $true -MailNickName $mailbox.name;
                $message = "Creating on-prem AD account for: " + $mailbox.UserPrincipalName;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            } else {
                [System.Windows.Messagebox]::Show("Invalid option in cfg file (shared)");
            }
            if ($newADUser) {
                $global:queue += $newMailbox;
            } else {
                $global:failed += $mailbox.name;
            }
        }
        $planName = $config.license;
        $location = $config.location;
        $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense;
        $license.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID;
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses;
        $licenses.AddLicenses = $license;
        try {
            while (!(Get-AzureADuser -SearchString $mailbox.UserPrincipalName)){
                Write-Host "Awaiting replication to Exchange Online.";
                Start-Sleep -Seconds 60;
            }
        } catch { Write-Host "Something went wrong."; }
        foreach ($newMailbox in $global:newMailboxes) {
           $mailbox = $null;
           $mailbox = $newMailbox;
           $aliases = @();
           $AADUser = $null;
           $AADUser = Get-AzureADUser -SearchString $mailbox.UserPrincipalName;
           Set-AzureADUser -ObjectID $AADUser.ObjectID -UsageLocation $location;
           $message = "Setting user: " + $AADUser.DisplayName + " location to: " + $location;
           Write-Host $message;
           $message | Out-File -FilePath C:\users.log -Append -NoClobber;
           if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
                Enable-Mailbox -Identity $mailbox.UserPrincipalName;
                Set-Mailbox -Identity $mailbox.UserPrincipalName -Type Shared;
                $message = "Enabling mailbox: " + $mailbox.UserPrincipalName + " and setting to shared mailbox.";
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            } elseif (($config.shared -eq "false") -or ($config.shared -eq "no")) {
                Enable-Mailbox -Identity $mailbox.UserPrincipalName;
                $message = "Enabling mailbox for: " + $mailbox.UserPrincipalName;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                Set-AzureADUserLicense -ObjectId $AADUser.ObjectId -AssignedLicenses $licenses;
                $message = "Setting license for: " + $mailbox.UserPrincipalName + " to: " + $license;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            } else {
                [System.Windows.Messagebox]::Show("Invalid option in cfg file (shared)");
                break;
            }
            $aliases = $mailbox.aliases;
            foreach ($address in $aliases) {
                $contact = $null;
                $contact2 = $null;
                $contact = Get-Contact -Identity $address;
                $contact2 = Get-MailContact -Identity $address;
                if ($contact) {
                    $backupContacts += $contact;
                    Remove-MailContact -Identity $contact.Identity;
                    $message =  "Removing contact for: " + $contact.Name;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                }
                if ($contact2) {
                    $backupContacts += $contact2;
                    Remove-MailContact -Identity $contact2.Identity;
                    $message =  "Removing contact for: " + $contact2.Name;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                }
                if (($address.contains("smtp:")) -or ($address.contains("SMTP:"))) {
                    Set-Mailbox -Identity $mailbox.UserPrincipalName -EmailAddresses @{add=$address};
                    $message = "Adding address: " + $address + " to mailbox: " + $mailbox.UserPrincipalName;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                } else {
                    Set-ADUser -Identity $mailbox.UserPrincipalName -add @{proxyAddresses=$address} -Server $config.destinationDC -Credential $global:hybridCreds;
                    $message = "Adding address: " + $address + " to mailbox: " + $mailbox.UserPrincipalName;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                }
            }
            if ($backupContacts) {
                $message = "Contacts backed up to C:\backupContacts.xml";
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                $backupContacts | Select * | Export-Clixml -Path C:\bacupContacts.xml;
            }
        }

    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }
}

############################################################################### End Phase Two ############################################################################################

################################################################################# Phase Three ############################################################################################

Function Start-Phase3 {

    $hybrid = $false;
    $cloud = $false;
    if (($hybridCheck.Checked -eq $true) -and ($cloudCheck.Checked -eq $false)) {
        $hybrid = $true;
        $cloud = $false;
    } elseif (($cloudCheck.Checked -eq $true) -and ($hybridCheck.Checked -eq $false)) {
        $hybrid = $false;
        $cloud = $true;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }

    if ($hybrid -and !($cloud)) {
        
        $newMailbox = $null;
        $DSWMailbox = $null;
        $mbObject = $null;
        $planName = $config.license;
        $location = $config.location;
        $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense;
        $license.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID;
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses;
        $licenses.AddLicenses = $license;
        foreach ($newMailbox in $global:queue) {
            $DSWMailbox = $null;
            $mbObject = $null;
            $DSWMailbox = $newMailbox;
            $AADUser = Get-AzureADUser -SearchString $DSWMailbox.UserPrincipalName;
            Set-AzureADUser -ObjectId $AADUser.ObjectID -UsageLocation $location;
            $message = "Setting user: " + $DSWMailbox.UserPrincipalName + " location to: " + $location;
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            if (($config.shared -notlike "true") -or ($config.shared -notlike "yes")) {
                Set-AzureADUserLicense -ObjectId $AADUser.ObjectId -AssignedLicenses $licenses;
                $message = "Setting user: " + $DSWMailbox.UserPrincipalName + " license";
            }
            $mbObject = $global:newMailboxes | Where-Object { $_.name -eq $DSWMailbox.Name };
            if (!$mbObject) {
                Write-Host "Something went wrong.";
                $global:failed += $DSWMailbox.name;
            } elseif (($DSWMailbox -isnot [array]) -and $mbObject) { 
               Set-ADUser -Identity $DSWMailbox.SamAccountName -EmailAddress $DSWMailbox.PrimarySmtpAddress -Server $config.destinationDC -Credential $global:hybridCreds;
               $message = "Setting AD User: " + $DSWMailbox.SamAccountName + " email address to: " + $DSWMailbox.PrimarySMTPAddress;
               Write-Host $message;
               $message | Out-File -FilePath C:\users.log -Append -NoClobber;
               $permissions = $mbObject.permissions;
               foreach ($permission in $permissions) {
                    if ($permission.AccessRights -like "FullAccess") {
                            $user = $permission.User;
                            if ($user.Contains("\")) {
                                $user = $user.Split("\");
                                $user = $user[1] + "@primowater.com";
                                Add-MailboxPermission -Identity $DSWMailbox.PrimarySmtpAddress -User $user -AccessRights "FullAccess" -Confirm:$false;
                                Add-RecipientPermission -Identity $DSWMailbox.PrimarySmtpAddress -Trustee $user -AccessRights "SendAs" -Confirm:$false;
                                $message = "Adding permission to: " + $DSWMailbox.PrimarySMTPAddress + " for user: " + $user;
                                Write-Host $message;
                                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                            } else {
                                Write-Host "Skipping user: $user";
                            }
                        if ($permission.AccessRights -like "ExtendedRight") {
                            $user = $permission.User;
                            if ($user.Contains("\")) {
                                $user = $user.Split("\");
                                $user = $user[1] + "@primowater.com";
                                Add-RecipientPermission -Identity $DSWMailbox.PrimarySmtpAddress -Trustee $user -AccessRights "SendAs" -Confirm:$false;
                                $message = "Adding permission to: " + $DSWMailbox.PrimarySMTPAddress + " for user: " + $user;
                                Write-Host $message;
                                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                            } else {
                                $message =  "Skipping user: " + $user + " on mailbox: " + $DSWMailbox.PrimarySMTPAddress;
                                Write-Host $message;
                                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                            }
                        }
                    }
                }
            } else {
                Write-Host "Something went wrong, more than one mailbox selected.";
            }
            $global:created += $newMailbox;
        }
        $global:created | Select * | Export-Clixml -Path C:\createdUsers.xml;
        $message = "Created users exported to: C:\createdUsers.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    
    } elseif ($cloud -and !($hybrid)) {

        $newMailbox = $null;
        $cloudMailbox= $null;
        $mbObject = $null;
        foreach ($newMailbox in $global:queue) {
            $cloudMailbox= $null;
            $mbObject = $null;
            $cloudMailbox = $newMailbox;
            $mbObject = $global:newMailboxes | Where-Object { $_.primarySMTPAddress -eq $cloudMailbox.PrimarySmtpAddress }
            if (!$mbObject) {
                Write-Host "Something went wrong.";
                $global:failed += $cloudMailbox.name;
            } elseif (($cloudMailbox-isnot [array]) -and $mbObject) { 
               Set-ADUser -Identity $cloudMailbox.SamAccountName -EmailAddress $cloudMailbox.PrimarySmtpAddress -Server $config.destinationDC;
               $message = "Setting email address for user: " + $cloudMailbox.SamAccountName + " to: " + $cloudMailbox.PrimarySMTPAddress;
               Write-Host $message;
               $message | Out-File -FilePath C:\users.log -Append -NoClobber;
               $permissions = $mbObject.permissions;
               foreach ($permission in $permissions) {
                    if ($permission.AccessRights -like "FullAccess") {
                            $user = $permission.User;
                            if ($user.Contains("\")) {
                                $user = $user.Split("\");
                                $user = $user[1] + "@primowater.com";
                                Add-MailboxPermission -Identity $cloudMailbox.PrimarySmtpAddress -User $user -AccessRights "FullAccess" -Confirm:$false;
                                Add-RecipientPermission -Identity $cloudMailbox.PrimarySmtpAddress -Trustee $user -AccessRights "SendAs" -Confirm:$false;
                                $message = "Adding permission to: " + $cloudMailbox.PrimarySMTPAddress + " for user: " + $user;
                                Write-Host $message;
                                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                            } else {
                                Write-Host "Skipping user: $user";
                            }
                        if ($permission.AccessRights -like "ExtendedRight") {
                            $user = $permission.User;
                            if ($user.Contains("\")) {
                                $user = $user.Split("\");
                                $user = $user[1] + "@primowater.com";
                                Add-RecipientPermission -Identity $cloudMailbox.PrimarySmtpAddress -Trustee $user -AccessRights "SendAs" -Confirm:$false;
                                $message = "Adding permission to: " + $cloudMailbox.PrimarySMTPAddress + " for user: " + $user;
                                Write-Host $message;
                                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                            } else {
                                Write-Host "Skipping user: $user";
                            }
                        }
                    }
                }
            } else {
                Write-Host "Something went wrong, more than one mailbox selected.";
            }
            $global:created += $newMailbox;
        }
        $global:created | Select * | Export-Clixml -Path C:\createdUsers.xml;
        $message = "Created users exported to: C:\createdUsers.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }
    Logout-DSSO365;
    Login-Source;
    $newMailbox = $null;
    foreach ($newMailbox in $global:queue) {
        $mailbox = $null;
        $mailbox = $newMailbox;
        try {
            New-MailContact -Name $mailbox.remoteRoutingAddress.Trim("SMTP:") -ExternalEmailAddress $mailbox.remoteRoutingAddress.Trim("SMTP:") -Alias $mailbox.userPrincipalName.Replace("@primowater.com", "") -OrganizationalUnit $path;
            $message = "Creating contact for: " + $mailbox.remoteRoutingAddress + " on source server.";
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            Set-Mailbox -Identity $mailbox.userPrincipalName.Replace("@primowater.com", "") -DeliverToMailboxAndForward $true -ForwardingAddress $mailbox.remoteRoutingAddress.Trim("SMTP:") -Confirm:$false;
            $message = "Setting forward for: " + $mailbox.userPrincipalName + " on source server to: " + $mailbox.remoteRoutingAddress;
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
        } catch { Write-Host "Unable to add forward for: "$mailbox.name; }
    }
    Logout-Source;
    if ($global:failed) {
        $global:failed | Select * | Export-Clixml -Path C:\failedUsers.xml;
        $message = "Failed users exported to: C:\failedUsers.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    }
    if ($global:conflict) {
        $global:conflict | Select * | Export-Clixml -Path C:\conflictUsers.xml;
        $message = "Conflicting users exported to: C:\conflictUsers.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    }
    $timestamp = Get-Date;
    $message = "Creation completed: " + $timestamp;
    Write-Host $message
    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    [System.Windows.Messagebox]::Show("Job Completed!");
}

############################################################################## End Phase Three ############################################################################################

############################################################################ Begin Groups-Phase1 Function #################################################################################

Function Groups-Phase1 {

    $timestamp = Get-Date;
    $timestamp | Out-File -FilePath C:\groups.log -Append -NoClobber;

    #### Prompt user for credentials for source, DSW on-prem, and DSS Office 365 ####
    if (!$global:sourceCreds){
        $global:sourceCreds = Get-Credential -Credential $config.sourcePrefix;
    }
    if (!$global:hybridCreds){
        $global:hybridCreds = Get-Credential -Credential "DSW\$env:USERNAME";
    }
    if (!$global:DSSCreds) {
        $global:DSSCreds = Get-Credential -Credential "@dsservices.onmicrosoft.com";
    }
    
    #### create new variables to hold the group objects ####

    $groupName = $null;
    $group = $null;
    $hybrid = $false;
    $cloud = $false;

    if (($hybridCheck.Checked -eq $true) -and ($cloudCheck.Checked -eq $false)) {
        $hybrid = $true;
        $cloud = $false;
    } elseif (($cloudCheck.Checked -eq $true) -and ($hybridCheck.Checked -eq $false)) {
        $hybrid = $false;
        $cloud = $true;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }

    if ($cloud -and ($config.groupType -like "Security")) {
        [System.Windows.Messagebox]::Show('Security groups not available as cloud-only resource.`nPlease change config file to "Distribution"');
        break;
    }

    if ($hybrid -and !($cloud)) {

        Login-Source;
        foreach ($account in $global:dataImport) {
            $aliases = $null;
            $samid = $null;
            $scope = $null;
            $category = $null;
            $newGroup = $null;
            $members = @();
            if ($account.($keyFieldComboBox.Text) -like "*@*") {
                $nameArray = ($account.($keyFieldComboBox.Text)).Split("@");
                $groupName = $nameArray[0];
            } else {
                $groupName = $account.($keyFieldComboBox.Text);
            }
            switch ($config.groupType) {
                "Security" {
                    $group = Get-ADGroup -Identity $groupName -Server $config.sourceDC -Properties *;
                    break;
                }
                "Distribution" {
                    $group = Get-DistributionGroup -Identity $groupName;
                    break;
                }
                Default {
                    [System.Windows.Messagebox]::Show("Invalid option in config (groupType)");
                    $global:failed += $account.($keyFieldComboBox.Text);
                    break;
                }
            }
            if ($group.Count -gt 1) {
                $global:failed += $account.($keyFieldComboBox.Text);
                Write-Host "Something went wrong with: ($account.keyFieldComboBox.Text)."
                break;
            } else {
                $groupName = $group.Name;
                if ($group.GroupCategory -like "Security") {
                    $addresses = "N/A";
                    $mail = "N/A";
                    $alias = "N/A"
                    $members = Get-ADGroupMember -Identity $groupName -Server $config.sourceDC;
                    $category = "Security";
                } else {
                    $addresses = $group.EmailAddresses;
                    if (!(@($addresses) -like "*primowater.com")) {
                        $addresses.Add("SMTP:"+$group.mailNickname+"@primowater.com");
                    }
                    $mail = $group.WindowsEmailAddress;
                    $alias = $group.Alias;
                    $members = Get-DistributionGroupMember -Identity $groupName;
                    $category = "Distribution";
                }
                $displayName = $group.DisplayName;
                $samid = $group.SamAccountName;
                $scope = $group.GroupType;
                $sid = $group.objectSid;
                $newGroup = [newGroup]::new($groupName, $displayName, $samid, $alias, $addresses, $category, $scope, $members, $SID, $mail);
                $global:newGroups += $newGroup;
            }
        }
        Logout-Source;
        $global:newGroups | Select * | Export-Clixml -Path C:\newGroups.xml;
        $message = "Queue of groups to be created exported to: C:\newGroups.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
        Groups-Phase2;
    }

    elseif ($cloud -and !($hybrid)) {

        Login-Source;
        Connect-AzureAD;
        foreach ($account in $global:dataImport) {
            $aliases = $null;
            $samid = $null;
            $scope = $null;
            $category = $null;
            $newGroup = $null;
            $members = @();
            if ($account.($keyFieldComboBox.Text) -like "*@*") {
                $nameArray = ($account.($keyFieldComboBox.Text)).Split("@");
                $groupName = $nameArray[0];
            } else {
                $groupName = $account.($keyFieldComboBox.Text);
            }
            switch ($config.groupType) {
                "Security" {
                    $group = Get-ADGroup -Identity $groupName -Server $config.sourceDC -Properties *;
                    break;
                }
                "Distribution" {
                    $group = Get-DistributionGroup -Identity $groupName;
                    break;
                }
                Default {
                    [System.Windows.Messagebox]::Show("Invalid option in config (groupType)");
                    $global:failed += $account.($keyFieldComboBox.Text);
                    break;
                }
            }
            if ($group.Count -gt 1) {
                $global:failed += $account.($keyFieldComboBox.Text);
                Write-Host "Something went wrong with: " $account.($keyFieldComboBox.Text)".";
                break;
            } else {
                $groupName = $group.Name;
                if ($group.GroupCategory -like "Security") {
                    $addresses.Add("N/A");
                    $mail = "N/A";
                    $alias = "N/A"
                    $members = Get-ADGroupMember -Identity $groupName -Server $config.sourceDC;
                    $category = "Security";
                } else {
                    $addresses = $group.ProxyAddresses;
                    if (!(@($addresses) -like "*primowater.com")) {
                        $addresses.Add("SMTP:"+$group.mailNickname+"@primowater.com");
                    }
                    $mail = $group.WindowsEmailAddress;
                    $alias = $group.Alias;
                    $members = Get-DistributionGroupMember -Identity $groupName;
                    $category = "Distribution";
                }
                $displayName = $group.DisplayName;
                $samid = $group.SamAccountName;
                $scope = $group.Scope;
            }
            $newGroup = [newGroup]::new($groupName, $displayName, $samid, $alias, $addresses, $category, $scope, $members, $SID, $mail);
            $global:newGroups += $newGroup;
        }
        Logout-Source;
        $global:newGroups | Select * | Export-Clixml -Path C:\newGroups.xml;
        $message = "Queue of groups to be created exported to: C:\newGroups.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
        Groups-Phase2;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }
}

###################################################################### End Groups-Phase1 Function ###################################################################################

###################################################################### Begin Groups-Phase2 Function #################################################################################

Function Groups-Phase2 {

    $hybrid = $false;
    $cloud = $false;
    $group = $null;
    if (($hybridCheck.Checked -eq $true) -and ($cloudCheck.Checked -eq $false)) {
        $hybrid = $true;
        $cloud = $false;
    } elseif (($cloudCheck.Checked -eq $true) -and ($hybridCheck.Checked -eq $false)) {
        $hybrid = $false;
        $cloud = $true;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }

    if ($hybrid -and !($cloud)) {

        $group = $null;
        $path = $config.OU;
        $destination = $config.destinationDC;
        Login-Hybrid;

        foreach ($group in $global:newGroups) {
            $newADGroup = $null;
            $name = $null;
            $scope = $null;
            $members = @();
            $aliases = @();
            $category = $null;
            $displayName = $null;
            $alias = $null;
            $name = $group.name;
            $smtp = $group.mail;
            $alias = $group.alias;
            $scope = $group.scope;
            $members = $group.members;
            $addresses = $group.addresses;
            $samid = $group.samid;
            $category = $group.category;
            $displayName = $group.displayName;
            if ($category -like "Distribution") {
                $newADGroup = New-DistributionGroup -Name $name -Alias $alias -SamAccountName $samid -DisplayName $displayName -PrimarySMTPAddress $smtp -OrganizationalUnit $path;
                $message =  "Creating new Distribution Group: " + $name;
                Write-Host $message;
                $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                if ($newADGroup) {
                    $global:queue += $newGroup;
                    foreach ($member in $members) {
                        if ($member.PrimarySmtpAddress -notlike "*@primowater.com") {
                            $contact = $null;
                            $contact2 = $null;
                            $contact = Get-Contact -Identity $member.PrimarySMTPAddress;
                            $contact2 = Get-MailContact -Identity $member.PrimarySMTPAddress;
                            if ((!$contact) -and (!$contact2)) {
                                $contact = New-MailContact -Name $member.Name -ExternalEmailAddress $member.PrimarySmtpAddress -OrganizationalUnit $path;
                                if ($contact) { 
                                    $message =  "Created mail contact: " + $member.Name;
                                    Write-Host $message;
                                    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                                } else {
                                    $message =  "Unable to create contact: " + $member.Name;
                                    Write-Host $message;
                                    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                                }
                            }
                        }
                        try {
                            Add-DistributionGroupMember -Identity $name -Member $member.primarySMTPAddress;
                            $message = "Adding member: " + $member.SamAccountName + " to group: " + $name;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        } catch { Write-Host "Failed to add member: " $member; }
                    }
                    foreach ($address in $addresses) {
                        $contact = $null;
                        $contact2 = $null;
                        $contact = Get-Contact -Identity $address.Substring(5);
                        $contact2 = Get-MailContact -Identity $address.Substring(5);
                        if ($contact) {
                            $backupContacts += $contact;
                            Remove-MailContact -Identity $contact.Identity;
                            $message =  "Removing contact for address:  " + $contact.Name;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        }
                        if ($contact2) {
                            $backupContacts += $contact2;
                            Remove-MailContact -Identity $contact2.Identity;
                            $message = "Removing contact for address: " + $contact2.Name;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        }
                        if ($address -clike "SMTP:*primowater.com") {
                            $address = $address;
                        } elseif ($address -clike "smtp:*primowater.com") {
                            $address = $address.Replace("smtp:", "SMTP:");
                        } elseif (($address -clike "SMTP:*") -and ($address -cnotlike "*primowater.com")) {
                            $address = $address.Replace("SMTP:", "smtp:");
                        } else {
                            $address = $address;
                        }
                        Set-DistributionGroup -Identity $name -EmailAddresses @{add=$address};
                        $message = "Adding address: " + $address + " to group: " + $name;
                        Write-Host $message;
                        $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                    }
                    Set-DistributionGroup -Identity $name -PrimarySMTPAddress $smtp;
                    $message = "Setting primary address for group: " + $name + " to: " + $alias;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                } else {
                    $global:failed += $group.name;
                }
            } elseif ($category -like "Security") {
                $newADGroup = New-ADGroup -Name $name -GroupCategory $category -GroupScope $scope -Path $path -SamAccountName $samid -Server $destination -DisplayName $displayName -Credential $global:hybridCreds -;
                $message =  "Creating new Distribution Group: " + $name;
                Write-Host $message;
                $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                if ($newADGroup) {
                    $global:queue += $newADGroup;
                    foreach ($member in $members) {
                        try {
                            Add-ADGroupMember -Identity $name -Members $member.SamAccountName -Server $destination -Credential $global:hybridCreds;
                            $message = "Adding member: " + $member.SamAccountName + " to group: " + $name;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        } catch { Write-Host "Failed to add member: " $member; }
                    }
                    Set-ADGroup -Identity $name -replace @{mail=$name+"@primowater.com"};
                    $message = "Setting group: " + $name + " email address to: " + $name+"@primowater.com";
                    Write-Host $message;
                    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                 } else {
                    $global:failed = $group.name;
                 }
            }
            $global:created += $group;
       }      
       Sync-ADC;
       $message = "Syncing AADConnect";
       Write-Host $message;
       $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
       $global:created | Select * | Export-Clixml -Path C:\createdGroups.xml;
       $message = "Created groups exported to C:\createdGroups.xml";
       Write-Host $message;
       $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
    }

    elseif ($cloud -and !($hybrid)) {
        Login-DSSO365;

        foreach ($group in $global:newGroups) {
            $newADGroup = $null;
            $name = $null;
            $scope = $null;
            $members = @();
            $alias = $group.mail;
            $category = $null;
            $displayName = $null;
            $name = $group.name;
            $scope = $group.scope;
            $members = $group.members;
            $addresses = $group.addresses;
            $samid = $group.samid;
            $category = $group.category;
            if ($category -like "Security") {
                $newADGroup = New-AzureADGroup -DisplayName $name -MailNickName $samid -MailEnabled $false;
                $message =  "Creating new Distribution Group: " + $name;
                Write-Host $message;
                $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                if ($newADGroup) {
                    $global:queue += $newGroup;
                    foreach ($member in $members) {
                        try {
                            Add-AzureADGroupMember -Identity $name -Members $member.SamAccountName;
                            $message = "Adding member: " + $member.SamAccountName + " to group: " + $name;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        } catch { Write-Host "Failed to add member: " $member; }
                    }
                } else {
                    $global:failed += $group.name;
                }
            } elseif ($category -like "Distribution") {
                $newADGroup = New-DistributionGroup -Name $name -Alias $alias -SamAccountName $samid -DisplayName $displayName;
                $message =  "Creating new Distribution Group: " + $name;
                Write-Host $message;
                $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                if ($newADGroup) {
                    $global:queue += $newGroup;
                    foreach ($member in $members) {
                        if ($member.PrimarySmtpAddress -notlike "*@primowater.com") {
                            $contact = $null;
                            $contact2 = $null;
                            $contact = Get-Contact -Identity $member.PrimarySMTPAddress;
                            $contact2 = Get-MailContact -Identity $member.PrimarySMTPAddress;
                            if ((!$contact) -and (!$contact2)) {
                                $contact = New-MailContact -Name $member.Name -ExternalEmailAddress $member.PrimarySmtpAddress;
                                if ($contact) {
                                    $message =  "Created mail contact: " + $member.Name;
                                    Write-Host $message;
                                    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                                } else {
                                    $message =  "Unable to create contact: " + $member.Name;
                                    Write-Host $message;
                                    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                                }
                            }
                        }
                        try {
                            Add-DistributionGroupMember -Identity $name -Member $member.primarySMTPAddress;
                            $message = "Adding member: " + $member.SamAccountName + " to group: " + $name;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        } catch { Write-Host "Failed to add member: " $member; }
                    }
                    $count = 0;
                    foreach ($address in $addresses) {
                        $contact = $null;
                        $contact2 = $null;
                        $contact = Get-Contact -Identity $address.Substring(5);
                        $contact2 = Get-MailContact -Identity $address.Substring(5);
                        if ($contact) {
                            $backupContacts += $contact;
                            Remove-MailContact -Identity $contact.Identity;
                            $message = "Remove-MailContact " + $contact.Name;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        }
                        if ($contact2) {
                            $backupContacts += $contact2;
                            Remove-MailContact -Identity $contact2.Identity;
                            $message = "Remove-MailContact " + $contact2.Name;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                        }
                        if ($address -clike "SMTP:*primowater.com") {
                            $address = $address;
                        } elseif ($address -clike "smtp:*primowater.com") {
                            $address = $address.Replace("smtp:", "SMTP:");
                        } elseif (($address -clike "SMTP:*") -and ($address -cnotlike "*primowater.com")) {
                            $address = $address.Replace("SMTP:", "smtp:");
                        } else {
                            $address = $address;
                        }
                        Set-DistributionGroup -Identity $name -EmailAddresses @{add=$address};
                        $message = "Adding address: " + $address + " to group: " + $name;
                        Write-Host $message;
                        $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                    }
                    Set-DistributionGroup -PrimarySMTPAddress $alias;
                    $message = "Setting primary address for group: " + $name + " to: " + $alias;
                    Write-Host $message;
                    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
                } else {
                    $global:failed += $group.name;
                    Write-Host "incorrect Group type specified.";
                }
                
            } else {
                [System.Windows.Messagebox]::Show("Incorrect option in config file (groupType)");
            }
            $global:created += $group;
        }
        $global:created | Select * | Export-Clixml -Path C:\createdGroups.xml;
        $message = "Created groups exported to C:\createdGroups.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
    } else {
        [System.Windows.Messagebox]::Show("Please choose a destination");
    }
    if ($global:failed) {
        $global:failed | Select * | Export-Clixml -Path C:\failedGroups.xml;
        $message = "Failed groups exported to: C:\failedGroups.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\groups.log -Append -NoClobber;
    }
    Groups-Phase3;
}

###################################################################### End Groups-Phase2 Function ###################################################################################

###################################################################### Begin Groups-Phase3 Function #################################################################################

Function Groups-Phase3 {
    
    $timestamp = Get-Date;
    $message = "Creation completed: " + $timestamp;
    Write-Host $message
    $message | Out-File -FilePath C:\groups.log -Append -NoClobber;

    [System.Windows.Messagebox]::Show("job complete");

 <# $newGroup = $null;
    $groupObject = $null;
    foreach ($newGroup in $global:queue) {
        $groupObject = $null;
        $group = $null;
        $group = $newGroup;
        $groupObject = $global:newGroups | Where-Object { $_.name -eq $newGroup.Name };
        if (!$groupObject) {
            Write-Host "Something went wrong.";
            $global:failed += $newGroup.name;
        }
    } #>
}

################################################################ Enable Mailboxes Only (Phase 1) ######################################################################################

Function Enable-Mailboxes {

    $timestamp = Get-Date;
    $timestamp | Out-File -FilePath C:\users.log -Append -NoClobber;

    #### Prompt user for credentials for source, DSW on-prem, and DSS Office 365 ####
    if (!$global:sourceCreds){
        $global:sourceCreds = Get-Credential -Credential $config.sourcePrefix;
    }
    if (!$global:hybridCreds){
        $global:hybridCreds = Get-Credential -Credential "DSW\$env:USERNAME";
    }
    if (!$global:DSSCreds) {
        $global:DSSCreds = Get-Credential -Credential "@dsservices.onmicrosoft.com";
    }

    Login-Source;
    Connect-AzureAD;
    Login-DSSO365;
    foreach ($account in $global:dataImport) {
        $newName = $null;
        $oldUPN = $null;
        $oldUPNSuffix = $null;
        $newUPN = $null;
        $aliases = @();
        $primarySMTPAddress = $null;
        $remoteRoutingAddress = $null;
        $givenName = $null;
        $surName = $null;
        $newMailbox = $null;
        $permissions = @();
        if ($account.($keyFieldComboBox.Text) -like "*@*") {
            $nameArray = ($account.($keyFieldComboBox.Text)).Split("@");
            $mbName = $nameArray[0];
        } else {
            $mbName = $account.($keyFieldComboBox.Text);
        }
        $mailbox = Get-Mailbox -Identity $mbName;
        if ($mailbox.Count -gt 1) {
            $global:failed += $account.($keyFieldComboBox.Text);
            Write-Host "Something went wrong with: ($keyFieldComboBox.Text)."
            break;
        } else {
            $primarySMTPAddress = $mailbox.primarySMTPAddress;
            $permissions = Get-MailboxPermission -Identity $primarySmtpAddress;
            $permissions += Get-Mailbox -Identity $primarySmtpAddress | Get-ADPermission;
            $sourceADAccount = Get-ADUser -Identity $mbName -Properties Name, SID, SamAccountName, ObjectClass, mail, mailNickName, GivenName, SurName, PrimaryGroup, primaryGroupID -Server $config.sourceDC;
            if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
                $newName = $mbName;
                $oldUPN = $sourceADAccount.UserPrincipalName.Split("@");
                $oldUPNSuffix = "@"+$oldUPN[1];
                $newUPN = $sourceADAccount.UserPrincipalName.Replace($oldUPNSuffix, "@primowater.com");
            } else {
                $newName = (($sourceADAccount.GivenName.SubString(0,1))+$sourceADAccount.SurName).toLower();
                $newUPN = $newName+"@primowater.com";
            }
            $aliases = $mailbox.EmailAddresses;
            if (!(@($aliases) -like "*"+$primarySMTPAddress)) {
                $aliases.Add("smtp:"+$primarySMTPAddress);
            }
            if (!(@($aliases) -like "*"+$newUPN)) {
                $aliases.Add("SMTP:"+$newUPN);
            }
            $remoteRoutingAddress = $newUPN.Replace("@primowater.com", "@dsservices.onmicrosoft.com");
            $givenName = $sourceADAccount.GivenName;
            $surName = $sourceADAccount.Surname;
            $newMailbox = [newMailbox]::new($newName, $newUPN, $permissions, $aliases, $primarySMTPAddress, $remoteRoutingAddress, $givenName, $surName);
            $global:newMailboxes += $newMailbox;
        }
    }
    Logout-Source;
    Enable-Phase2;
    $message = "Groups to be created exported to C:\newGroups.xml";
    Write-Host $message;
    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    $global:newMailboxes | Select * | Export-Clixml -Path C:\newGroups.xml;
}

################################################################ End Enable-Mailboxes (Phase 1) #####################################################################################

################################################################ Enable Mailboxes Only (Phase 2) ####################################################################################

Function Enable-Phase2 {

    $planName = $config.license;
    $location = $config.location;
    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense;
    $license.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID;
    $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses;
    $licenses.AddLicenses = $license;
        
    foreach ($newMailbox in $global:newMailboxes) {
        $mailbox = $null;
        $mailbox = $newMailbox;
        $aliases = @();
        $AADUser = $null;
	<# try {
            while (!(Get-AzureADuser -SearchString $mailbox.UserPrincipalName)){
            $message =  "Awaiting replication to Exchange Online or account: " + $mailbox.UserPrincipalName + " does not exist.";
	Write-Host $message;
	$message | Out-File -FilePath C:\users.log -Append -NoClobber;
            Start-Sleep -Seconds 60;
            }
        } catch { Write-Host "Something went wrong."; } #>
        $AADUser = Get-AzureADUser -SearchString $mailbox.UserPrincipalName;
        Set-AzureADUser -ObjectID $AADUser.ObjectID -UsageLocation $location;
        $message = "Setting user: " + $AADUser.DisplayName + " location to: " + $location;
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
        if (($config.shared -eq "true") -or ($config.shared -eq "yes")) {
            Enable-Mailbox -Identity $mailbox.UserPrincipalName;
            Set-Mailbox -Identity $mailbox.UserPrincipalName -Type Shared;
            $message = "Setting mailbox: " + $mailbox.UserPrincipalName + " to shared mailbox.";
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
        } elseif (($config.shared -eq "false") -or ($config.shared -eq "no")) {
            Enable-Mailbox -Identity $mailbox.UserPrincipalName;
            $message = "Enabling mailbox for: " + $mailbox.UserPrincipalName;
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            Set-AzureADUserLicense -ObjectId $AADUser.ObjectId -AssignedLicenses $licenses;
            $message = "Setting license for: " + $mailbox.UserPrincipalName + " to: " + $license;
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
        } else {
            [System.Windows.Messagebox]::Show("Invalid option in cfg file (shared)");
            break;
        }
        $aliases = $mailbox.aliases;
        foreach ($address in $aliases) {
            $contact = $null;
            $contact2 = $null;
            $contact = Get-Contact -Identity $address;
            $contact2 = Get-MailContact -Identity $address;
            if ($contact) {
                $backupContacts += $contact;
                Remove-MailContact -Identity $contact.Identity;
                $message =  "Removing contact for: " + $contact.Name;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            }
            if ($contact2) {
                $backupContacts += $contact2;
                Remove-MailContact -Identity $contact2.Identity;
                $message =  "Removing contact for: " + $contact2.Name;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            }
            if (($address.contains("smtp:")) -or ($address.contains("SMTP:"))) {
                Set-Mailbox -Identity $mailbox.UserPrincipalName -EmailAddresses @{add=$address};
                $message = "Adding address: " + $address + " to mailbox: " + $mailbox.UserPrincipalName;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            } else {
                Set-Mailbox -Identity $mailbox.UserPrincipalName -EmailAddresses @{add=$address};
                $message = "Adding address: " + $address + " to mailbox: " + $mailbox.UserPrincipalName;
                Write-Host $message;
                $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            }
        }
        if ($AADUser) {
            $global:queue += $newMailbox;
        } else {
            $global:failed += $mailbox.name;
        }
    }
    $message = "Mailboxes to be enabled exported to C:\enabled.xml";
    Write-Host $message;
    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    $global:queue | Select * | Export-Clixml -Path C:\enabled.xml;
    if ($global:failed) {
        $message = "Failed mailboxes exported to C:\failed.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
        $global:queue | Select * | Export-Clixml -Path C:\failed.xml;
    }
    Enable-Phase3;
}

################################################################ End Enable-Mailboxes (Phase 2) #####################################################################################

Function Enable-Phase3 {

    $newMailbox = $null;
    $cloudMailbox= $null;
    $mbObject = $null;
    foreach ($newMailbox in $global:queue) {
        $cloudMailbox= $null;
        $mbObject = $null;
        $cloudMailbox = $newMailbox;
        $mbObject = $global:newMailboxes | Where-Object { $_.primarySMTPAddress -eq $cloudMailbox.PrimarySmtpAddress }
        if (!$mbObject) {
            Write-Host "Something went wrong.";
            $global:failed += $cloudMailbox.name;
        } elseif (($cloudMailbox-isnot [array]) -and $mbObject) { 
            Set-Mailbox -Identity $cloudMailbox.SamAccountName -PrimarySMTPAddress $cloudMailbox.PrimarySmtpAddress;
            $message = "Setting email address for user: " + $cloudMailbox.SamAccountName + " to: " + $cloudMailbox.PrimarySMTPAddress;
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            $permissions = $mbObject.permissions;
            foreach ($permission in $permissions) {
                if ($permission.AccessRights -like "FullAccess") {
                        $user = $permission.User;
                        if ($user.Contains("\")) {
                            $user = $user.Split("\");
                            $user = $user[1] + "@primowater.com";
                            Add-MailboxPermission -Identity $cloudMailbox.PrimarySmtpAddress -User $user -AccessRights "FullAccess" -Confirm:$false;
                            Add-RecipientPermission -Identity $cloudMailbox.PrimarySmtpAddress -Trustee $user -AccessRights "SendAs" -Confirm:$false;
                            $message = "Adding permission to: " + $cloudMailbox.PrimarySMTPAddress + " for user: " + $user;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                        } else {
                            Write-Host "Skipping user: $user";
                        }
                    if ($permission.AccessRights -like "ExtendedRight") {
                        $user = $permission.User;
                        if ($user.Contains("\")) {
                            $user = $user.Split("\");
                            $user = $user[1] + "@primowater.com";
                            Add-RecipientPermission -Identity $cloudMailbox.PrimarySmtpAddress -Trustee $user -AccessRights "SendAs" -Confirm:$false;
                            $message = "Adding permission to: " + $cloudMailbox.PrimarySMTPAddress + " for user: " + $user;
                            Write-Host $message;
                            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
                        } else {
                            Write-Host "Skipping user: $user";
                        }
                    }
                }
            }
        } else {
            Write-Host "Something went wrong, more than one mailbox selected.";
        }
        $global:created += $newMailbox;
    }
    $global:created | Select * | Export-Clixml -Path C:\createdUsers.xml;
    $message = "Created users exported to: C:\createdUsers.xml";
    Write-Host $message;
    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
   
    Logout-DSSO365;
    Login-Source;
    $newMailbox = $null;
    foreach ($newMailbox in $global:queue) {
        $mailbox = $null;
        $mailbox = $newMailbox;
        try {
            New-MailContact -Name $mailbox.remoteRoutingAddress.Trim("SMTP:") -ExternalEmailAddress $mailbox.remoteRoutingAddress.Trim("SMTP:") -Alias $mailbox.userPrincipalName.Replace("@primowater.com", "") -OrganizationalUnit $path;
            $message = "Creating contact for: " + $mailbox.remoteRoutingAddress + " on source server.";
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
            Set-Mailbox -Identity $mailbox.userPrincipalName.Replace("@primowater.com", "") -DeliverToMailboxAndForward $true -ForwardingAddress $mailbox.remoteRoutingAddress.Trim("SMTP:") -Confirm:$false;
            $message = "Setting forward for: " + $mailbox.userPrincipalName + " on source server to: " + $mailbox.remoteRoutingAddress;
            Write-Host $message;
            $message | Out-File -FilePath C:\users.log -Append -NoClobber;
        } catch { Write-Host "Unable to add forward for: "$mailbox.name; }
    }
    Logout-Source;
    if ($global:failed) {
        $global:failed | Select * | Export-Clixml -Path C:\failedUsers.xml;
        $message = "Failed users exported to: C:\failedUsers.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    }
    if ($global:conflict) {
        $global:conflict | Select * | Export-Clixml -Path C:\conflictUsers.xml;
        $message = "Conflicting users exported to: C:\conflictUsers.xml";
        Write-Host $message;
        $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    }
    $timestamp = Get-Date;
    $message = "Creation completed: " + $timestamp;
    Write-Host $message;
    $message | Out-File -FilePath C:\users.log -Append -NoClobber;
    [System.Windows.Messagebox]::Show("Job Completed!");

}

#################################################### Login in to Source Exchange Server using Exchange configuration ################################################################

Function Login-Source {

    $global:Source = $null;
    if ($global:sourceCreds -eq $null) {
        $global:sourceCreds = Get-Credential -Credential $config.sourcePrefix;
    }
    $global:Source = New-PSSession -ConnectionURI $config.sourceURI -Credential $global:sourceCreds -ConfigurationName Microsoft.Exchange;
    Import-PSSession $global:Source -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null;

}

###################################################### Login in to DSW Office 365 using Exchange configuration  ######################################################################

Function Login-DSSO365 {

    $global:DSSO365 = $null;
    if ($global:DSSCreds -eq $null) {
        $global:DSSCreds = Get-Credential -Credential "@dsservices.onmicrosoft.com";
    }
    $global:DSSO365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $config.destinationURI -Credential $global:DSSCreds -Authentication Basic -AllowRedirection;
    Import-PSSession $global:DSSO365 -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null;

}

######################################################## Login in to DSW on prem using Exchange configuration  ########################################################################

Function Login-Hybrid {

    $global:hybrid = $null;
    if ($global:hybridCreds -eq $null) {
        $global:hybridCreds = Get-Credential -Credential "$env:USERDOMAIN\$env:USERNAME";
    }
    $global:hybrid = New-PSSession -ConnectionURI http://atmexhyb01.dsw.net/PowerShell -Credential $global:hybridcreds -ConfigurationName Microsoft.Exchange;
    Import-PSSession $global:hybrid -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue | Out-Null;

}

################################################### Login in to DSW ADConnect server and sync new accounts #############################################################################

Function Sync-ADC {

    $global:ADC = $null;
    if ($global:hybridCreads = $null) {
        $global:hybridCreds = Get-Credential -Credential "$env:USERDOMAIN\$env:USERNAME";
    }
    $global:ADC = New-PSSession -ComputerName "ATMDSYC01" -Credential $global:hybridCreds;
    Invoke-Command -Session $global:ADC -ScriptBlock { Start-ADSyncSyncCycle };
    Remove-PSSession $global:ADC;

}


############################################################### Log off of Source Exchange server #######################################################################################
Function Logout-Source {

    Remove-PSSession $global:Source;

}

################################################################### Log off of DSW Office365 ############################################################################################

Function Logout-DSSO365 {

    Remove-PSSession $global:DSSO365;

}

############################################################### Log off of DSW ADConnect Server #########################################################################################

Function Logout-ADC {

    Remove-PSSession $global:ADC;

}

############################################################ Log off of on-prem Hybrid Exchange Server ##################################################################################

Function Logout-Hybrid {

    Remove-PSSession $global:hybrid;

}

$mainWindow.ShowDialog()| Out-Null;