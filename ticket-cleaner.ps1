cd $PSScriptRoot;
$global:config = Import-Csv "ticket-cleaner-config.csv";
$global:config.scsm_path = $global:config.scsm_path + "\PowerShell\System.Center.Service.Manager.psd1";
Import-Module $global:config.scsm_path;
Add-Type -AssemblyName System.Windows.Forms;

$global:browser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('MyDocuments') 
    Filter = 'SpreadSheet (*.csv)|*.csv'
}
$global:saveFile = New-Object System.Windows.Forms.SaveFileDialog -Property @{
    InitialDirectory = $PSScriptRoot
    Filter = 'SpreadSheet (*.csv)|*.csv'
}
$global:_menu = @('Import CSV', 'Generate CSV of old tickets', 'Generate CSV of resolved/completed tickets', 'Run CSV', 'Help', 'Quit');
$global:menu_options = @(1, 2, 3, 4, 5, 'q');
$global:current_csv = [pscustomobject]@{path='';obj=''};
$global:max_age_days = 30;
[System.Collections.ArrayList]$failed = @();
$global:class_instances = @('mna', 'rwa', 'srq', 'inc');
$global:today = Get-Date;

function get-class($class_name) {
    return Get-SCSMClass -ComputerName $global:config.server -Name $class_name;
}

$global:classes = [pscustomobject]@{
                  mna=[pscustomobject]@{class=get-class 'System.WorkItem.Activity.ManualActivity';close='Cancelled';ignore='Completed'}; 
                  rwa=[pscustomobject]@{class=get-class 'System.WorkItem.Activity.ReviewActivity';close='Cancelled';ignore='Completed'};
                  srq=[pscustomobject]@{class=get-class 'System.WorkItem.ServiceRequest';close='Closed';search=@('Completed', 'Cancelled', 'Failed');};
                  inc=[pscustomobject]@{class=get-class 'System.WorkItem.Incident';close='Closed';search='Resolved'};
                  };

function title {
    write-host "###############################";
    write-host "#                             #";
    write-host "#    Ticket Cleaning Tool     #";
    write-host "#         Written By          #";
    write-host "#        Arkam Mazrui         #";
    write-host "#     With inspiration from   #";
    write-host "#        Steve Brown          #";
    write-host "#    arkam.mazrui@gmail.com   #";
    write-host "#                             #";
    write-host "###############################";
    write-host "Current CSV: $($current_csv.path)";
    write-host "";
}

function get-option {
    $in = '';
    while ($global:menu_options -inotcontains $in) {
        $in = read-host "Please enter a valid option";
    }
    return $in;
}

function import-ticket-csv {
    $global:browser.ShowDialog();
    try {
        $global:current_csv.obj = Import-Csv $global:browser.FileName;
        $global:current_csv.path = $global:browser.FileName;
    } catch {
        write-host -ForegroundColor Red "Failed to import $path.";
    }
}

function generate-csv($arr_obj) {
    $global:saveFile.ShowDialog();
    $arr_obj | out-file $global:saveFile.FileName;
}

function ret-ticket-class($ticket_id) {
    return $global:classes.($ticket_id.substring(0,3).ToLower());
}

function get-ticket-class($ticket_id) {
    return ret-ticket-class($ticket_id).class;
}

function ticket-too-old($ticket) {
    $span = New-TimeSpan -Start $ticket.'#TimeAdded' -End $global:today;
    return $span.Days -ge $global:max_age_days;
}

function get-old-open-tickets {
    Write-Host -ForegroundColor Gray "Relax, this will most likely take 30 minutes to an hour.";
    [System.Collections.ArrayList]$open_tickets = @();
    foreach ($class_instance in $global:class_instances) {
        $class = ret-ticket-class($class_instance);
        Write-Host -ForegroundColor Yellow "Now working on $class_instance .";
        $tickets = Get-SCSMClassInstance -ComputerName $global:config.server -Class $class.class | ?{$_.Status.DisplayName -ne $class.close -and $_.Status.DisplayName -ne $class.ignore -and (ticket-too-old($_));};
        $open_tickets += $tickets;
    }
    Write-Host -ForegroundColor Yellow "Now retrieving IDs of tickets.";
    $open_tickets = $open_tickets | %{$_.Id};
    generate-csv($open_tickets);
}

function get-resolved-and-completed {
    Write-Host -ForegroundColor Gray "Relax, this will most likely take 30 minutes to an hour.";
    [System.Collections.ArrayList]$completed_tickets = @();
    Write-Host -ForegroundColor Yellow "Now working on Incidents.";
    $resolved_inc = Get-SCSMClassInstance -ComputerName $global:config.server -Class $global:classes.inc.class | ?{$_.Status.DisplayName -eq $global:classes.inc.search};
    Write-Host -ForegroundColor Yellow "Now working on service requests.";
    $completed_srq = Get-SCSMClassInstance -ComputerName $global:config.server -Class $global:classes.srq.class | ?{$global:classes.srq.search -contains $_.Status.DisplayName};
    Write-Host -ForegroundColor Yellow "Now combining lists.";
    $completed_tickets = $resolved_inc + $completed_inc;
    Write-Host -ForegroundColor Yellow "Now retrieving IDs of tickets.";
    $completed_tickets = $completed_tickets | %{$_.Id};
    generate-csv($completed_tickets);
}

function run-csv {
    if ($global:current_csv.path -ne '') {
        $confirm = read-host "Approximately $($global:current_csv.obj.Length) work items will be closed. Proceed? (y/n)";
        if ($confirm.ToLower() -ne 'y') {
            return;
        }
        foreach ($id in $global:current_csv.obj) {
            $class = get-ticket-class($id.Id);
            $class_index = $id.Id.Substring(0,3).ToLower();
            $class_close = $global:classes."$class_index".close;
            if ($class -eq $null) {
                write-host -ForegroundColor Red "Failed to obtain class for $($id.Id).";
            } else {
                $o = Get-SCSMClassInstance -ComputerName $global:config.server -Class $class -Filter "Id -eq $($id.Id)";
                if ($o -eq $null) {
                    Write-Host -ForegroundColor Red "Failed to find $($id.Id) on $($global:config.server).";
                    $failed.Add($id.Id) | Out-Null;
                    continue;
                }
                if ($o.Status.DisplayName -eq $class_close) {
                    Write-Host -ForegroundColor Yellow "$($id.Id) is already closed.";
                } else {
                    $o.Status = $class_close;
                    try {
                        Update-SCSMClassInstance $o;
                        #Write-Host -ForegroundColor Green "Successfully closed $($id.Id).";
                    } catch {
                        Write-Host -ForegroundColor Red "Failed to close $($id.Id).";
                        $failed.Add($id.Id) | Out-Null;
                    }
                }
            }
        }
        Write-Host -ForegroundColor Gray "Completed running current csv.";
        $failed | Out-File "$($global:current_csv.path)-failed.txt";
        $failed = @();
    } else {
        import-ticket-csv;
        run-csv;
    }
}

function proces-option {
    $option = get-option;
    switch ($option) {
        1{import-ticket-csv;break;}
        2{get-old-open-tickets;break;}
        3{get-resolved-and-completed;break;}
        4{run-csv;break;}
        5{break;}
        'q'{exit;}
    }
}

function menu {
    while ($true) {
        cls;
        title;
        for ($i = 0; $i -lt $global:_menu.Count; $i++) {
            write-host "$($global:menu_options[$i]): $($global:_menu[$i])";
        }
        proces-option;
    }
}

menu;