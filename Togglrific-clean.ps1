#######################################################
###                                                 ###
###  Filename: Togglrific.ps1                       ###
###  Author:   Craig Boroson                        ###
###  Version:  1.1                                  ###
###  Date:     March 23, 2016                       ###
###  Purpose:  Collect data from the Toggl website  ###
###            related to hours worked for each     ###
###            customer.  Apply mutipliers for      ###
###            off-hours work                       ###
###                                                 ###
#######################################################

# Note: The key below is associated to Tara Boroson's Toggl account
#       It will need to be changed if this Toggl account is removed or disabled.
$username = "<redacted>"
$pass = "api_token"
$pair = "$($username):$($pass)"
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)
$basicAuthValue = "Basic $base64"
$headers = @{ Authorization = $basicAuthValue }
$contentType = "application/json"
$workspace_id = "656157" # Note: this is HA's unique workspace identifier
$rounding_precision = 2 # This is used to round the results in the report to this many decimal places
$AllHours = @()

#Authorization
##############
Invoke-RestMethod -Uri https://www.toggl.com/api/v8/me -Headers $headers -ContentType $contentType

# Populate the customer list
$uriReport = "https://toggl.com/api/v8/workspaces/$workspace_id/clients"
$AllClients = Invoke-RestMethod -Uri $uriReport -Headers $headers -ContentType $contentType
$CustomerList = $AllClients.name | Sort-Object

# Populate the engineer list
$uriReport = "https://toggl.com/api/v8/workspaces/$workspace_id/users"
$AllEngineers = Invoke-RestMethod -Uri $uriReport -Headers $headers -ContentType $contentType
$EngineerList = $AllEngineers.fullname | Sort-Object

# Populate the bundle list
$uriReport = "https://toggl.com/api/v8/workspaces/$workspace_id/projects"
$AllBundles = Invoke-RestMethod -Uri $uriReport -Headers $headers -ContentType $contentType
#$AllBundles | foreach {$_.name -replace '[^\p{L}\p{Nd}/ /]', '-'}
#$BundleList = $AllBundles.name | Sort-Object


# Sort the output when the header is clicked
function datagrid1_OnColumnHeaderMouseClick ( $EventArgs ) {
    $SortProperty = $datagrid1.Columns[$EventArgs.ColumnIndex].HeaderText
    $SortedGridData = $Datagrid1.datasource | Sort-Object -Property $SortProperty

    $array = New-Object System.Collections.ArrayList
    $array.AddRange( $SortedGridData )
    $datagrid1.DataSource = $array
    }

function filter_bundles ( $Customer ) {
    $CustomerID = $AllClients | Where {$_.name -eq $customer}
    $BundleList = $AllBundles | where {$_.cid -eq $CustomerID.id} | Sort-Object
    
    $Obj_bundles.Items.Clear()
    $BundleList | ForEach-Object { [void] $obj_bundles.Items.Add($_.name) }
    $Form.Refresh()
}


function export_data ([Windows.Forms.DataGridView] $grid) {
    if ($grid.RowCount -eq 0) { return } # nothing to do
    $now = get-date -Format M-d-yyyy_h.m.s
    $file = "$env:TEMP\toggl_export_$now.csv" 
    $datagrid1.Rows | select -expand DataBoundItem | export-csv $file -NoType
    Invoke-Item $file
}


function fetch_data ( $Customer ) {

    # Build the URL for submitting to Toggl
    $enddate = $obj_enddate.selectionstart
    $startdate = $obj_startdate.SelectionStart
    $customerID = $($AllClients | where {$_.name -eq $obj_customers.SelectedItems}).id

    # Derive the id's for each selected engineer
    $userText = ""
    if ($obj_engineers.SelectedItems -ne "") {
        $EngineerIds = @()
        foreach ($engineer in $obj_engineers.SelectedItems) {
            $EngineerIds += $($AllEngineers | where {$_.fullname -eq $engineer}).id
            }
        $EngineerIds = $EngineerIds -join ","
        $UserText = "&user_ids=$EngineerIds"
    }

    # Derive the id's for each selected bundle
    $CustomerID = $AllClients | Where {$_.name -eq $customer}
    $BundleList = $AllBundles | where {$_.cid -eq $CustomerID.id} | Sort-Object
    $BundleIDs = $BundleList.id -join ","
    $bundleText = "&project_ids=$BundleIDs"

    if ($obj_bundles.SelectedItems -ne "") {
        $BundleId = @()
        $BundleId = ($AllBundles | where {$_.cid -eq $CustomerID.id -and $_.name -eq $obj_bundles.SelectedItems}).id
        $BundleText = "&project_ids=$BundleId"
    }


    #Reports Request Parameters
    ###########################
    # user_agent: string, required, the name of your application or your email address so we can get in touch in case you're doing something wrong.
    # workspace_id: integer, required. The workspace whose data you want to access.
    # since: string, ISO 8601 date (YYYY-MM-DD), by default until - 6 days.
    # until: string, ISO 8601 date (YYYY-MM-DD), by default today
    # billable: possible values: yes/no/both, default both
    # client_ids: client ids separated by a comma, 0 if you want to filter out time entries without a client
    # project_ids: project ids separated by a comma, 0 if you want to filter out time entries without a project
    # user_ids: user ids separated by a comma
    # tag_ids: tag ids separated by a comma, 0 if you want to filter out time entries without a tag
    # task_ids: task ids separated by a comma, 0 if you want to filter out time entries without a task
    # time_entry_ids: time entry ids separated by a comma
    # description: string, time entry description
    # without_description: true/false, filters out the time entries which do not have a description ('(no description)')
    # order_field:
    # - date/description/duration/user in detailed reports
    # - title/duration/amount in summary reports
    # - title/day1/day2/day3/day4/day5/day6/day7/week_total in weekly report
    # order_desc: on/off, on for descending and off for ascending order
    # distinct_rates: on/off, default off
    # rounding: on/off, default off, rounds time according to workspace settings
    # display_hours: decimal/minutes, display hours with minutes or as a decimal number, default minutes


    #Billable Report
    ##############
    # This pulls all Toggl data for the selected period that are flagged as billable
    $uriReport = "https://toggl.com/reports/api/v2/details?user_agent=api_test&workspace_id=$workspace_id&billable=yes&since=$($startdate.ToString("yyyy-MM-dd"))&until=$($enddate.ToString("yyyy-MM-dd"))$UserText$BundleText"
    $TogglResponse = Invoke-RestMethod -Uri $uriReport -Headers $headers -ContentType $contentType
    $responseTotal = $TogglResponse.total_count
    $pageNum = 1
    $Billable = @()
    while ($responseTotal -gt 0)
    { 
        $TogglResponse = Invoke-RestMethod -Uri $uriReport+"&page="+$pageNum -Headers $headers -ContentType $contentType
        $TogglResponseData = $TogglResponse.data
        $Billable += $TogglResponseData
        $responseTotal = $responseTotal - $TogglResponse.per_page 
        $pageNum++
    }

    #$Billable | foreach {$_.project -replace '[^\p{L}\p{Nd}/ /]', '-'}

    # Filter out unwanted items
    #$Billable = $Billable | where {$_.pid -eq $BundleIDs}

    $Billable | Add-Member -MemberType NoteProperty -Name "adjusted_hours" -Value $null
    $Billable | Add-Member -MemberType NoteProperty -Name "adjustment_reason" -Value ""
    $Billable | Add-Member -MemberType NoteProperty -Name "Questionable" -Value $false

    # Convert all durations from milliseconds to hours
    $billable | foreach {$_.dur = ($_.dur / 1000 / 60 / 60)}

    # Look for questionable entries
    $Billable | where {$_.client -eq $Null} | foreach {$_.Questionable = "Missing client in billable entry"}
    $Billable | where {$_.tags -match "Verify with Sales"} | foreach {$_.Questionable = "Entry tagged as Verify with Sales"}
    $Billable | where {$_.dur -gt 10 -and ($_.tags -join ",") -notmatch "Off-Hours"} | foreach {$_.Questionable = "Possible off-hours work not tagged as such"}
    $Billable | where {($_.end).day -match "0|6" -and ($_.tags -join ",") -notmatch "Off-Hours"} | foreach {$_.Questionable = "Possible off-hours work not tagged as such"}
    $Billable | where {($_.tags -join ",") -match "Pre-sales"} | foreach {$_.Questionable = "Pre-sales item marked as billable"}
    $Billable | where {($_.tags -join ",") -match "Travel"} | foreach {$_.Questionable = "Travel item marked as billable"}
    $Billable | where {($_.tags -join ",") -match "On-Hours" -and ($_.tags -join ",") -match "Off-Hours"} | foreach {$_.Questionable = "Entry tagged as On-Hours and Off-Hours"}
    $Billable | where {($_.tags -join ",") -match "On-Site" -and ($_.tags -join ",") -match "Remote"} | foreach {$_.Questionable = "Entry tagged as On-Site and Remote"}
    $Billable | where {($_.tags -join ",") -match "Off-Hours" -and ($_.tags -join ",") -match "On-Call"} | foreach {$_.Questionable = "Entry tagged as Off-hours and On-call"}
    $Billable | where {($_.tags -join ",") -match "On-Site" -and ($_.tags -join ",") -match "On-Call"} | foreach {$_.Questionable = "Entry tagged as On-Site and On-call"}


    # Note:  Before September 1, 2016, the off-hours totals were mulitplied by 1.5.
    #        After August 31, 2016, the off-hours entries were not modified.  Therfore,
    #        the calculation below takes these differences into account.
    # Mulitply off-hours work by 1.5
    $Billable | where {$_.tags -match "Off-hours" `
                        -and $_.description -notmatch "^OH:"} | foreach {$_.adjusted_hours = $_.dur * 1.5; $_.adjustment_reason = "Off-hours work (*1.5)"}

    # Mulitply on-call hours by 0.5
    $Billable | where {($_.tags -join ",") -match "On-Call"} | foreach {$_.adjusted_hours = $_.dur * 0.5; $_.adjustment_reason = "on-call work (*0.5)"}

    # Iterate through each billable item and adjust time for contractual minimums
    # On-site, work has a 4-hour minimum for Zone 1 unless it's a tagged as Travel
    $Billable | where {$_.tags -match "On-Site" `
                        -and $_.description -notmatch "^OM:" `
                        -and $_.dur -lt 4 `
                        -and $_.adjusted_hours -lt 4 `
                        -and $_.project -notmatch "-Z2" `
                        -and $_.project -notmatch "-Z3" `
                        -and ($_.tags -join ",") -notmatch "Travel"} | foreach {$_.adjusted_hours = 4; $_.adjustment_reason = "on-site minimum for zone 1"}

    # On-site, work has an 8-hour minimum for Zone 2 unless it's a tagged as Travel
    $Billable | where {$_.tags -match "On-Site" `
                        -and $_.description -notmatch "^OM:" `
                        -and $_.dur -lt 8 `
                        -and $_.adjusted_hours -lt 8 `
                        -and $_.project -match "-Z2" `
                        -and $_.project -notmatch "-Z3" `
                        -and ($_.tags -join ",") -notmatch "Travel"} | foreach {$_.adjusted_hours = 8; $_.adjustment_reason = "on-site minimum for zone 2"}

    # Look for adjusted on-site items for the same engineer on the same day
    For ($i=1; $i -le 31; $i++) {
        $a = $Billable | where {$_.adjustment_reason -ne "" `
                        -and $(get-date $_.start).day -eq $i `
                        -and $_.tags -match "On-Site" `
                        -and $_.tags -notmatch "Off-Hours"} | group user | where {$_.count -gt 1} | Select-Object -ExpandProperty group | group project | where {$_.count -gt 1} | Select-Object -ExpandProperty group 
        # All zone 1 combined entries greater than 4 should be left alone
        if (($a.dur | Measure-Object -sum).sum -ge 4 -and $a.project -notmatch "-Z2" -and $_.project -notmatch "-Z3") {             
            foreach ($record in $a) {
                $billable | where {$_.id -eq $record.id} | foreach {$_.adjusted_hours = $record.dur; $_.adjustment_reason = "On-site work in zone 1 split between multiple entries with combined total greater than 4"}
            }
        }

        # All zone 1 combined entries less than 4 should be rounded up to a total of 4 hours
        if (($a.dur | Measure-Object -sum).sum -lt 4 -and $a.project -notmatch "-Z2" -and $_.project -notmatch "-Z3") {             
            foreach ($record in $a) {
                $billable | where {$_.id -eq $record.id} | foreach {$_.adjusted_hours = (4/$a.count) ; $_.adjustment_reason = "On-site work in zone 1 split between multiple entries with combined total less than 4"}
            }
        }

        # All zone 2 combined entries greater than 8 should be left alone
        if (($a.dur | Measure-Object -sum).sum -ge 8 -and $a.project -match "-Z2" ) {             
            foreach ($record in $a) {
                $billable | where {$_.id -eq $record.id} | foreach {$_.adjusted_hours = $record.dur; $_.adjustment_reason = "On-site work in zone 2 split between multiple entries with combined total greater than 8"}
            }
        }

        # All zone 2 combined entries less than 8 should be rounded up to a total of 8 hours
        if (($a.dur | Measure-Object -sum).sum -lt 4 -and $a.project -match "-Z2" ) {             
            foreach ($record in $a) {
                $billable | where {$_.id -eq $record.id} | foreach {$_.adjusted_hours = (8/$a.count) ; $_.adjustment_reason = "On-site work in zone 2 split between multiple entries with combined total less than 8"}
            }
        }

    }

    # Copy the actual hours worked to the adjusted column to replace zeros
    $Billable | where {$_.adjustment_reason -eq ""} | foreach {$_.adjusted_hours = $_.dur}


    # Round up all durations to the nearest 15-minute interval
    $billable | foreach {
        $_.adjusted_hours = [math]::Ceiling($_.adjusted_hours * 4 ) / 4
        }


    $AllHours = @()
    $Billable | foreach {

        $tempobj = New-Object System.Object
        $tempobj | Add-Member -MemberType NoteProperty -Name "Customer" -Value $_.client
        $tempobj | Add-Member -MemberType NoteProperty -Name "Engineer" -Value $_.user
        $tempobj | Add-Member -MemberType NoteProperty -Name "Bundle" -Value $_.project
        $tempobj | Add-Member -MemberType NoteProperty -Name "Description" -Value $_.description
        # Convert from milliseconds to hours and rounds the result to a defined number of decimal places
        $tempobj | Add-Member -MemberType NoteProperty -Name "Hours" -Value $_.dur
        $tempobj | Add-Member -MemberType NoteProperty -Name "AdjustedHours" -Value ([decimal]$_.adjusted_hours)
        $tempobj | Add-Member -MemberType NoteProperty -Name "Adjusted" -Value $_.adjustment_reason
        $tempobj | Add-Member -MemberType NoteProperty -Name "Billable" -Value $_.is_billable
        $tempobj | Add-Member -MemberType NoteProperty -Name "Questionable" -Value $_.Questionable
        $tempobj | Add-Member -MemberType NoteProperty -Name "Tags" -Value $($_.tags -join ";") 
        $tempobj | Add-Member -MemberType NoteProperty -Name "LastUpdated" -Value $(get-date $_.updated)
        $tempobj | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $(get-date $_.start)
        # Toggl sometimes fails to record the end time which causes a script error if isn't trapped
        $ErrorActionPreference = "SilentlyContinue"
        $tempobj | Add-Member -MemberType NoteProperty -Name "EndTime" -Value $(get-date $_.end)

        # Append the current user's summary to a growing array 
        $AllHours += $tempobj
        }


    # Calculate total adjusted billable hours
    $TotalHours = ($AllHours.adjustedhours | Measure-Object -sum).sum
    $label9.Text = "Total Adjusted Hours: $TotalHours"

    # Build array for displaying the output in the grid control
    $array = New-Object System.Collections.ArrayList 
    $array.AddRange($AllHours)
    $dataGrid1.Datasource = $Array
    $Form.Refresh()

}

function GenerateForm { 
    # Build date selector
    #####################
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    Add-Type -AssemblyName System.Windows.Forms
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null 
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null 

    $Form = New-Object system.Windows.Forms.Form
        $Form.Text = "Total Billable Hours per Customer"
        $Form.TopMost = $false
        $Form.Width = 700
        $Form.Height = 860

    $obj_customers = New-Object system.windows.Forms.ListBox
        $obj_customers.Text = "listBox"
        $obj_customers.Width = 323
        $obj_customers.Height = 120
        $obj_customers.location = new-object system.drawing.point(330,40)
        #$obj_customers.SelectionMode = "MultiExtended"
        $obj_customers.Add_mouseclick({filter_bundles $obj_customers.SelectedItem })
        $CustomerList | ForEach-Object { [void] $obj_customers.Items.Add($_) }
        $Form.controls.Add($obj_customers)

    $groupbox2 = New-Object System.Windows.Forms.GroupBox
        $groupbox2.text = "Customer"
        $groupbox2.width = 346
        $groupbox2.Height = 150
        $groupbox2.sendtoback()
        $groupbox2.Location = new-object system.drawing.point(320,14)
        $form.Controls.Add($groupbox2)

    $obj_bundles = New-Object system.windows.Forms.ListBox
        $obj_bundles.Text = "listBox"
        $obj_bundles.Width = 323
        $obj_bundles.Height = 120
        $obj_bundles.location = new-object system.drawing.point(330,198)
        #$obj_bundles.SelectionMode = "MultiExtended"
        If ($BundleList) {$BundleList | ForEach-Object { [void] $obj_bundles.Items.Add($_) }}
        $Form.controls.Add($obj_bundles)

    $groupbox3 = New-Object System.Windows.Forms.GroupBox
        $groupbox3.text = "Bundle(s)"
        $groupbox3.width = 346
        $groupbox3.Height = 150
        $groupbox3.sendtoback()
        $groupbox3.Location = new-object system.drawing.point(320,170)
        $form.Controls.Add($groupbox3)

    $obj_engineers = New-Object system.windows.Forms.ListBox
        $obj_engineers.Text = "listBox"
        $obj_engineers.Width = 323
        $obj_engineers.Height = 120
        $obj_engineers.location = new-object system.drawing.point(330,352)
        $obj_engineers.SelectionMode = "MultiExtended"
        $EngineerList | ForEach-Object { [void] $obj_engineers.Items.Add($_) }
        $Form.controls.Add($obj_engineers)

    $groupbox4 = New-Object System.Windows.Forms.GroupBox
        $groupbox4.text = "Engineer(s)"
        $groupbox4.width = 346
        $groupbox4.Height = 150
        $groupbox4.sendtoback()
        $groupbox4.Location = new-object system.drawing.point(320,330)
        $form.Controls.Add($groupbox4)

    $label7 = New-Object system.windows.Forms.Label
        $label7.Text = "Start Date"
        $label7.AutoSize = $true
        $label7.Width = 25
        $label7.Height = 10
        $label7.location = new-object system.drawing.point(23,14)
        $label7.Font = "Microsoft Sans Serif,10"
        $Form.controls.Add($label7)

    $label8 = New-Object system.windows.Forms.Label
        $label8.Text = "End Date"
        $label8.AutoSize = $true
        $label8.Width = 25
        $label8.Height = 10
        $label8.location = new-object system.drawing.point(23,219)
        $label8.Font = "Microsoft Sans Serif,10"
        $Form.controls.Add($label8)

    $label9 = New-Object system.windows.Forms.Label
        $label9.Text = "Total Adjusted Hours: "
        $label9.AutoSize = $true
        $label9.Width = 95
        $label9.Height = 10
        $label9.location = new-object system.drawing.point(23,772)
        $label9.Font = "Microsoft Sans Serif,10"
        $Form.controls.Add($label9)

    $obj_startdate = New-Object System.Windows.Forms.MonthCalendar 
        $obj_startdate.ShowTodayCircle = $False
        $obj_startdate.location = New-object System.Drawing.Point(23,43)
        $obj_startdate.MaxSelectionCount = 1
        $obj_startdate.setdate($(get-date -day 1))
        $form.Controls.Add($obj_startdate) 

    $obj_enddate = New-Object System.Windows.Forms.MonthCalendar 
        $obj_enddate.ShowTodayCircle = $False
        $obj_enddate.location = New-object System.Drawing.Point(23,246)
        $obj_enddate.MaxSelectionCount = 1
        $form.Controls.Add($obj_enddate) 

    $groupbox1 = New-Object System.Windows.Forms.GroupBox
        $groupbox1.width = 2
        $groupbox1.Height = 440
        $groupbox1.Location = new-object system.drawing.point(290,40)
        $form.Controls.Add($groupbox1)

    $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(23,440)
        $OKButton.Size = New-Object System.Drawing.Size(228,30)
        $OKButton.Text = "GET MY DATA, DAMMIT!"
        #$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $OKButton.FlatStyle = "Standard"
        $OKButton.Add_Click({fetch_data $obj_customers.SelectedItem})
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(610,772)
        $CancelButton.Size = New-Object System.Drawing.Size(60,30)
        $CancelButton.Text = "Quit"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $CancelButton.FlatStyle = "Standard"
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)

    $ExportButton = New-Object System.Windows.Forms.Button
        $ExportButton.Location = New-Object System.Drawing.Point(516,772)
        $ExportButton.Size = New-Object System.Drawing.Size(60,30)
        $ExportButton.Text = "Export"
        $ExportButton.FlatStyle = "Standard"
        $ExportButton.Add_MouseClick( { export_data $datagrid1 } )
        $form.Controls.Add($ExportButton)

    $dataGrid1 = New-Object System.Windows.Forms.DataGridView
        $System_Drawing_Size = New-Object System.Drawing.Size 
        $System_Drawing_Size.Width = 652 
        $System_Drawing_Size.Height = 250 
        $dataGrid1.Size = $System_Drawing_Size 
        $dataGrid1.DataBindings.DefaultDataSourceUpdateMode = 0 
        $dataGrid1.Name = "dataGrid1" 
        $dataGrid1.DataMember = "" 
        $dataGrid1.TabIndex = 0 
        $System_Drawing_Point = New-Object System.Drawing.Point 
        $System_Drawing_Point.X = 18 
        $System_Drawing_Point.Y = 510 
        $dataGrid1.Location = $System_Drawing_Point 
        $dataGrid1.add_ColumnHeaderMouseClick( { datagrid1_OnColumnHeaderMouseClick $_ } )
        $dataGrid1.AllowUserToAddRows = $False
        $dataGrid1.AllowUserToDeleteRows = $False 
        $dataGrid1.ReadOnly = $True 
        $form.Controls.Add($dataGrid1) 


    $result = $Form.ShowDialog()
    #$Form.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {exit}

}
generateform