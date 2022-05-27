Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$msForm = [System.Windows.Forms.Form]::new()
$msForm.BackColor = [System.Drawing.Color]::Black
$msForm.AutoSize = $true
$msForm.MaximizeBox = $false

# TODO: Add more difficulty options
$height = 18 # 16 + 2 empty rows
$width = 34 # 32 + 2 empty rows
$mines = 99
$safeCells = (($height-2)*($width-2))-$mines
$global:dug = 0
$cellDim = 28

$msForm.Text = "Minesweeper - $($mines) mines"

# Max hash value to represent a cell
$maxHash = $((($height-1) * $width) + ($width-1))
$topRight = ($width-1)
$bottomLeft = (($height-1) * $width)

# Keeps track of first dig in a game
$global:newGame = 1

# List of to-clear cells in digChain
$cleared = New-Object collections.arrayList
# List of mines placed
$mineHash = New-Object collections.arrayList

############
# incMineBorder
# Increments the number of nearby mines for each cell
# Param: mineHash - hash of mine currently being placed
Function incMineBorder($mineHash)
{
    # list of cells surrounding the mine
    $borderHash = (($mineHash - $width)-1),($mineHash - $width),(($mineHash - $width)+1),($mineHash-1),($mineHash+1),(($mineHash + $width)-1),($mineHash + $width),(($mineHash + $width)+1)

    forEach ($hash in $borderHash)
    {
        if ($msForm.Controls[$hash].tag -eq "X" -or $msForm.Controls[$hash].tag -eq "empty")
        {
            # dont increment mine or empty border cell
            $msForm.Controls[$hash].Text = ""
        }
        else
        {
            $msForm.Controls[$hash].tag++
            switch ($msForm.Controls[$hash].tag)
            {
                1 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::CadetBlue}
                2 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::LimeGreen}
                3 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::Tomato}
                4 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::Violet}
                5 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::Red}
                6 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::Yellow}
                7 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::Cyan}
                8 {$msForm.Controls[$hash].ForeColor = [System.Drawing.Color]::Magenta}
            }
        }
    }
}

############
# placeMines
# Randomly places mines on field
Function placeMines($cellHash)
{
    # Range for Get-Random
    $inputRange = 0..$maxHash
    # Exclude invis border from Get-Random
    $exclude = 0..$topRight
    $exclude += $bottomLeft..$maxHash

    $i = $width
    while ($i -lt $bottomLeft)
    {
        $exclude += ,$i
        $i += $width
    }
    $i = ($width-1) + $width
    while ($i -lt $maxHash)
    {
        $exclude += ,$i
        $i += $width
    }

    # Exclude at and around first dug cell
    $exclude += $cellHash,(($cellHash - $width)-1),($cellHash - $width),(($cellHash - $width)+1),($cellHash-1),([int]$cellHash+1),(([int]$cellHash + $width)-1),([int]$cellHash + $width),(([int]$cellHash + $width)+1)

    # Iterate for number of mines, only increment i when one is placed
    for ($i = 0; $i -lt $mines;)
    {
        $randomRange = $inputRange | Where-Object { $exclude -notcontains $_ }

        $randomHash = Get-Random -InputObject $randomRange

        # exclude this hash from Get-Random
        $exclude += ,$randomHash
		
        $msForm.Controls[$randomHash].tag = "X"
        $mineHash.Add($randomHash)

        # Increment surrounding cells
        incMineBorder($randomHash)

        $i++
    }
}

############
# digChain
# Clears cells not next to any mines that are connected to cell dug, and spreads out
# Param: name - name/hash of initial cell dug
Function digChain($name)
{
    $cell = $msForm.Controls[$name]
    if ($cell.Enabled -eq $true)
    {
        if ($cell.tag -eq "X" -or $cell.tag -eq "empty" -or $cell.BackColor -eq $([System.Drawing.Color]::Black))
        {
            # if mine, empty border cell, or already dug.. do nothing
        }
        else
        {
            $cell.BackColor = [System.Drawing.Color]::Black
            $global:dug++

            if ($global:dug -eq $safeCells)
            {
                gameOver("win")
            }

            if ($cell.tag -gt 0)
            {
                $cell.Text = "$([int]$cell.tag)"
            }
            else
            {
                $col = $name % $width
                $row = [math]::Truncate($name/$width)
                # Begin adding cells to cleared based on position of current cell
                if ($col -gt 1)
                {
                    $cleared.Add($name-1)
                    if ($row -gt 1)
                    {
                        $cleared.Add(($name-$width)-1)
                    }
                }
                if ($col -lt ($width-2))
                {
                    $cleared.Add([int]$name+1)
                    if ($row -gt 1)
                    {
                        $cleared.Add([int]($name-$width)+1)
                    }
                }
                if ($row -gt 1)
                {
                    $cleared.Add($name-$width)
                }
                if ($row -lt ($height-2))
                {
                    $cleared.Add([int]$name+$width)
                    if ($col -gt 1)
                    {
                        $cleared.Add(([int]$name+$width)-1)
                    }
                    if ($col -lt ($width-2))
                    {
                        $cleared.Add(([int]$name+$width)+1)
                    }
                }
                $cell.Enabled = $false
            }
        }
    }
}

############
# dig
# Digs current cell
# Param: cell - clicked cell
Function dig([System.Windows.Forms.Button]$cell)
{
    # If the first dig, place mines
    if ($global:newGame -eq 1)
    {
        placeMines($cell.Name)
        $global:newGame = 0
        $global:startTime = (Get-Date)
    }

    # Only dig unknown cells
    if ($cell.Text -eq "")
    {
        if ($cell.tag -eq "X")
        {
            gameOver("lose")
        }
        else
        {
            $cleared.Add($cell.Name)
            while($cleared.Count -gt 0)
            {
                $name = $cleared[0]
                digChain($name)
                $cleared.Remove($name)
            }
        }
    }
}

############
# gameOver
Function gameOver($reason)
{
    $endTime = (Get-Date)

    # main end game form
    $endForm = [System.Windows.Forms.Form]::new()
    $endForm.BackColor = [System.Drawing.Color]::Black
    $endForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $endForm.MaximizeBox = $false
    $endForm.MinimizeBox = $false
    $endForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent

    if ($reason -eq "lose")
    {
        forEach ($mine in $mineHash)
        {
            # Highlight all mines
            $msForm.Controls[$mine].BackColor = [System.Drawing.Color]::Maroon
            $msForm.Controls[$mine].ForeColor = [System.Drawing.Color]::Snow
            if ($msForm.Controls[$mine].Text -ne "M")
            {
                $msForm.Controls[$mine].Text = "X"
            }
        }
        $endForm.Size = [System.Drawing.Size]::new(235, 85)
        $endForm.Text = "Game Over"
    }
    else
    {
        $endForm.Size = [System.Drawing.Size]::new(235, 125)
        $endForm.Text = "You Win!"

        # Time and record text box
        $timeBox = [System.Windows.Forms.TextBox]::new()
        $timeBox.ReadOnly = $true
        $timeBox.Multiline = $true
        $timeBox.Enabled = $true
        $timeBox.Size = [System.Drawing.Size]::new(150, 30)
        $timeBox.Location = [System.Drawing.Size]::new(5, 45)
        $timeBox.BorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $timeBox.BackColor = [System.Drawing.Color]::Black
        $timeBox.ForeColor = [System.Drawing.Color]::LightCyan

        # Calc time diff
        $tDiff = $endTime - $global:startTime
        $record = [Environment]::GetEnvironmentVariable('MSTime', 'User')

        if ($record -eq $null)
        {
            # New record
            $newRecord = $true
            $timeBox.Text = "Time: $($tDiff)$([Environment]::NewLine)Record: $($tDiff)"
        }
        else
        {
            if (($tDiff - $record) -lt 0)
            {
                # New record
                $newRecord = $true
                $timeBox.Text = "Time: $($tDiff)$([Environment]::NewLine)Record: $($tDiff)"
            }
            else
            {
                $timeBox.Text = "Time: $($tDiff)$([Environment]::NewLine)Record: $($record)"
            }
        }
        $endForm.Controls.Add($timeBox)
    }

    # New game button
    $replayBtn = [System.Windows.Forms.Button]::new()
    $replayBtn.Size = [System.Drawing.Size]::new(100, 30)
    $replayBtn.Location = [System.Drawing.Size]::new(5, 5)
    $replayBtn.BackColor = [System.Drawing.Color]::FromArgb(43, 46, 54)
    $replayBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $replayBtn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Black
    $replayBtn.FlatAppearance.BorderSize = 0
    $replayBtn.Text = "New Game"
    $replayBtn.ForeColor = [System.Drawing.Color]::White
    $replayBtn.add_mouseDown({
        switch($_.Button){
            "Left" {
                $endForm.Close()
                replay
            }
        }
    })

    # Exit button
    $quitBtn = [System.Windows.Forms.Button]::new()
    $quitBtn.Size = [System.Drawing.Size]::new(100, 30)
    $quitBtn.Location = [System.Drawing.Size]::new(110, 5)
    $quitBtn.BackColor = [System.Drawing.Color]::FromArgb(43, 46, 54)
    $quitBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $quitBtn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Black
    $quitBtn.FlatAppearance.BorderSize = 0
    $quitBtn.Text = "Quit"
    $quitBtn.ForeColor = [System.Drawing.Color]::White
    $quitBtn.add_mouseDown({
        switch($_.Button){
            "Left" {
                $endForm.Close()
                $msForm.Close()
            }
        }
    })

    $endForm.Controls.Add($replayBtn)
    $endForm.Controls.Add($quitBtn)
    $endForm.ShowDialog()

    # Done after showdialog to not cause obvious slowdown
    if ($newRecord)
    {
        [Environment]::SetEnvironmentVariable('MSTime', $tDiff, 'User')
    }
}

############
# replay
# Prepares field for new game
Function replay()
{
    $global:newGame = 1
    $global:dug = 0
    $mineHash.Clear()
    $msForm.Controls.Clear()
    drawMinefield
}

############
# flag
# Flags clicked cell
# Param: cell - clicked cell
Function flag([System.Windows.Forms.Button]$cell)
{
    # Cant flag before mines placed
	if ($global:newGame -eq 0)
	{
		# Cant flag already exposed cells
		if ($cell.BackColor -ne [System.Drawing.Color]::Black)
		{
			if ($cell.Text -ne "M")
			{
				$cell.Text = "M"
				# Store original forecolor
				$cell.FlatAppearance.BorderColor = $cell.ForeColor
				# Replace it with white for M text
				$cell.ForeColor = [System.Drawing.Color]::Snow
				$cell.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(43, 46, 54)
			}
			else
			{
				$cell.Text = ""
				$cell.ForeColor = $cell.FlatAppearance.BorderColor
				$cell.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Black
			}
		}
	}
}

############
# drawMinefield
# Draws the initial minefield
Function drawMinefield()
{
    $msForm.Enabled = $false
    $cellSize = [System.Drawing.Size]::new($cellDim, $cellDim)
    $cellHide = [System.Drawing.Size]::new(0, 0)

    $cellYPos = 2
    for ($y = 0; $y -lt $height; $y++)
    {
        $cellXPos = 2
        for ($x = 0; $x -lt $width; $x++)
        {
            $cellBtn = [System.Windows.Forms.Button]::new()

            # Name unique int hashed from x,y
            $cellBtn.Name = $( ($y * $width) + $x)

            if ($x -eq 0 -or $y -eq 0 -or $x -eq ($width-1) -or $y -eq ($height-1))
            {
                # Outer border of invis cells
                $cellBtn.Enabled = $false
                $cellBtn.Size = $cellHide
                $cellBtn.Tag = "empty"
            }
            else
            {
                $cellBtn.Location = [System.Drawing.Size]::new($cellXPos, $cellYPos)
                $cellBtn.BackColor = [System.Drawing.Color]::FromArgb(43, 46, 54)
                $cellBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $cellBtn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Black
                $cellBtn.FlatAppearance.BorderSize = 0
                $cellBtn.Size = $cellSize
                $cellBtn.Tag = 0
                $cellBtn.add_mouseDown({
                    switch($_.Button){
                        "Left" {
                            dig($this)
                        }
                        "Right" {
                            flag($this)
                        }
                    }
                })
                # draw over invis columns
                $cellXPos += $cellDim + 2
            }
            $msForm.Controls.Add($cellBtn)
        }
        if ($y -ne 0)
        {
            # Draw first real row over invis one
            $cellYPos += $cellDim + 2
        }
    }
    $msForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $msForm.Enabled = $true
}

# main
drawMinefield

$msForm.ShowDialog()