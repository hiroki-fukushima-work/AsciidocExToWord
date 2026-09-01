# Requires Pester 3.x
$scriptPath = Join-Path $PSScriptRoot 'Convert-AsciiDocToWord.ps1'
. $scriptPath -TestMode

function New-TempJson {
    param([string]$Content)
    $path = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllText($path, $Content)
    return $path
}

# ---------------------------------------------------------------------------
# Invoke-JsonTableDirective
# ---------------------------------------------------------------------------

Describe 'Invoke-JsonTableDirective' {

    Context 'Normal - flat array' {
        It 'returns header row plus data rows' {
            $json = '[{"id":"1","name":"alpha","type":"A"},{"id":"2","name":"beta","type":"B"}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success              | Should Be $true
                $r.TableInfo.MaxColumns | Should Be 3
                $r.TableInfo.Rows.Count | Should Be 3
                $r.TableInfo.Rows[0][0].Text     | Should Be 'id'
                $r.TableInfo.Rows[0][0].IsHeader | Should Be $true
                $r.TableInfo.Rows[0][1].Text     | Should Be 'name'
                $r.TableInfo.Rows[1][0].Text     | Should Be '1'
                $r.TableInfo.Rows[1][0].IsHeader | Should Be $false
                $r.TableInfo.Rows[1][1].Text     | Should Be 'alpha'
                $r.TableInfo.Rows[2][0].Text     | Should Be '2'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'converts null to empty string' {
            $json = '[{"id":"1","value":null}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success | Should Be $true
                $r.TableInfo.Rows[1][1].Text | Should Be ''
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'converts boolean to lowercase true/false' {
            $json = '[{"id":"1","active":true,"disabled":false}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success | Should Be $true
                $r.TableInfo.Rows[1][1].Text | Should Be 'true'
                $r.TableInfo.Rows[1][2].Text | Should Be 'false'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }
    }

    Context 'File not found' {
        It 'returns File not found error' {
            $r = Invoke-JsonTableDirective -FilePath './nonexistent_99999.json' -BaseDirectory $PSScriptRoot
            $r.Success   | Should Be $false
            $r.ErrorCode | Should Be 'File not found'
        }
    }

    Context 'Invalid JSON' {
        It 'returns Invalid JSON error for malformed array' {
            $file = New-TempJson '[{invalid json'
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success   | Should Be $false
                $r.ErrorCode | Should Match '^Invalid JSON'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }
    }

    Context 'Nested objects in JSON table' {
        It 'silently excludes nested object property and outputs remaining scalar columns' {
            $json = '[{"id":"1","nested":{"key":"value"}}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success              | Should Be $true
                $r.TableInfo.MaxColumns | Should Be 1
                $r.TableInfo.Rows[0][0].Text | Should Be 'id'
                $r.TableInfo.Rows[1][0].Text | Should Be '1'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'silently excludes nested array property and outputs remaining scalar columns' {
            $json = '[{"id":"1","tags":["a","b"]}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success              | Should Be $true
                $r.TableInfo.MaxColumns | Should Be 1
                $r.TableInfo.Rows[0][0].Text | Should Be 'id'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'returns Unsupported JSON structure when ALL properties are nested (no scalar columns)' {
            $json = '[{"nested":{"key":"value"}}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success   | Should Be $false
                $r.ErrorCode | Should Be 'Unsupported JSON structure'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'includes properties from later elements missing in the first element' {
            $json = '[{"id":"1","name":"alpha"},{"id":"2","name":"beta","extra":"x"}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success              | Should Be $true
                $r.TableInfo.MaxColumns | Should Be 3
                $r.TableInfo.Rows[0][2].Text | Should Be 'extra'
                $r.TableInfo.Rows[1][2].Text | Should Be ''
                $r.TableInfo.Rows[2][2].Text | Should Be 'x'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'returns Unsupported JSON structure when root is plain object not array' {
            $json = '{"id":"1","name":"test"}'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTableDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success   | Should Be $false
                $r.ErrorCode | Should Be 'Unsupported JSON structure'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

    }
}

# ---------------------------------------------------------------------------
# Invoke-JsonTreeDirective
# ---------------------------------------------------------------------------

Describe 'Invoke-JsonTreeDirective' {

    Context 'Normal - nested object' {
        It 'builds correct tree for a nested object' {
            $json = '{"UserList":{"SearchArea":{"UserId":"s","UserName":"s"},"ResultGrid":"s"}}'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTreeDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success  | Should Be $true
                $r.TreeText | Should Match 'UserList'
                $r.TreeText | Should Match 'SearchArea'
                $r.TreeText | Should Match 'UserId'
                $r.TreeText | Should Match 'UserName'
                $r.TreeText | Should Match 'ResultGrid'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'labels array elements by priority key (name)' {
            $json = '[{"name":"item1","value":1},{"name":"item2","value":2}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTreeDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success  | Should Be $true
                $r.TreeText | Should Match 'item1'
                $r.TreeText | Should Match 'item2'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }

        It 'labels array elements with [index] when no priority key exists' {
            $json = '[{"x":1},{"x":2}]'
            $file = New-TempJson $json
            try {
                $r = Invoke-JsonTreeDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success  | Should Be $true
                $r.TreeText | Should Match '\[0\]'
                $r.TreeText | Should Match '\[1\]'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }
    }

    Context 'File not found' {
        It 'returns File not found error' {
            $r = Invoke-JsonTreeDirective -FilePath './nonexistent_99999.json' -BaseDirectory $PSScriptRoot
            $r.Success   | Should Be $false
            $r.ErrorCode | Should Be 'File not found'
        }
    }

    Context 'Invalid JSON' {
        It 'returns Invalid JSON error' {
            $file = New-TempJson '{invalid json here'
            try {
                $r = Invoke-JsonTreeDirective -FilePath $file -BaseDirectory $PSScriptRoot
                $r.Success   | Should Be $false
                $r.ErrorCode | Should Match '^Invalid JSON'
            }
            finally { Remove-Item $file -ErrorAction SilentlyContinue }
        }
    }
}

# ---------------------------------------------------------------------------
# Mixed directives
# ---------------------------------------------------------------------------

Describe 'Mixed directives' {

    It 'table and tree directives work independently' {
        $tableJson = '[{"id":"1","name":"alpha"},{"id":"2","name":"beta"}]'
        $treeJson  = '{"root":{"child1":"v1","child2":"v2"}}'
        $tableFile = New-TempJson $tableJson
        $treeFile  = New-TempJson $treeJson
        try {
            $tr = Invoke-JsonTableDirective -FilePath $tableFile -BaseDirectory $PSScriptRoot
            $gr = Invoke-JsonTreeDirective  -FilePath $treeFile  -BaseDirectory $PSScriptRoot

            $tr.Success              | Should Be $true
            $tr.TableInfo.Rows.Count | Should Be 3

            $gr.Success  | Should Be $true
            $gr.TreeText | Should Match 'root'
            $gr.TreeText | Should Match 'child1'
            $gr.TreeText | Should Match 'child2'
        }
        finally {
            Remove-Item $tableFile -ErrorAction SilentlyContinue
            Remove-Item $treeFile  -ErrorAction SilentlyContinue
        }
    }

    It 'file-not-found error in one directive does not affect the other' {
        $tableJson = '[{"id":"1","name":"ok"}]'
        $tableFile = New-TempJson $tableJson
        try {
            $tr = Invoke-JsonTableDirective -FilePath $tableFile      -BaseDirectory $PSScriptRoot
            $gr = Invoke-JsonTreeDirective  -FilePath './missing.json' -BaseDirectory $PSScriptRoot

            $tr.Success   | Should Be $true
            $gr.Success   | Should Be $false
            $gr.ErrorCode | Should Be 'File not found'
        }
        finally { Remove-Item $tableFile -ErrorAction SilentlyContinue }
    }
}
