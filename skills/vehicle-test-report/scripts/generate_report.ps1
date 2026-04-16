<#
.SYNOPSIS
    整车功能性能测试报告生成器 - 从JSON数据自动生成Word报告

.DESCRIPTION
    读取 report_data.json（或指定数据文件），自动填充到报告模板中，
    生成包含完整5章内容的Word文档(.docx)。

.PARAMETER DataFile
    JSON数据文件路径（默认: config/report_data_template.json）

.PARAMETER OutputFile
    输出Word文件路径（默认: vehicle_test_report.docx）

.EXAMPLE
    .\generate_report.ps1
    .\generate_report.ps1 -DataFile "my_test_data.json" -OutputFile "report.docx"
#>

param(
    [string]$DataFile = (Join-Path $PSScriptRoot "..\config\report_data_template.json"),
    [string]$OutputFile = (Join-Path $PSScriptRoot "..\output\vehicle_test_report.docx")
)

# ============================================================
# 辅助函数
# ============================================================

function Get-JsonValue {
    <# 安全获取JSON值，null返回占位符 #>
    param($Obj, $Key, [string]$Placeholder = "【待填写】")
    if ($null -eq $Obj -or $null -eq $Obj.$Key) { return $Placeholder }
    return $Obj.$Key
}

function Set-CellColor {
    param($Cell, [string]$Color = "D9E2F3")
    $shading = $Cell.Shading
    $shading.BackgroundPatternColor = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic
    # 使用RGB值设置底色
    switch ($Color) {
        "D9E2F3" { $Cell.Shading.BackgroundPatternColor = 14247641 }  # 蓝色表头
        "FFF2CC" { $Cell.Shading.BackgroundPatternColor = 13421619 }  # 黄色要求行
        "E2EFDA" { $Cell.Shading.BackgroundPatternColor = 14408658 }  # 绿色通过
        "FCE4EC" { $Cell.Shading.BackgroundPatternColor = 15655996 }  # 红色失败
    }
}

function Add-TableFromData {
    <#
    .SYNOPSIS
        从ABS测试数据生成多车速、三次测试的表格

    .PARAMETER WordDoc
        Word文档对象
    .PARAMETER TestData
        测试数据对象（包含test_conditions, test_items, subjective_evaluation）
    .PARAMETER SectionTitle
        章节标题（如"5.1.1 干沥青直线制动"）
    .PARAMETER IsCurve
        是否为弯道测试（显示弯道半径）
    #>
    param(
        $WordDoc,
        $TestData,
        [string]$SectionTitle,
        [bool]$IsCurve = $false
    )

    $selection = $WordDoc.ActiveWindow.Selection

    # 测试条件
    $conditions = $TestData.test_conditions
    if ($null -ne $conditions) {
        $selection.TypeText("测试条件：")
        $selection.TypeParagraph()
        $selection.TypeText("  • 测试路面：$(Get-JsonValue $conditions '测试路面')")
        $selection.TypeParagraph()
        if ($IsCurve) {
            $selection.TypeText("  • 弯道半径：$(Get-JsonValue $conditions '弯道半径')m")
            $selection.TypeParagraph()
        }
        $selection.TypeText("  • 路面附着系数：$(Get-JsonValue $conditions '路面附着系数')")
        $selection.TypeParagraph()
        $selection.TypeText("  • 测试环境：温度$(Get-JsonValue $conditions '测试温度')℃，湿度$(Get-JsonValue $conditions '测试湿度')%")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
    }

    # 测试结果记录表
    $selection.TypeText("测试结果记录表：")
    $selection.TypeParagraph()

    $testItems = $TestData.test_items
    if ($null -eq $testItems -or $testItems.Count -eq 0) {
        $selection.TypeText("【无测试数据】")
        $selection.TypeParagraph()
        return
    }

    # 计算表格行数：表头1行 + 每个车速(3次测试+1平均) + 要求值1行
    $totalRows = 1
    foreach ($item in $testItems) {
        $runCount = 0
        if ($null -ne $item.test_runs) { $runCount = $item.test_runs.Count }
        $totalRows += $runCount + 1  # 测试次数 + 平均行
    }
    $totalRows += 1  # 要求值行

    # 列数：测试项目(车速) + 序号 + 7个指标 + 结论 = 10列
    $colCount = 10

    # 创建表格
    $range = $selection.Range
    $table = $WordDoc.Tables.Add($range, $totalRows, $colCount)
    $table.Style = "网格型"
    $table.Borders.Enable = $true

    # === 表头 ===
    $headers = @("测试项目`n(车速)", "测试`n序号", "平均减速度`n(m/s²)", "制动距离`n(m)",
                 "减速度相邻`n峰谷差值(m/s²)", "转向修正角`n(deg)", "车轮抱死`n时间(s)",
                 "附着系数`n利用率(%)", "主观`n评分", "结论")

    for ($i = 0; $i -lt $colCount; $i++) {
        $cell = $table.Cell(1, $i + 1)
        $cell.Range.Text = $headers[$i]
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Size = 9
        $cell.Range.ParagraphFormat.Alignment = 1  # 居中
        Set-CellColor $cell "D9E2F3"
    }

    # === 数据行 ===
    $rowIdx = 2
    foreach ($item in $testItems) {
        $speed = Get-JsonValue $item "车速" "—"
        $testRuns = $item.test_runs
        $average = $item.average

        if ($null -ne $testRuns) {
            foreach ($run in $testRuns) {
                $seqNo = Get-JsonValue $run "序号" "—"
                $isFirst = ($seqNo -eq 1 -or $seqNo -eq "1")

                $table.Cell($rowIdx, 1).Range.Text = if ($isFirst) { "$speed km/h" } else { "" }
                $table.Cell($rowIdx, 2).Range.Text = "$seqNo"
                $table.Cell($rowIdx, 3).Range.Text = "$(Get-JsonValue $run '平均减速度')"
                $table.Cell($rowIdx, 4).Range.Text = "$(Get-JsonValue $run '制动距离')"
                $table.Cell($rowIdx, 5).Range.Text = "$(Get-JsonValue $run '减速度峰谷差值')"
                $table.Cell($rowIdx, 6).Range.Text = "$(Get-JsonValue $run '转向修正角')"
                $table.Cell($rowIdx, 7).Range.Text = "$(Get-JsonValue $run '车轮抱死时间')"
                $table.Cell($rowIdx, 8).Range.Text = "$(Get-JsonValue $run '附着系数利用率')"
                $table.Cell($rowIdx, 9).Range.Text = "$(Get-JsonValue $run '主观评分')"
                $table.Cell($rowIdx, 10).Range.Text = "$(Get-JsonValue $run '结论')"

                # 居中对齐
                for ($c = 1; $c -le $colCount; $c++) {
                    $table.Cell($rowIdx, $c).Range.ParagraphFormat.Alignment = 1
                    $table.Cell($rowIdx, $c).Range.Font.Size = 9
                }

                $rowIdx++
            }
        }

        # 平均值行
        if ($null -ne $average) {
            $table.Cell($rowIdx, 1).Range.Text = ""
            $table.Cell($rowIdx, 2).Range.Text = "平均"
            $table.Cell($rowIdx, 3).Range.Text = "$(Get-JsonValue $average '平均减速度')"
            $table.Cell($rowIdx, 4).Range.Text = "$(Get-JsonValue $average '制动距离')"
            $table.Cell($rowIdx, 5).Range.Text = "$(Get-JsonValue $average '减速度峰谷差值')"
            $table.Cell($rowIdx, 6).Range.Text = "$(Get-JsonValue $average '转向修正角')"
            $table.Cell($rowIdx, 7).Range.Text = "$(Get-JsonValue $average '车轮抱死时间')"
            $table.Cell($rowIdx, 8).Range.Text = "$(Get-JsonValue $average '附着系数利用率')"
            $table.Cell($rowIdx, 9).Range.Text = "$(Get-JsonValue $average '主观评分')"
            $table.Cell($rowIdx, 10).Range.Text = "$(Get-JsonValue $average '结论')"

            # 平均值行加粗
            for ($c = 1; $c -le $colCount; $c++) {
                $table.Cell($rowIdx, $c).Range.Font.Bold = $true
                $table.Cell($rowIdx, $c).Range.Font.Size = 9
                $table.Cell($rowIdx, $c).Range.ParagraphFormat.Alignment = 1
            }

            $rowIdx++
        }
    }

    # === 要求值行 ===
    $req = $testItems[0].requirements
    if ($null -ne $req) {
        $table.Cell($rowIdx, 1).Range.Text = "要求值"
        $table.Cell($rowIdx, 2).Range.Text = "-"
        $table.Cell($rowIdx, 3).Range.Text = "$(Get-JsonValue $req '平均减速度')"
        $table.Cell($rowIdx, 4).Range.Text = "$(Get-JsonValue $req '制动距离')"
        $table.Cell($rowIdx, 5).Range.Text = "$(Get-JsonValue $req '减速度峰谷差值')"
        $table.Cell($rowIdx, 6).Range.Text = "$(Get-JsonValue $req '转向修正角')"
        $table.Cell($rowIdx, 7).Range.Text = "$(Get-JsonValue $req '车轮抱死时间')"
        $table.Cell($rowIdx, 8).Range.Text = "$(Get-JsonValue $req '附着系数利用率')"
        $table.Cell($rowIdx, 9).Range.Text = "$(Get-JsonValue $req '主观评分')"
        $table.Cell($rowIdx, 10).Range.Text = "通过"

        for ($c = 1; $c -le $colCount; $c++) {
            Set-CellColor $table.Cell($rowIdx, $c) "FFF2CC"
            $table.Cell($rowIdx, $c).Range.Font.Italic = $true
            $table.Cell($rowIdx, $c).Range.Font.Size = 9
            $table.Cell($rowIdx, $c).Range.ParagraphFormat.Alignment = 1
        }
    }

    # 移动到表格下方
    $table.Select()
    $selection.MoveDown(5, 1)  # wdLine, 1行

    # 主观评价
    $subjective = Get-JsonValue $TestData "subjective_evaluation"
    $selection.TypeParagraph()
    $selection.TypeText("主观评价：$subjective")
    $selection.TypeParagraph()
}

# ============================================================
# 主流程
# ============================================================

# 解析数据文件路径
$DataFile = [System.IO.Path]::GetFullPath($DataFile)
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

if (-not (Test-Path $DataFile)) {
    Write-Error "数据文件不存在: $DataFile"
    exit 1
}

# 确保输出目录存在
$outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 读取JSON数据
Write-Host "读取数据文件: $DataFile"
$jsonText = Get-Content $DataFile -Raw -Encoding UTF8
$data = $jsonText | ConvertFrom-Json

# 启动Word应用
Write-Host "启动Word应用..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

try {
    $doc = $word.Documents.Add()
    $selection = $word.ActiveWindow.Selection

    # ============================================================
    # 封面
    # ============================================================
    $selection.Font.Size = 22
    $selection.Font.Bold = $true
    $selection.ParagraphFormat.Alignment = 1  # 居中
    $selection.TypeText("整车功能性能测试报告")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Font.Size = 14
    $selection.Font.Bold = $false
    $selection.TypeText("项目名称：$(Get-JsonValue $data 'project_name')")
    $selection.TypeParagraph()
    $selection.TypeText("报告编号：$(Get-JsonValue $data 'report_id')")
    $selection.TypeParagraph()
    $selection.TypeText("报告日期：$(Get-JsonValue $data 'report_date')")
    $selection.TypeParagraph()

    # 分页
    $selection.InsertNewPage()

    # ============================================================
    # 第一章 前言
    # ============================================================
    $selection.Font.Size = 16
    $selection.Font.Bold = $true
    $selection.ParagraphFormat.Alignment = 0  # 左对齐
    $selection.TypeText("第一章 前言")
    $selection.TypeParagraph()

    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    $foreword = Get-JsonValue $data 'foreword'
    $selection.TypeText($foreword)
    $selection.TypeParagraph()

    $selection.InsertNewPage()

    # ============================================================
    # 第二章 概述
    # ============================================================
    $selection.Font.Size = 16
    $selection.Font.Bold = $true
    $selection.TypeText("第二章 概述")
    $selection.TypeParagraph()

    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("2.1 测试对象")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    $selection.TypeText("$(Get-JsonValue $data 'test_object')")
    $selection.TypeParagraph()

    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("2.2 测试目标")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    if ($null -ne $data.test_objectives) {
        foreach ($obj in $data.test_objectives) {
            $selection.TypeText("  • $obj")
            $selection.TypeParagraph()
        }
    }

    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("2.3 测试周期")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    $selection.TypeText("$(Get-JsonValue $data 'test_period')")
    $selection.TypeParagraph()

    $selection.InsertNewPage()

    # ============================================================
    # 第三章 测试版本说明
    # ============================================================
    $selection.Font.Size = 16
    $selection.Font.Bold = $true
    $selection.TypeText("第三章 测试版本说明")
    $selection.TypeParagraph()

    # 3.1 版本信息
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("3.1 测试版本信息")
    $selection.TypeParagraph()

    $verTable = $doc.Tables.Add($selection.Range, 5, 2)
    $verTable.Style = "网格型"
    $verData = @(
        @("项目", "版本信息"),
        @("软件版本", "$(Get-JsonValue $data 'software_version')"),
        @("硬件版本", "$(Get-JsonValue $data 'hardware_version')"),
        @("固件版本", "$(Get-JsonValue $data 'firmware_version')"),
        @("版本变更说明", "$(Get-JsonValue $data 'version_change')")
    )
    for ($r = 0; $r -lt 5; $r++) {
        for ($c = 0; $c -lt 2; $c++) {
            $verTable.Cell($r + 1, $c + 1).Range.Text = $verData[$r][$c]
            $verTable.Cell($r + 1, $c + 1).Range.Font.Size = 10
        }
        if ($r -eq 0) {
            Set-CellColor $verTable.Cell($r + 1, 1) "D9E2F3"
            Set-CellColor $verTable.Cell($r + 1, 2) "D9E2F3"
            $verTable.Cell($r + 1, 1).Range.Font.Bold = $true
            $verTable.Cell($r + 1, 2).Range.Font.Bold = $true
        }
    }
    $verTable.Select()
    $selection.MoveDown(5, 1)
    $selection.TypeParagraph()

    # 3.2 测试环境
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("3.2 测试环境描述")
    $selection.TypeParagraph()

    $envTable = $doc.Tables.Add($selection.Range, 5, 2)
    $envTable.Style = "网格型"
    $envData = @(
        @("项目", "描述"),
        @("测试场地", "$(Get-JsonValue $data 'test_site')"),
        @("测试设备", "$(Get-JsonValue $data 'test_equipment')"),
        @("测试工具", "$(Get-JsonValue $data 'test_tools')"),
        @("环境条件", "$(Get-JsonValue $data 'environment')")
    )
    for ($r = 0; $r -lt 5; $r++) {
        for ($c = 0; $c -lt 2; $c++) {
            $envTable.Cell($r + 1, $c + 1).Range.Text = $envData[$r][$c]
            $envTable.Cell($r + 1, $c + 1).Range.Font.Size = 10
        }
        if ($r -eq 0) {
            Set-CellColor $envTable.Cell($r + 1, 1) "D9E2F3"
            Set-CellColor $envTable.Cell($r + 1, 2) "D9E2F3"
            $envTable.Cell($r + 1, 1).Range.Font.Bold = $true
            $envTable.Cell($r + 1, 2).Range.Font.Bold = $true
        }
    }
    $envTable.Select()
    $selection.MoveDown(5, 1)
    $selection.TypeParagraph()

    # 3.3 引用的测试设计
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("3.3 引用的测试设计")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    if ($null -ne $data.test_design_refs) {
        foreach ($ref in $data.test_design_refs) {
            $selection.TypeText("  • $ref")
            $selection.TypeParagraph()
        }
    }

    # 3.4 测试通过标准
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("3.4 测试通过标准")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    if ($null -ne $data.pass_criteria) {
        foreach ($criteria in $data.pass_criteria) {
            $selection.TypeText("  • $criteria")
            $selection.TypeParagraph()
        }
    }

    $selection.InsertNewPage()

    # ============================================================
    # 第四章 概要测试结论
    # ============================================================
    $selection.Font.Size = 16
    $selection.Font.Bold = $true
    $selection.TypeText("第四章 概要测试结论")
    $selection.TypeParagraph()

    # 4.1 测试结论总结
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("4.1 测试结论总结")
    $selection.TypeParagraph()

    $sumTable = $doc.Tables.Add($selection.Range, 5, 2)
    $sumTable.Style = "网格型"
    $sumData = @(
        @("统计项", "数值"),
        @("测试用例总数", "$(Get-JsonValue $data 'total_cases')"),
        @("通过用例数", "$(Get-JsonValue $data 'passed_cases')"),
        @("失败用例数", "$(Get-JsonValue $data 'failed_cases')"),
        @("测试结论", "$(Get-JsonValue $data 'conclusion')")
    )
    for ($r = 0; $r -lt 5; $r++) {
        for ($c = 0; $c -lt 2; $c++) {
            $sumTable.Cell($r + 1, $c + 1).Range.Text = $sumData[$r][$c]
            $sumTable.Cell($r + 1, $c + 1).Range.Font.Size = 10
        }
        if ($r -eq 0) {
            Set-CellColor $sumTable.Cell($r + 1, 1) "D9E2F3"
            Set-CellColor $sumTable.Cell($r + 1, 2) "D9E2F3"
            $sumTable.Cell($r + 1, 1).Range.Font.Bold = $true
            $sumTable.Cell($r + 1, 2).Range.Font.Bold = $true
        }
    }
    $sumTable.Select()
    $selection.MoveDown(5, 1)
    $selection.TypeParagraph()

    # 4.2 关键风险
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("4.2 关键风险和规避措施")
    $selection.TypeParagraph()

    if ($null -ne $data.risks -and $data.risks.Count -gt 0) {
        $riskTable = $doc.Tables.Add($selection.Range, $data.risks.Count + 1, 4)
        $riskTable.Style = "网格型"

        # 表头
        $riskHeaders = @("风险描述", "风险等级", "规避措施", "状态")
        for ($i = 0; $i -lt 4; $i++) {
            $riskTable.Cell(1, $i + 1).Range.Text = $riskHeaders[$i]
            $riskTable.Cell(1, $i + 1).Range.Font.Bold = $true
            $riskTable.Cell(1, $i + 1).Range.Font.Size = 10
            Set-CellColor $riskTable.Cell(1, $i + 1) "D9E2F3"
        }

        for ($i = 0; $i -lt $data.risks.Count; $i++) {
            $risk = $data.risks[$i]
            $riskTable.Cell($i + 2, 1).Range.Text = "$(Get-JsonValue $risk 'description')"
            $riskTable.Cell($i + 2, 2).Range.Text = "$(Get-JsonValue $risk 'level')"
            $riskTable.Cell($i + 2, 3).Range.Text = "$(Get-JsonValue $risk 'mitigation')"
            $riskTable.Cell($i + 2, 4).Range.Text = "$(Get-JsonValue $risk 'status')"
            for ($c = 1; $c -le 4; $c++) {
                $riskTable.Cell($i + 2, $c).Range.Font.Size = 10
            }
        }
        $riskTable.Select()
        $selection.MoveDown(5, 1)
        $selection.TypeParagraph()
    }

    $selection.InsertNewPage()

    # ============================================================
    # 第五章 测试项目及结果
    # ============================================================
    $selection.Font.Size = 16
    $selection.Font.Bold = $true
    $selection.TypeText("第五章 测试项目及结果")
    $selection.TypeParagraph()

    # 5.1 ABS测试结果
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("5.1 ABS测试结果")
    $selection.TypeParagraph()

    # 5.1.1 干沥青直线制动
    $selection.Font.Size = 13
    $selection.Font.Bold = $true
    $selection.TypeText("5.1.1 干沥青直线制动")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    if ($null -ne $data.abs_straight_braking) {
        Add-TableFromData -WordDoc $doc -TestData $data.abs_straight_braking -SectionTitle "5.1.1" -IsCurve $false
    } else {
        $selection.TypeText("【无数据】")
        $selection.TypeParagraph()
    }

    # 5.1.2 干沥青弯道制动
    $selection.Font.Size = 13
    $selection.Font.Bold = $true
    $selection.TypeText("5.1.2 干沥青弯道制动")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    if ($null -ne $data.abs_curve_braking) {
        Add-TableFromData -WordDoc $doc -TestData $data.abs_curve_braking -SectionTitle "5.1.2" -IsCurve $true
    } else {
        $selection.TypeText("【无数据】")
        $selection.TypeParagraph()
    }

    # 5.1.3 湿沥青直线制动
    $selection.Font.Size = 13
    $selection.Font.Bold = $true
    $selection.TypeText("5.1.3 湿沥青直线制动")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    if ($null -ne $data.abs_wet_straight_braking) {
        Add-TableFromData -WordDoc $doc -TestData $data.abs_wet_straight_braking -SectionTitle "5.1.3" -IsCurve $false
    } else {
        $selection.TypeText("【无数据】")
        $selection.TypeParagraph()
    }

    # 5.1.4 对开路面制动
    $selection.Font.Size = 13
    $selection.Font.Bold = $true
    $selection.TypeText("5.1.4 对开路面制动")
    $selection.TypeParagraph()
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    if ($null -ne $data.abs_split_braking) {
        Add-TableFromData -WordDoc $doc -TestData $data.abs_split_braking -SectionTitle "5.1.4" -IsCurve $false
    } else {
        $selection.TypeText("【无数据】")
        $selection.TypeParagraph()
    }

    # 5.2 TCS测试结果
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.TypeText("5.2 TCS测试结果")
    $selection.TypeParagraph()

    $selection.Font.Size = 12
    $selection.Font.Bold = $false

    if ($null -ne $data.tcs_results -and $data.tcs_results.Count -gt 0) {
        $tcsTable = $doc.Tables.Add($selection.Range, $data.tcs_results.Count + 1, 5)
        $tcsTable.Style = "网格型"

        $tcsHeaders = @("用例编号", "测试项目", "预期结果", "实际结果", "结论")
        for ($i = 0; $i -lt 5; $i++) {
            $tcsTable.Cell(1, $i + 1).Range.Text = $tcsHeaders[$i]
            $tcsTable.Cell(1, $i + 1).Range.Font.Bold = $true
            $tcsTable.Cell(1, $i + 1).Range.Font.Size = 10
            Set-CellColor $tcsTable.Cell(1, $i + 1) "D9E2F3"
        }

        for ($i = 0; $i -lt $data.tcs_results.Count; $i++) {
            $tcs = $data.tcs_results[$i]
            $tcsTable.Cell($i + 2, 1).Range.Text = "$(Get-JsonValue $tcs 'case_id')"
            $tcsTable.Cell($i + 2, 2).Range.Text = "$(Get-JsonValue $tcs 'test_item')"
            $tcsTable.Cell($i + 2, 3).Range.Text = "$(Get-JsonValue $tcs 'expected')"
            $tcsTable.Cell($i + 2, 4).Range.Text = "$(Get-JsonValue $tcs 'actual')"
            $tcsTable.Cell($i + 2, 5).Range.Text = "$(Get-JsonValue $tcs 'conclusion')"
            for ($c = 1; $c -le 5; $c++) {
                $tcsTable.Cell($i + 2, $c).Range.Font.Size = 10
            }
        }
        $tcsTable.Select()
        $selection.MoveDown(5, 1)
        $selection.TypeParagraph()
    } else {
        $selection.TypeText("【无TCS测试数据】")
        $selection.TypeParagraph()
    }

    # ============================================================
    # 保存文档
    # ============================================================
    Write-Host "保存报告: $OutputFile"
    $doc.SaveAs([ref]$OutputFile, [ref]16)  # wdFormatDocumentDefault = 16 (.docx)
    $doc.Close()
    Write-Host "报告生成完成！"

} catch {
    Write-Error "生成报告时出错: $_"
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
