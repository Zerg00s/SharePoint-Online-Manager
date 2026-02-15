using ClosedXML.Excel;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Service for exporting data to Excel (XLSX) format.
/// </summary>
public class ExcelExporter
{
    /// <summary>
    /// Exports document comparison results to an Excel workbook with overview and sites detail sheets.
    /// </summary>
    public void ExportDocumentCompareReport(DocumentCompareResult result, string filePath)
    {
        using var workbook = new XLWorkbook();

        // Overview Sheet (bird's eye view)
        var overviewSheet = workbook.Worksheets.Add("Overview");
        CreateOverviewSheet(overviewSheet, result);

        // Sites Detail Sheet (detailed per-site breakdown)
        var sitesSheet = workbook.Worksheets.Add("Sites Detail");
        CreateSitesDetailSheet(sitesSheet, result);

        workbook.SaveAs(filePath);
    }

    private static void CreateOverviewSheet(IXLWorksheet sheet, DocumentCompareResult result)
    {
        int row = 1;

        // Title
        sheet.Cell(row, 1).Value = "Document Compare Report - Overview";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 1).Style.Font.FontSize = 16;
        row += 2;

        // Execution Info Section
        sheet.Cell(row, 1).Value = "Execution Details";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
        sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightBlue;
        row++;

        AddSummaryRow(sheet, ref row, "Executed At", result.ExecutedAt.ToString("yyyy-MM-dd HH:mm:ss"));
        AddSummaryRow(sheet, ref row, "Completed At", result.CompletedAt?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");
        AddSummaryRow(sheet, ref row, "Duration", FormatDuration(result.Duration));
        AddSummaryRow(sheet, ref row, "Overall Status", result.Success ? "Success" : "Failed");

        var statusCell = sheet.Cell(row - 1, 2);
        statusCell.Style.Fill.BackgroundColor = result.Success ? XLColor.LightGreen : XLColor.LightCoral;

        if (result.ThrottleRetryCount > 0)
        {
            AddSummaryRow(sheet, ref row, "Throttle Retries", result.ThrottleRetryCount.ToString());
        }
        row++;

        // Site Pair Statistics Section
        sheet.Cell(row, 1).Value = "Site Pair Statistics";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
        sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightBlue;
        row++;

        AddSummaryRow(sheet, ref row, "Total Site Pairs", result.TotalPairsProcessed.ToString());
        AddSummaryRow(sheet, ref row, "Successful Pairs", result.SuccessfulPairs.ToString());
        AddSummaryRow(sheet, ref row, "Failed Pairs", result.FailedPairs.ToString());

        if (result.FailedPairs > 0)
        {
            sheet.Cell(row - 1, 2).Style.Fill.BackgroundColor = XLColor.LightCoral;
        }
        row++;

        // Document Statistics Section
        var (found, sizeIssues, sourceOnly, targetOnly, _) = result.GetSummary();
        var totalSourceDocs = found + sourceOnly;
        var totalTargetDocs = found + targetOnly;

        sheet.Cell(row, 1).Value = "Document Statistics";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
        sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightBlue;
        row++;

        AddSummaryRow(sheet, ref row, "Total Source Documents", totalSourceDocs.ToString());
        AddSummaryRow(sheet, ref row, "Total Target Documents", totalTargetDocs.ToString());
        row++;

        // Found - green
        sheet.Cell(row, 1).Value = "Documents Found";
        sheet.Cell(row, 2).Value = found;
        sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
        row++;

        // Size Issues - yellow (concerning size differences)
        sheet.Cell(row, 1).Value = "Size Issues (0 bytes or <30%)";
        sheet.Cell(row, 2).Value = sizeIssues;
        if (sizeIssues > 0)
        {
            sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightYellow;
        }
        row++;

        // Source Only - red
        sheet.Cell(row, 1).Value = "Source Only (Not Migrated)";
        sheet.Cell(row, 2).Value = sourceOnly;
        if (sourceOnly > 0)
        {
            sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightCoral;
        }
        row++;

        // Target Only - orange
        sheet.Cell(row, 1).Value = "Target Only (Extra on Target)";
        sheet.Cell(row, 2).Value = targetOnly;
        if (targetOnly > 0)
        {
            sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.Apricot;
        }
        row += 2;

        // Size Statistics Section
        sheet.Cell(row, 1).Value = "Size Statistics";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
        sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightBlue;
        row++;

        AddSummaryRow(sheet, ref row, "Total Source Size", FormatFileSize(result.TotalSourceSizeBytes));
        AddSummaryRow(sheet, ref row, "Total Target Size", FormatFileSize(result.TotalTargetSizeBytes));
        row++;

        // Version Statistics Section
        sheet.Cell(row, 1).Value = "Version Statistics";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
        sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightBlue;
        row++;

        AddSummaryRow(sheet, ref row, "Avg Source Versions/Doc", result.OverallAvgSourceVersions.ToString("F2"));
        AddSummaryRow(sheet, ref row, "Avg Target Versions/Doc", result.OverallAvgTargetVersions.ToString("F2"));

        // Highlight version difference if significant
        if (Math.Abs(result.OverallAvgSourceVersions - result.OverallAvgTargetVersions) > 0.5)
        {
            sheet.Cell(row - 1, 2).Style.Fill.BackgroundColor = XLColor.LightYellow;
        }
        row += 2;

        // Migration Completeness
        sheet.Cell(row, 1).Value = "Migration Completeness";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 2).Value = $"{result.MigrationCompletenessPercent:F1}%";
        sheet.Cell(row, 2).Style.Font.Bold = true;
        if (result.MigrationCompletenessPercent >= 99)
        {
            sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
        }
        else if (result.MigrationCompletenessPercent >= 90)
        {
            sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightYellow;
        }
        else
        {
            sheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightCoral;
        }

        // Auto-fit columns
        sheet.Column(1).Width = 35;
        sheet.Column(2).Width = 25;
    }

    private static void CreateSitesDetailSheet(IXLWorksheet sheet, DocumentCompareResult result)
    {
        // Headers
        var headers = new[]
        {
            "Source Site URL",
            "Source Site Title",
            "Target Site URL",
            "Target Site Title",
            "Source Docs",
            "Target Docs",
            "Found",
            "Size Issues",
            "Source Only",
            "Target Only",
            "% Found",
            "% Not Found",
            "% Target Only",
            "Source Size",
            "Target Size",
            "Avg Src Versions",
            "Avg Tgt Versions",
            "Libraries",
            "Status",
            "Error"
        };

        for (int i = 0; i < headers.Length; i++)
        {
            var cell = sheet.Cell(1, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
            cell.Style.Alignment.WrapText = true;
        }

        // Freeze header row
        sheet.SheetView.FreezeRows(1);

        // Data rows
        int row = 2;
        foreach (var site in result.SiteResults)
        {
            int col = 1;
            sheet.Cell(row, col++).Value = site.SourceSiteUrl;
            sheet.Cell(row, col++).Value = site.SourceSiteTitle;
            sheet.Cell(row, col++).Value = site.TargetSiteUrl;
            sheet.Cell(row, col++).Value = site.TargetSiteTitle;
            sheet.Cell(row, col++).Value = site.TotalSourceDocuments;
            sheet.Cell(row, col++).Value = site.TotalTargetDocuments;
            sheet.Cell(row, col++).Value = site.FoundCount;
            sheet.Cell(row, col++).Value = site.SizeIssueCount;
            sheet.Cell(row, col++).Value = site.SourceOnlyCount;
            sheet.Cell(row, col++).Value = site.TargetOnlyCount;

            // Percentages
            var percentFoundCell = sheet.Cell(row, col++);
            percentFoundCell.Value = site.PercentFound / 100;
            percentFoundCell.Style.NumberFormat.Format = "0.0%";

            var percentNotFoundCell = sheet.Cell(row, col++);
            percentNotFoundCell.Value = site.PercentNotFound / 100;
            percentNotFoundCell.Style.NumberFormat.Format = "0.0%";
            if (site.PercentNotFound > 0)
            {
                percentNotFoundCell.Style.Fill.BackgroundColor = XLColor.LightCoral;
            }

            var percentTargetOnlyCell = sheet.Cell(row, col++);
            percentTargetOnlyCell.Value = site.PercentTargetOnly / 100;
            percentTargetOnlyCell.Style.NumberFormat.Format = "0.0%";
            if (site.PercentTargetOnly > 5)
            {
                percentTargetOnlyCell.Style.Fill.BackgroundColor = XLColor.Apricot;
            }

            // Sizes
            sheet.Cell(row, col++).Value = FormatFileSize(site.TotalSourceSizeBytes);
            sheet.Cell(row, col++).Value = FormatFileSize(site.TotalTargetSizeBytes);

            // Version averages
            var avgSrcVersionCell = sheet.Cell(row, col++);
            avgSrcVersionCell.Value = site.AvgSourceVersions;
            avgSrcVersionCell.Style.NumberFormat.Format = "0.00";

            var avgTgtVersionCell = sheet.Cell(row, col++);
            avgTgtVersionCell.Value = site.AvgTargetVersions;
            avgTgtVersionCell.Style.NumberFormat.Format = "0.00";

            // Highlight if versions differ significantly (versions not migrated)
            if (site.AvgSourceVersions > 1.5 && site.AvgTargetVersions < 1.5)
            {
                avgTgtVersionCell.Style.Fill.BackgroundColor = XLColor.LightYellow;
            }

            sheet.Cell(row, col++).Value = site.LibrariesProcessed;
            sheet.Cell(row, col++).Value = site.Success ? "Success" : "Failed";
            sheet.Cell(row, col++).Value = site.ErrorMessage ?? "";

            // Row color coding based on status
            if (!site.Success)
            {
                for (int c = 1; c <= headers.Length; c++)
                {
                    if (sheet.Cell(row, c).Style.Fill.BackgroundColor == XLColor.NoColor)
                    {
                        sheet.Cell(row, c).Style.Fill.BackgroundColor = XLColor.LightCoral;
                    }
                }
            }
            else if (site.PercentFound >= 99)
            {
                // Only color status cell green for fully migrated sites
                sheet.Cell(row, headers.Length - 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
            }

            row++;
        }

        // Add totals row
        row++;
        sheet.Cell(row, 1).Value = "TOTALS";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Row(row).Style.Font.Bold = true;
        sheet.Row(row).Style.Fill.BackgroundColor = XLColor.LightGray;

        var (totalFound, totalSizeIssues, totalSourceOnly, totalTargetOnly, _) = result.GetSummary();
        var grandTotalSource = totalFound + totalSourceOnly;
        var grandTotalTarget = totalFound + totalTargetOnly;

        sheet.Cell(row, 5).Value = grandTotalSource;
        sheet.Cell(row, 6).Value = grandTotalTarget;
        sheet.Cell(row, 7).Value = totalFound;
        sheet.Cell(row, 8).Value = totalSizeIssues;
        sheet.Cell(row, 9).Value = totalSourceOnly;
        sheet.Cell(row, 10).Value = totalTargetOnly;

        var totalPercentFound = grandTotalSource > 0 ? (double)totalFound / grandTotalSource : 1;
        sheet.Cell(row, 11).Value = totalPercentFound;
        sheet.Cell(row, 11).Style.NumberFormat.Format = "0.0%";

        var totalPercentNotFound = grandTotalSource > 0 ? (double)totalSourceOnly / grandTotalSource : 0;
        sheet.Cell(row, 12).Value = totalPercentNotFound;
        sheet.Cell(row, 12).Style.NumberFormat.Format = "0.0%";

        var totalPercentTargetOnly = grandTotalTarget > 0 ? (double)totalTargetOnly / grandTotalTarget : 0;
        sheet.Cell(row, 13).Value = totalPercentTargetOnly;
        sheet.Cell(row, 13).Style.NumberFormat.Format = "0.0%";

        sheet.Cell(row, 14).Value = FormatFileSize(result.TotalSourceSizeBytes);
        sheet.Cell(row, 15).Value = FormatFileSize(result.TotalTargetSizeBytes);
        sheet.Cell(row, 16).Value = result.OverallAvgSourceVersions;
        sheet.Cell(row, 16).Style.NumberFormat.Format = "0.00";
        sheet.Cell(row, 17).Value = result.OverallAvgTargetVersions;
        sheet.Cell(row, 17).Style.NumberFormat.Format = "0.00";
        sheet.Cell(row, 18).Value = result.SiteResults.Sum(s => s.LibrariesProcessed);

        // Auto-fit columns
        sheet.Columns().AdjustToContents();

        // Set minimum widths for URL columns
        sheet.Column(1).Width = Math.Max(sheet.Column(1).Width, 40);
        sheet.Column(3).Width = Math.Max(sheet.Column(3).Width, 40);
    }

    private static void AddSummaryRow(IXLWorksheet sheet, ref int row, string label, string value)
    {
        sheet.Cell(row, 1).Value = label;
        sheet.Cell(row, 2).Value = value;
        row++;
    }

    private static string FormatDuration(TimeSpan duration)
    {
        if (duration.TotalHours >= 1)
        {
            return $"{(int)duration.TotalHours}h {duration.Minutes}m {duration.Seconds}s";
        }
        if (duration.TotalMinutes >= 1)
        {
            return $"{duration.Minutes}m {duration.Seconds}s";
        }
        return $"{duration.Seconds}s";
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB", "TB"];
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len /= 1024;
        }
        return $"{len:F2} {sizes[order]}";
    }
}
