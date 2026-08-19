using OpenXMLOffice.Presentation_2007;
using OpenXMLOffice.Global_2007;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using Pen = System.Windows.Media.Pen;

namespace SchedulerApp.Services
{
    public class ExportSchedule
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime CreatedDate { get; set; }
        public Schedule SourceSchedule { get; set; }
        public ExportSettings Settings { get; set; } = new ExportSettings();
        public string OutputPath { get; set; }
        public bool IsReady { get; set; } = true;
        [Newtonsoft.Json.JsonIgnore]
        public List<BitmapImage> SlidePreviews { get; set; } = new List<BitmapImage>();
        public int CurrentPreviewSlide { get; set; } = 0;
    }

    public class ExportSettings
    {
        public double FontSize { get; set; } = 35;
        public double CellPadding { get; set; } = 8;
        public double CellMargin { get; set; } = 3;
        public bool ShowGridLines { get; set; } = true;
        public bool ShowCellBackground { get; set; } = true;
        public string HeaderColor { get; set; } = "#BFDBFE";
        public string CellColor { get; set; } = "#E0F2FE";
        public string TextColor { get; set; } = "#1F2937";
        public string NameCellColor { get; set; } = "#E0F2FE";
        public string TimeCellColor { get; set; } = "#DBEAFE";
        public string DaysRowColor { get; set; } = "#BFDBFE";
        public string TemplateName { get; set; } = "BlackWhite";
        public string BackgroundImagePath { get; set; } = "Icons/Black-White-Template.png";
        public double HeaderFontSize { get; set; } = 40;
        public double TimeColumnWidth { get; set; } = 120;
        public double DayRowHeight { get; set; } = 60;
        public double CellHeight { get; set; } = 80;
        public bool UseColorCoding { get; set; } = true;
        public bool IncludeWeekends { get; set; } = true;
        public int SlidesPerWeek { get; set; } = 1;
        public double CellSpacing { get; set; } = 2;
        public double CellBorderRadius { get; set; } = 4;
        public double CellWidthScale { get; set; } = 0.7;
        public double CellHeightScale { get; set; } = 0.6;
        public string BackgroundColor { get; set; } = "#FFFFFF";
        public string TitleText { get; set; } = "Weekly Schedule";
        public string TitleColor { get; set; } = "#1F2937";
    }

    public class TemplateInfo
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string ColorHex { get; set; }
        public string Icon { get; set; }
        public string BackgroundImagePath { get; set; }
        public string TemplateFilePath { get; set; }
        public bool IsCustom { get; set; }
    }

    public class ExportResult
    {
        public bool Success { get; set; }
        public string PptxPath { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class ExportService
    {
        private static List<ExportSchedule> _exportSchedules = new List<ExportSchedule>();
        private static List<TemplateInfo> _templates = new List<TemplateInfo>();
        private static Dispatcher _dispatcher;
        private static StateSave _stateSave = new StateSave();

        static ExportService()
        {
            InitializeTemplates();
            LoadExportsFromDisk();
        }

        private static void LoadExportsFromDisk()
        {
            try
            {
                var saved = _stateSave.LoadExports();
                if (saved != null)
                {
                    _exportSchedules.Clear();
                    _exportSchedules.AddRange(saved);
                }
            }
            catch { }
        }

        private static void SaveExportsToDisk()
        {
            try
            {
                _stateSave.SaveAllExports(_exportSchedules);
            }
            catch { }
        }

        public static void SetDispatcher(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        private static void InitializeTemplates()
        {
            _templates = new List<TemplateInfo>
            {
                new TemplateInfo
                {
                    Name = "BlackWhite",
                    DisplayName = "Black & White",
                    Description = "Monochrome classic look",
                    ColorHex = "#111827",
                    Icon = "Icons/Black-White-Template.png",
                    BackgroundImagePath = "Icons/Black-White-Template.png",
                    TemplateFilePath = "Templates/Black-White-Template.pptx"
                },
                new TemplateInfo
                {
                    Name = "BlueDoodle",
                    DisplayName = "Blue Doodle",
                    Description = "Playful blue hand-drawn design",
                    ColorHex = "#3B82F6",
                    Icon = "Icons/Blue-Doodle-Template.png",
                    BackgroundImagePath = "Icons/Blue-Doodle-Template.png",
                    TemplateFilePath = "Templates/Blue-Doodle-Template.pptx"
                },
                new TemplateInfo
                {
                    Name = "BlueDots",
                    DisplayName = "Blue Dots",
                    Description = "Polka dot blue pattern",
                    ColorHex = "#60A5FA",
                    Icon = "Icons/Blue-Dots-Template.png",
                    BackgroundImagePath = "Icons/Blue-Dots-Template.png",
                    TemplateFilePath = "Templates/Blue-Dots-Template.pptx"
                },
                new TemplateInfo
                {
                    Name = "BlueWatercolor",
                    DisplayName = "Blue Watercolor",
                    Description = "Soft watercolor wash",
                    ColorHex = "#93C5FD",
                    Icon = "Icons/Blue-Watercolor-Template.png",
                    BackgroundImagePath = "Icons/Blue-Watercolor-Template.png",
                    TemplateFilePath = "Templates/Blue-Watercolor-Template.pptx"
                },
                new TemplateInfo
                {
                    Name = "CreamDarkBrown",
                    DisplayName = "Cream & Brown",
                    Description = "Warm cream and dark brown tones",
                    ColorHex = "#D97706",
                    Icon = "Icons/Cream-Dark-Brown-Template.png",
                    BackgroundImagePath = "Icons/Cream-Dark-Brown-Template.png",
                    TemplateFilePath = "Templates/Cream-Dark-Brown-Template.pptx"
                },
                new TemplateInfo
                {
                    Name = "GreyBlack",
                    DisplayName = "Grey & Black",
                    Description = "Sleek grey and black design",
                    ColorHex = "#374151",
                    Icon = "Icons/Grey-Black-Template.png",
                    BackgroundImagePath = "Icons/Grey-Black-Template.png",
                    TemplateFilePath = "Templates/Grey-Black-Template.pptx"
                },
                new TemplateInfo
                {
                    Name = "PeachGreen",
                    DisplayName = "Peach & Green",
                    Description = "Fresh peach and green tones",
                    ColorHex = "#F97316",
                    Icon = "Icons/Peach-Green-Template.png",
                    BackgroundImagePath = "Icons/Peach-Green-Template.png",
                    TemplateFilePath = "Templates/Peach-Green-Template.pptx"
                }
            };
        }

        public static List<ExportSchedule> GetExportSchedules()
        {
            return _exportSchedules;
        }

        public static ExportSchedule GetExportScheduleById(string id)
        {
            return _exportSchedules.FirstOrDefault(s => s.Id == id);
        }

        public static ExportSchedule CreateExportItem(Schedule schedule, string customName = null)
        {
            var exportSchedule = new ExportSchedule
            {
                Id = Guid.NewGuid().ToString(),
                Name = customName ?? $"{schedule.Name}_Export_{DateTime.Now:yyyyMMdd_HHmmss}",
                CreatedDate = DateTime.Now,
                SourceSchedule = schedule,
                Settings = new ExportSettings
                {
                    IncludeWeekends = schedule.IncludeWeekends,
                    CellSpacing = 2,
                    CellBorderRadius = 4,
                    CellWidthScale = 0.7,
                    CellHeightScale = 0.6
                }
            };

            _exportSchedules.Add(exportSchedule);
            SaveExportsToDisk();
            return exportSchedule;
        }

        public static void AddExportSchedule(Schedule schedule, string customName = null)
        {
            CreateExportItem(schedule, customName);
        }

        public static bool DeleteExportSchedule(string id)
        {
            var schedule = _exportSchedules.FirstOrDefault(s => s.Id == id);
            if (schedule != null)
            {
                bool removed = _exportSchedules.Remove(schedule);
                if (removed)
                {
                    _stateSave.DeleteExport(id);
                    SaveExportsToDisk();
                }
                return removed;
            }
            return false;
        }

        public static bool RenameExportSchedule(string id, string newName)
        {
            var schedule = _exportSchedules.FirstOrDefault(s => s.Id == id);
            if (schedule != null && !string.IsNullOrWhiteSpace(newName))
            {
                schedule.Name = newName;
                SaveExportsToDisk();
                return true;
            }
            return false;
        }

        public static List<TemplateInfo> GetAvailableTemplates()
        {
            return _templates;
        }

        public static void AddUserTemplate(string name, string imagePath)
        {
            if (!_templates.Any(t => t.Name == name))
            {
                var newTemplate = new TemplateInfo
                {
                    Name = name,
                    DisplayName = name,
                    Description = "Custom template",
                    ColorHex = "#6366F1",
                    Icon = null,
                    BackgroundImagePath = imagePath,
                    IsCustom = true
                };
                _templates.Add(newTemplate);
            }
        }

        public static void RemoveTemplate(string name)
        {
            if (name != "Default")
            {
                var template = _templates.FirstOrDefault(t => t.Name == name && t.IsCustom);
                if (template != null)
                {
                    _templates.Remove(template);
                }
            }
        }

        public static void UpdateExportSettings(string exportId, ExportSettings settings)
        {
            var exportSchedule = _exportSchedules.FirstOrDefault(s => s.Id == exportId);
            if (exportSchedule != null)
            {
                exportSchedule.Settings = settings;
                SaveExportsToDisk();
            }
        }

        /// <summary>
        /// Exports the schedule to a PowerPoint (.pptx) file.
        /// </summary>
        /// <param name="exportSchedule">The export item containing the schedule and settings.</param>
        /// <param name="progress">Optional progress reporter for UI feedback.</param>
        /// <param name="outputPath">Optional explicit output path; otherwise a timestamped file is created in Documents\SchedulerExports.</param>
        /// <returns>An ExportResult with the generated .pptx path or an error message.</returns>
        public static async Task<ExportResult> ExportToPowerPointAsync(ExportSchedule exportSchedule, IProgress<string> progress = null, string outputPath = null)
        {
            return await Task.Run(() =>
            {
                ExportResult result = new ExportResult();
                try
                {
                    progress?.Report("Preparing export...");

                    var schedule = exportSchedule.SourceSchedule;
                    var settings = exportSchedule.Settings;

                    string pptxPath;
                    if (!string.IsNullOrEmpty(outputPath))
                    {
                        pptxPath = outputPath;
                    }
                    else
                    {
                        string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SchedulerExports");
                        Directory.CreateDirectory(outputDir);

                        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        string fileName = $"{exportSchedule.Name.Replace(" ", "_")}_{timestamp}.pptx";
                        pptxPath = Path.Combine(outputDir, fileName);
                    }

                    progress?.Report("Creating PowerPoint presentation...");
                    ExportToPowerPointFile(pptxPath, exportSchedule);

                    result.PptxPath = pptxPath;
                    result.Success = true;

                    exportSchedule.OutputPath = pptxPath;
                    progress?.Report("Export completed successfully!");

                    return result;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                    throw new Exception($"Failed to export: {ex.Message}", ex);
                }
            });
        }

        private static void ExportToPowerPointFile(string filePath, ExportSchedule exportSchedule)
        {
            var schedule = exportSchedule.SourceSchedule;
            var settings = exportSchedule.Settings;

            settings.IncludeWeekends = schedule.IncludeWeekends;

            // Find the selected template file
            string templatePath = null;
            if (!string.IsNullOrEmpty(settings.TemplateName) && settings.TemplateName != "Default")
            {
                var template = _templates.FirstOrDefault(t => t.Name == settings.TemplateName);
                if (template != null && !string.IsNullOrEmpty(template.TemplateFilePath))
                {
                    templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, template.TemplateFilePath);
                    if (!File.Exists(templatePath))
                    {
                        templatePath = Path.Combine(Directory.GetCurrentDirectory(), template.TemplateFilePath);
                    }
                }
            }

            // Create PowerPoint presentation - use template as base if selected
            PowerPoint powerPoint;
            string templateBackgroundImage = null;
            long backgroundWidth = 12192000;  // default 13.33"
            long backgroundHeight = 6858000;  // default 7.5"
            if (!string.IsNullOrEmpty(templatePath) && File.Exists(templatePath))
            {
                powerPoint = new PowerPoint(templatePath, true, new PowerPointProperties());

                // Remove all template slides so no extra blank slide remains.
                // The template design will be re-applied as a background image on
                // each new schedule slide instead.
                int templateSlideCount = powerPoint.GetSlideCount();
                for (int i = templateSlideCount - 1; i >= 0; i--)
                {
                    powerPoint.RemoveSlideByIndex(i);
                }

                // Resolve the template's background PNG (from Icons folder)
                var template = _templates.FirstOrDefault(t => t.Name == settings.TemplateName);
                if (template != null && !string.IsNullOrEmpty(template.BackgroundImagePath))
                {
                    templateBackgroundImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, template.BackgroundImagePath);
                    if (!File.Exists(templateBackgroundImage))
                    {
                        templateBackgroundImage = Path.Combine(Directory.GetCurrentDirectory(), template.BackgroundImagePath);
                    }
                }

                // Read the template's actual slide dimensions (sldSz cx/cy in EMU)
                // so the background picture fills the whole slide exactly.
                long slidCx, slidCy;
                if (TryGetTemplateSlideSize(templatePath, out slidCx, out slidCy) && slidCx > 0 && slidCy > 0)
                {
                    backgroundWidth = slidCx;
                    backgroundHeight = slidCy;
                }
            }
            else
            {
                powerPoint = new PowerPoint();
            }

            try
            {
                // Get schedule weeks
                var weeks = GetScheduleWeeks(schedule, settings);

                // Create slides for each week
                foreach (var week in weeks)
                {
                    // Add blank slide
                    powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
                }

                // Generate content for each slide
                for (int weekIndex = 0; weekIndex < weeks.Count; weekIndex++)
                {
                    var week = weeks[weekIndex];
                    Slide slide = powerPoint.GetSlideByIndex(weekIndex);

                    if (slide != null)
                    {
                        // Add template background image as full-slide picture.
                        // Read the PNG's original size and scale to COVER the slide
                        if (!string.IsNullOrEmpty(templateBackgroundImage) && File.Exists(templateBackgroundImage))
                        {
                            try
                            {
                                long picW, picH;
                                if (GetImageFillMetrics(templateBackgroundImage, backgroundWidth, backgroundHeight, out picW, out picH))
                                {
                                    var pictureSetting = new PictureSetting
                                    {
                                        x = 0,
                                        y = 0,
                                        width = (uint)picW,
                                        height = (uint)picH
                                    };
                                    slide.AddPicture(templateBackgroundImage, pictureSetting);
                                }
                            }
                            catch { }
                        }

                        GenerateSlideContent(slide, schedule, settings, week);
                    }
                }

                powerPoint.SaveAs(filePath);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Reads the original pixel dimensions of a PNG/JPG so the background picture
        /// can be scaled to fill the whole slide while preserving aspect ratio.
        /// Returns a scale factor and the scaled width/height in EMU.
        /// </summary>
        private static bool GetImageFillMetrics(string imagePath, long slideCx, long slideCy,
            out long outWidthEmu, out long outHeightEmu)
        {
            outWidthEmu = slideCx;
            outHeightEmu = slideCy;
            try
            {
                using (var img = System.Drawing.Image.FromFile(imagePath))
                {
                    double origW = img.Width;
                    double origH = img.Height;
                    if (origW <= 0 || origH <= 0) return false;

                    // "Cover" scaling: scale so the image fully covers the slide,
                    // preserving aspect ratio (may overflow edges, which is fine for a background).
                    double scaleW = slideCx / origW;
                    double scaleH = slideCy / origH;
                    double scale = Math.Max(scaleW, scaleH);

                    outWidthEmu = (long)(origW * scale);
                    outHeightEmu = (long)(origH * scale);
                    return outWidthEmu > 0 && outHeightEmu > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Reads the template's actual slide dimensions (sldSz cx/cy in EMU) from
        /// its presentation.xml so a full-slide background picture fills the whole slide.
        /// </summary>
        private static bool TryGetTemplateSlideSize(string templatePath, out long cx, out long cy)
        {
            cx = 0;
            cy = 0;
            try
            {
                using (var zip = ZipFile.OpenRead(templatePath))
                {
                    var entry = zip.GetEntry("ppt/presentation.xml");
                    if (entry == null) return false;

                    using (var reader = new StreamReader(entry.Open()))
                    {
                        string xml = reader.ReadToEnd();
                        // Find sldSz element: <p:sldSz cx="18288000" cy="10287000" type="screen16x9"/>
                        int idx = xml.IndexOf("sldSz");
                        if (idx < 0) return false;

                        // Extract cx="..."
                        int cxStart = xml.IndexOf("cx=\"", idx);
                        if (cxStart < 0) return false;
                        cxStart += 4;
                        int cxEnd = xml.IndexOf("\"", cxStart);
                        if (cxEnd < 0) return false;
                        long.TryParse(xml.Substring(cxStart, cxEnd - cxStart), out cx);

                        // Extract cy="..."
                        int cyStart = xml.IndexOf("cy=\"", idx);
                        if (cyStart < 0) return false;
                        cyStart += 4;
                        int cyEnd = xml.IndexOf("\"", cyStart);
                        if (cyEnd < 0) return false;
                        long.TryParse(xml.Substring(cyStart, cyEnd - cyStart), out cy);

                        return cx > 0 && cy > 0;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        private static void GenerateSlideContent(Slide slide, Schedule schedule, ExportSettings settings, DateTime[] weekRange)
        {
            AddTitleTextBox(slide, settings, weekRange);
            CreateScheduleGrid(slide, schedule, settings, weekRange);
        }

        /// <summary>
        /// Builds a title from the slide's date range, e.g.:
        ///   Sep 24 - 28        (same month: show month once)
        ///   Sep 28 - Oct 04    (different months: show both months)
        /// When weekends are excluded, the title reflects only the weekday dates
        /// (e.g. Nov 17 - 21, not Nov 17 - 23).
        /// </summary>
        private static string BuildSlideTitle(DateTime[] weekRange, bool includeWeekends)
        {
            var start = weekRange[0];
            var end = weekRange[1];

            // Skip weekend days at the start if weekends are excluded
            if (!includeWeekends)
            {
                while (start.DayOfWeek == DayOfWeek.Saturday || start.DayOfWeek == DayOfWeek.Sunday)
                {
                    start = start.AddDays(1);
                }
                while (end.DayOfWeek == DayOfWeek.Saturday || end.DayOfWeek == DayOfWeek.Sunday)
                {
                    end = end.AddDays(-1);
                }
            }

            string startMonth = start.ToString("MMM");
            string endMonth = end.ToString("MMM");
            string startDay = start.ToString("dd");
            string endDay = end.ToString("dd");

            if (start.Year == end.Year && start.Month == end.Month)
            {
                return $"{startMonth} {startDay} - {endDay}";
            }

            return $"{startMonth} {startDay} - {endMonth} {endDay}";
        }

        /// <summary>
        /// Adds a title textbox at the top of the slide.
        /// Position/size derived as proportions of the 20" x 11.25" slide:
        ///   width  = 8.08" / 20"    = 0.404   x2.5 -> ~20.2" (capped at 95% of slide width)
        ///   height = 1.78" / 11.25" = 0.158   x2.5 -> ~4.45"
        ///   x      = centered horizontally
        ///   y      = 0"   / 11.25"  = 0
        /// </summary>
        private static void AddTitleTextBox(Slide slide, ExportSettings settings, DateTime[] weekRange)
        {
            try
            {
                long slideWidth = 18288000;   // 20"
                long slideHeight = 10287000;  // 11.25"

                long w = (long)(slideWidth * 0.404 * 2.5);
                if (w > (long)(slideWidth * 0.95)) w = (long)(slideWidth * 0.95);
                long h = (long)(slideHeight * 0.158 * 2.5);   // ~4.45"
                long x = (long)((slideWidth - w) / 2);        // centered horizontally
                long y = 0;                                   // top of slide

                var textBoxSetting = new TextBoxSetting
                {
                    x = (int)x,
                    y = (int)y,
                    width = (int)w,
                    height = (int)h,
                    horizontalAlignment = HorizontalAlignmentValues.CENTER
                };

                var textBlock = new TextBlock();
                textBlock.textValue = BuildSlideTitle(weekRange, settings.IncludeWeekends);
                textBlock.fontSize = 100;
                textBlock.fontColor = settings.TitleColor.Replace("#", "");
                textBlock.isBold = false;
                textBlock.fontFamily = "Forte";

                textBoxSetting.textBlocks = new[] { textBlock };

                slide.AddTextBox(textBoxSetting);
            }
            catch { }
        }

        private static void CreateScheduleGrid(Slide slide, Schedule schedule, ExportSettings settings, DateTime[] weekRange)
        {
            int days = CalculateDaysInWeek(weekRange[0], weekRange[1], settings.IncludeWeekends);
            int intervals = schedule.ShiftIntervals;

            if (days <= 0 || intervals <= 0) return;

            long slideWidth = 18288000;   // 20"
            long slideHeight = 10287000;  // 11.25"

            long tableWidth = (long)(slideWidth * 0.8215);   // ~16.43"
            long tableHeight = (long)(slideHeight * 0.731);  // ~8.22" (renders ~8.88" after library padding)
            long gridStartX = (long)(slideWidth * 0.0895);   // ~1.79" from top-left
            long gridStartY = (long)(slideHeight * 0.167);   // ~1.88" from top-left

            long cellSpacing = (long)(settings.CellSpacing * 1200);

            long availableWidth = tableWidth - (days * cellSpacing);
            long availableHeight = tableHeight - (intervals * cellSpacing);

            long cellWidth = availableWidth / (days + 1);
            long cellHeight = availableHeight / (intervals + 1);

            if (cellWidth <= 0 || cellHeight <= 0) return;

            long totalGridWidth = (days + 1) * cellWidth + days * cellSpacing;
            long totalGridHeight = (intervals + 1) * cellHeight + intervals * cellSpacing;

            // Determine the max number of names in any single cell on this slide.
            // If cells are crowded, shrink the data/header fonts so text fits inside cells.
            int maxNamesInCell = 1;
            for (int interval = 0; interval < intervals; interval++)
            {
                DateTime dc = weekRange[0];
                int dcCount = 0;
                while (dcCount < days && dc <= weekRange[1])
                {
                    if (!settings.IncludeWeekends &&
                       (dc.DayOfWeek == DayOfWeek.Saturday || dc.DayOfWeek == DayOfWeek.Sunday))
                    {
                        dc = dc.AddDays(1);
                        continue;
                    }
                    string cid = GetCellId(schedule, dc, interval);
                    if (schedule.CellAssignments.ContainsKey(cid))
                    {
                        maxNamesInCell = Math.Max(maxNamesInCell, schedule.CellAssignments[cid].Count);
                    }
                    dc = dc.AddDays(1);
                    dcCount++;
                }
            }

            // Font scale: fewer names -> normal size, many names -> smaller.
            // The data font is clamped to a minimum of 14 so names stay readable.
            double fontScale = 1.0;
            if (maxNamesInCell > 3) fontScale = 0.55;
            else if (maxNamesInCell == 3) fontScale = 0.65;
            else if (maxNamesInCell == 2) fontScale = 0.8;
            else fontScale = 1.0;

            int dataFontSize = Math.Max(14, (int)(settings.FontSize * 0.5 * fontScale));

            // Build table data
            var tableRows = new List<TableRow>();

            // Header row (Time + Day headers)
            var headerRow = new TableRow
            {
                height = (int)cellHeight,
                rowBackground = settings.DaysRowColor.Replace("#", ""),
                textColor = settings.TextColor.Replace("#", ""),
                tableCells = new List<TableCell>()
            };

            // Time header cell
            var timeHeaderCell = new TableCell();
            timeHeaderCell.textValue = "Time";
            timeHeaderCell.rowSpan = 1;
            timeHeaderCell.columnSpan = 1;
            timeHeaderCell.cellBackground = settings.TimeCellColor.Replace("#", "");
            timeHeaderCell.textColor = settings.TextColor.Replace("#", "");
            timeHeaderCell.fontSize = (int)(settings.HeaderFontSize * 0.5);
            timeHeaderCell.fontFamily = "Forte";
            timeHeaderCell.isBold = true;
            timeHeaderCell.horizontalAlignment = HorizontalAlignmentValues.CENTER;
            timeHeaderCell.verticalAlignment = VerticalAlignmentValues.MIDDLE;
            headerRow.tableCells.Add(timeHeaderCell);

            // Day headers
            DateTime currentDate = weekRange[0];
            int dayCount = 0;

            while (dayCount < days && currentDate <= weekRange[1])
            {
                if (!settings.IncludeWeekends &&
                   (currentDate.DayOfWeek == DayOfWeek.Saturday ||
                    currentDate.DayOfWeek == DayOfWeek.Sunday))
                {
                    currentDate = currentDate.AddDays(1);
                    continue;
                }

                string dayLabel = $"{currentDate:ddd}\n{currentDate:MM/dd}";

                var dayHeaderCell = new TableCell();
                dayHeaderCell.textValue = dayLabel;
                dayHeaderCell.rowSpan = 1;
                dayHeaderCell.columnSpan = 1;
                dayHeaderCell.cellBackground = settings.DaysRowColor.Replace("#", "");
                dayHeaderCell.textColor = settings.TextColor.Replace("#", "");
                dayHeaderCell.fontSize = (int)(settings.HeaderFontSize * 0.5);
                dayHeaderCell.fontFamily = "Forte";
                dayHeaderCell.isBold = true;
                dayHeaderCell.horizontalAlignment = HorizontalAlignmentValues.CENTER;
                dayHeaderCell.verticalAlignment = VerticalAlignmentValues.MIDDLE;
                headerRow.tableCells.Add(dayHeaderCell);

                currentDate = currentDate.AddDays(1);
                dayCount++;
            }

            tableRows.Add(headerRow);

            // Time rows
            for (int interval = 0; interval < intervals; interval++)
            {
                double startTime = schedule.OpeningHour + (interval * schedule.ShiftLengthHours);
                double endTime = Math.Min(startTime + schedule.ShiftLengthHours, schedule.ClosingHour);

                string timeLabel = $"{FormatTimeFromHour(startTime)}\nto\n{FormatTimeFromHour(endTime)}";

                var timeRow = new TableRow
                {
                    height = (int)cellHeight,
                    rowBackground = settings.TimeCellColor.Replace("#", ""),
                    textColor = settings.TextColor.Replace("#", ""),
                    tableCells = new List<TableCell>()
                };

                // Time cell
                var timeCell = new TableCell();
                timeCell.textValue = timeLabel;
                timeCell.rowSpan = 1;
                timeCell.columnSpan = 1;
                timeCell.cellBackground = settings.TimeCellColor.Replace("#", "");
                timeCell.textColor = settings.TextColor.Replace("#", "");
                timeCell.fontSize = (int)(settings.HeaderFontSize * 0.5);
                timeCell.fontFamily = "Forte";
                timeCell.isBold = true;
                timeCell.horizontalAlignment = HorizontalAlignmentValues.CENTER;
                timeCell.verticalAlignment = VerticalAlignmentValues.MIDDLE;
                timeRow.tableCells.Add(timeCell);

                // Day cells for this time slot
                currentDate = weekRange[0];
                dayCount = 0;

                while (dayCount < days && currentDate <= weekRange[1])
                {
                    if (!settings.IncludeWeekends &&
                       (currentDate.DayOfWeek == DayOfWeek.Saturday ||
                        currentDate.DayOfWeek == DayOfWeek.Sunday))
                    {
                        currentDate = currentDate.AddDays(1);
                        continue;
                    }

                    string cellId = GetCellId(schedule, currentDate, interval);
                    List<string> names = schedule.CellAssignments.ContainsKey(cellId)
                        ? schedule.CellAssignments[cellId]
                        : new List<string>();

                    string cellText = names.Count > 0 ? string.Join("\n", names) : "";
                    string cellColor = settings.NameCellColor;

                    var dataCell = new TableCell();
                    dataCell.textValue = cellText;
                    dataCell.rowSpan = 1;
                    dataCell.columnSpan = 1;
                    dataCell.cellBackground = cellColor.Replace("#", "");
                    dataCell.textColor = settings.TextColor.Replace("#", "");
                    dataCell.fontSize = dataFontSize;
                    dataCell.fontFamily = "Forte";
                    dataCell.isBold = false;
                    dataCell.horizontalAlignment = HorizontalAlignmentValues.CENTER;
                    dataCell.verticalAlignment = VerticalAlignmentValues.MIDDLE;
                    timeRow.tableCells.Add(dataCell);

                    currentDate = currentDate.AddDays(1);
                    dayCount++;
                }

                tableRows.Add(timeRow);
            }

            // Create table setting
            var tableSetting = new TableSetting();
            tableSetting.name = "ScheduleGrid";
            tableSetting.height = (uint)(totalGridHeight + cellSpacing * intervals);
            tableSetting.width = (uint)totalGridWidth;
            tableSetting.x = (uint)gridStartX;
            tableSetting.y = (uint)gridStartY;
            tableSetting.tableColumnWidth = new List<float>();

            // Set column widths
            for (int i = 0; i <= days; i++)
            {
                tableSetting.tableColumnWidth.Add(cellWidth);
            }

            // Add table to slide
            slide.AddTable(tableRows.ToArray(), tableSetting);
        }

        private static string GetCellId(Schedule schedule, DateTime date, int interval)
        {
            int dayIndex = GetDayIndex(schedule, date);
            return $"cell_{dayIndex}_{interval}";
        }

        private static int GetDayIndex(Schedule schedule, DateTime date)
        {
            var startDate = schedule.StartDate.Date;
            var targetDate = date.Date;
            int daysSinceStart = (int)(targetDate - startDate).TotalDays;

            if (schedule.IncludeWeekends)
                return daysSinceStart;

            // When weekends are excluded, cell IDs use sequential weekday indices
            // (0 = first weekday, 1 = second weekday, etc.). This matches how
            // SetupPage populates CellAssignments (dates are stored by weekday index,
            // skipping weekends).
            int weekdayCount = 0;
            for (int i = 0; i <= daysSinceStart; i++)
            {
                var day = startDate.AddDays(i);
                if (day.DayOfWeek != DayOfWeek.Saturday && day.DayOfWeek != DayOfWeek.Sunday)
                {
                    weekdayCount++;
                }
            }

            return Math.Max(0, weekdayCount - 1);
        }

        private static int CalculateDaysInWeek(DateTime startDate, DateTime endDate, bool includeWeekends)
        {
            int days = 0;
            DateTime currentDate = startDate;

            while (currentDate <= endDate)
            {
                if (includeWeekends || (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday))
                {
                    days++;
                }
                currentDate = currentDate.AddDays(1);
            }

            return days;
        }

        private static string FormatTimeFromHour(double hour)
        {
            int hourInt = (int)Math.Floor(hour);
            int minutes = (int)Math.Round((hour - hourInt) * 60);

            if (minutes == 60)
            {
                hourInt++;
                minutes = 0;
            }

            string ampm = hourInt >= 12 ? "PM" : "AM";
            int displayHour = hourInt > 12 ? hourInt - 12 : (hourInt == 0 ? 12 : hourInt);

            return $"{displayHour}:{minutes:D2}{ampm}";
        }

        public static async Task<List<BitmapImage>> GenerateSlidePreviewsAsync(ExportSchedule exportSchedule, Action<int, int> progressCallback = null)
        {
            return await Task.Run(() =>
            {
                var previews = new List<BitmapImage>();
                var schedule = exportSchedule.SourceSchedule;
                var settings = exportSchedule.Settings;

                // Honor the user's weekend decision from the source schedule itself
                settings.IncludeWeekends = schedule.IncludeWeekends;

                var weeks = GetScheduleWeeks(schedule, settings);

                for (int weekIndex = 0; weekIndex < weeks.Count; weekIndex++)
                {
                    progressCallback?.Invoke(weekIndex + 1, weeks.Count);

                    var preview = GenerateWeekPreview(schedule, settings, weeks[weekIndex]);
                    if (preview != null)
                    {
                        previews.Add(preview);
                    }
                }

                return previews;
            });
        }

        private static BitmapImage GenerateWeekPreview(Schedule schedule, ExportSettings settings, DateTime[] weekRange)
        {
            try
            {
                var drawingVisual = new DrawingVisual();
                using (var drawingContext = drawingVisual.RenderOpen())
                {
                    // Draw template background image if selected
                    if (!string.IsNullOrEmpty(settings.BackgroundImagePath))
                    {
                        try
                        {
                            BitmapImage bgImage = null;
                            if (settings.BackgroundImagePath.StartsWith("pack://"))
                            {
                                var uri = new Uri(settings.BackgroundImagePath, UriKind.Absolute);
                                var streamInfo = System.Windows.Application.GetResourceStream(uri);
                                if (streamInfo != null)
                                {
                                    using (var ms = new MemoryStream())
                                    {
                                        streamInfo.Stream.CopyTo(ms);
                                        ms.Seek(0, SeekOrigin.Begin);
                                        bgImage = new BitmapImage();
                                        bgImage.BeginInit();
                                        bgImage.CacheOption = BitmapCacheOption.OnLoad;
                                        bgImage.StreamSource = ms;
                                        bgImage.EndInit();
                                        bgImage.Freeze();
                                    }
                                }
                            }
                            else
                            {
                                string bgPath = settings.BackgroundImagePath;
                                if (!Path.IsPathRooted(bgPath))
                                {
                                    string baseDirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, bgPath);
                                    string cwdPath = Path.Combine(Directory.GetCurrentDirectory(), bgPath);
                                    bgPath = File.Exists(baseDirPath) ? baseDirPath
                                           : File.Exists(cwdPath) ? cwdPath
                                           : bgPath;
                                }
                                if (File.Exists(bgPath))
                                {
                                    bgImage = new BitmapImage(new Uri(bgPath, UriKind.Absolute));
                                }
                            }

                            if (bgImage != null)
                            {
                                drawingContext.DrawImage(bgImage, new System.Windows.Rect(0, 0, 800, 450));
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        var backgroundColor = ColorHelper.ConvertFromString(settings.BackgroundColor);
                        if (backgroundColor is Color bgColor)
                        {
                            var backgroundBrush = new SolidColorBrush(bgColor);
                            drawingContext.DrawRectangle(backgroundBrush, null, new System.Windows.Rect(0, 0, 800, 450));
                        }
                    }

                    var titleText = new FormattedText(
                        BuildSlideTitle(weekRange, settings.IncludeWeekends),
                        System.Globalization.CultureInfo.CurrentCulture,
                        System.Windows.FlowDirection.LeftToRight,
                        new Typeface("Forte"),
                        settings.FontSize * 0.25,
                        new SolidColorBrush((Color)ColorHelper.ConvertFromString(settings.TitleColor)),
                        VisualTreeHelper.GetDpi(drawingVisual).PixelsPerDip);

                    drawingContext.DrawText(titleText, new System.Windows.Point(20, 20));

                    int days = CalculateDaysInWeek(weekRange[0], weekRange[1], settings.IncludeWeekends);
                    int intervals = Math.Min(8, schedule.ShiftIntervals);

                    double cellSpacing = settings.CellSpacing;
                    double cellWidth = (760.0 - (days * cellSpacing)) / (days + 1) * settings.CellWidthScale;
                    double cellHeight = (350.0 - (intervals * cellSpacing)) / (intervals + 1) * settings.CellHeightScale;
                    double startX = 20;
                    double startY = 70;

                    double totalGridWidth = (days + 1) * cellWidth + days * cellSpacing;
                    startX += (760 - totalGridWidth) / 2;

                    double timeX = startX;
                    double timeY = startY;

                    var timeHeaderColor = ColorHelper.ConvertFromString(settings.TimeCellColor);
                    if (timeHeaderColor is Color timeColor)
                    {
                        var timeHeaderBrush = new SolidColorBrush(timeColor);
                        drawingContext.DrawRectangle(timeHeaderBrush, new Pen(Brushes.Gray, 1),
                            new System.Windows.Rect(timeX, timeY, cellWidth, cellHeight));
                    }

                    var timeText = new FormattedText(
                        "Time",
                        System.Globalization.CultureInfo.CurrentCulture,
                        System.Windows.FlowDirection.LeftToRight,
                        new Typeface("Arial"),
                        settings.HeaderFontSize * 0.25,
                        new SolidColorBrush((Color)ColorHelper.ConvertFromString(settings.TextColor)),
                        VisualTreeHelper.GetDpi(drawingVisual).PixelsPerDip);

                    double textX = timeX + (cellWidth - timeText.Width) / 2;
                    double textY = timeY + (cellHeight - timeText.Height) / 2;
                    drawingContext.DrawText(timeText, new System.Windows.Point(textX, textY));

                    DateTime currentDate = weekRange[0];
                    int dayCount = 0;

                    while (dayCount < days && currentDate <= weekRange[1])
                    {
                        if (!settings.IncludeWeekends &&
                           (currentDate.DayOfWeek == DayOfWeek.Saturday ||
                            currentDate.DayOfWeek == DayOfWeek.Sunday))
                        {
                            currentDate = currentDate.AddDays(1);
                            continue;
                        }

                        double x = startX + ((dayCount + 1) * (cellWidth + cellSpacing));
                        double y = startY;

                        var daysRowColor = ColorHelper.ConvertFromString(settings.DaysRowColor);
                        if (daysRowColor is Color dayColor)
                        {
                            var dayBrush = new SolidColorBrush(dayColor);
                            drawingContext.DrawRectangle(dayBrush, new Pen(Brushes.Gray, 1),
                                new System.Windows.Rect(x, y, cellWidth, cellHeight));
                        }

                        var dayLabel = $"{currentDate:ddd}\n{currentDate:MM/dd}";
                        var dayText = new FormattedText(
                            dayLabel,
                            System.Globalization.CultureInfo.CurrentCulture,
                            System.Windows.FlowDirection.LeftToRight,
                            new Typeface("Arial"),
                            settings.HeaderFontSize * 0.25,
                            new SolidColorBrush((Color)ColorHelper.ConvertFromString(settings.TextColor)),
                            VisualTreeHelper.GetDpi(drawingVisual).PixelsPerDip);

                        textX = x + (cellWidth - dayText.Width) / 2;
                        textY = y + (cellHeight - dayText.Height) / 2;
                        drawingContext.DrawText(dayText, new System.Windows.Point(textX, textY));

                        currentDate = currentDate.AddDays(1);
                        dayCount++;
                    }

                    for (int interval = 0; interval < intervals; interval++)
                    {
                        double startTime = schedule.OpeningHour + (interval * schedule.ShiftLengthHours);
                        double endTime = Math.Min(startTime + schedule.ShiftLengthHours, schedule.ClosingHour);

                        double x = startX;
                        double y = startY + ((interval + 1) * (cellHeight + cellSpacing));

                        var timeCellColor = ColorHelper.ConvertFromString(settings.TimeCellColor);
                        if (timeCellColor is Color tcColor)
                        {
                            var timeColumnBrush = new SolidColorBrush(tcColor);
                            drawingContext.DrawRectangle(timeColumnBrush, new Pen(Brushes.Gray, 1),
                                new System.Windows.Rect(x, y, cellWidth, cellHeight));
                        }

                        var timeLabel = $"{FormatTimeFromHour(startTime)}";
                        var timeLabelText = new FormattedText(
                            timeLabel,
                            System.Globalization.CultureInfo.CurrentCulture,
                            System.Windows.FlowDirection.LeftToRight,
                            new Typeface("Arial"),
                            settings.HeaderFontSize * 0.25,
                            new SolidColorBrush((Color)ColorHelper.ConvertFromString(settings.TextColor)),
                            VisualTreeHelper.GetDpi(drawingVisual).PixelsPerDip);

                        textX = x + (cellWidth - timeLabelText.Width) / 2;
                        textY = y + (cellHeight - timeLabelText.Height) / 2;
                        drawingContext.DrawText(timeLabelText, new System.Windows.Point(textX, textY));
                    }

                    currentDate = weekRange[0];
                    dayCount = 0;

                    while (dayCount < days && currentDate <= weekRange[1])
                    {
                        if (!settings.IncludeWeekends &&
                           (currentDate.DayOfWeek == DayOfWeek.Saturday ||
                            currentDate.DayOfWeek == DayOfWeek.Sunday))
                        {
                            currentDate = currentDate.AddDays(1);
                            continue;
                        }

                        for (int interval = 0; interval < intervals; interval++)
                        {
                            string cellId = GetCellId(schedule, currentDate, interval);
                            List<string> names = schedule.CellAssignments.ContainsKey(cellId)
                                ? schedule.CellAssignments[cellId]
                                : new List<string>();

                            double x = startX + ((dayCount + 1) * (cellWidth + cellSpacing));
                            double y = startY + ((interval + 1) * (cellHeight + cellSpacing));

                            SolidColorBrush cellBrush;
                            var nameCellColor = ColorHelper.ConvertFromString(settings.NameCellColor);
                            if (nameCellColor is Color ncColor)
                            {
                                cellBrush = new SolidColorBrush(ncColor);
                            }
                            else
                            {
                                cellBrush = new SolidColorBrush(Colors.White);
                            }

                            drawingContext.DrawRectangle(cellBrush, new Pen(Brushes.Gray, 0.5),
                                new System.Windows.Rect(x, y, cellWidth, cellHeight));

                            if (names.Count > 0)
                            {
                                var nameText = new FormattedText(
                                    string.Join("\n", names),
                                    System.Globalization.CultureInfo.CurrentCulture,
                                    System.Windows.FlowDirection.LeftToRight,
                                    new Typeface("Arial"),
                                    settings.FontSize * 0.20,
                                    new SolidColorBrush((Color)ColorHelper.ConvertFromString(settings.TextColor)),
                                    VisualTreeHelper.GetDpi(drawingVisual).PixelsPerDip);

                                textX = x + (cellWidth - nameText.Width) / 2;
                                textY = y + (cellHeight - nameText.Height) / 2;
                                drawingContext.DrawText(nameText, new System.Windows.Point(textX, textY));
                            }
                        }

                        currentDate = currentDate.AddDays(1);
                        dayCount++;
                    }
                }

                var renderTarget = new RenderTargetBitmap(800, 450, 96, 96, PixelFormats.Pbgra32);
                renderTarget.Render(drawingVisual);

                var bitmapImage = new BitmapImage();
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(renderTarget));

                using (var stream = new MemoryStream())
                {
                    encoder.Save(stream);
                    stream.Seek(0, SeekOrigin.Begin);

                    bitmapImage.BeginInit();
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.StreamSource = stream;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze();
                }

                return bitmapImage;
            }
            catch
            {
                return CreateErrorPreview();
            }
        }

        private static BitmapImage CreateErrorPreview()
        {
            var drawingVisual = new DrawingVisual();
            using (var drawingContext = drawingVisual.RenderOpen())
            {
                drawingContext.DrawRectangle(Brushes.White, null, new System.Windows.Rect(0, 0, 800, 450));

                var errorText = new FormattedText(
                    "Preview Unavailable",
                    System.Globalization.CultureInfo.CurrentCulture,
                    System.Windows.FlowDirection.LeftToRight,
                    new Typeface("Arial"),
                    20,
                    Brushes.Red,
                    VisualTreeHelper.GetDpi(drawingVisual).PixelsPerDip);

                drawingContext.DrawText(errorText, new System.Windows.Point(100, 100));
            }

            var renderTarget = new RenderTargetBitmap(800, 450, 96, 96, PixelFormats.Pbgra32);
            renderTarget.Render(drawingVisual);

            var bitmapImage = new BitmapImage();
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(renderTarget));

            using (var stream = new MemoryStream())
            {
                encoder.Save(stream);
                stream.Seek(0, SeekOrigin.Begin);

                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.StreamSource = stream;
                bitmapImage.EndInit();
                bitmapImage.Freeze();
            }

            return bitmapImage;
        }

        private static List<DateTime[]> GetScheduleWeeks(Schedule schedule, ExportSettings settings)
        {
            var weeks = new List<DateTime[]>();
            var currentDate = schedule.StartDate;
            var endDate = schedule.EndDate;

            while (currentDate <= endDate)
            {
                // When weekends are excluded, never start a slide on a weekend day.
                if (!settings.IncludeWeekends)
                {
                    while (currentDate.DayOfWeek == DayOfWeek.Saturday ||
                           currentDate.DayOfWeek == DayOfWeek.Sunday)
                    {
                        currentDate = currentDate.AddDays(1);
                    }
                }

                if (currentDate > endDate) break;

                var weekStart = currentDate;
                DateTime weekEnd;

                if (settings.SlidesPerWeek == 1)
                {
                    if (settings.IncludeWeekends)
                    {
                        weekEnd = weekStart.AddDays(6);
                    }
                    else
                    {
                        // Weekday-only week: end at the nearest Friday (Mon->Fri,
                        // or the first Friday after a mid-week start). Progress at
                        // most one working week (4 days) from the week start.
                        int daysToFriday = (int)DayOfWeek.Friday - (int)weekStart.DayOfWeek;
                        if (daysToFriday < 0) daysToFriday += 7;
                        int daysToAdd = Math.Min(4, daysToFriday);
                        weekEnd = weekStart.AddDays(daysToAdd);
                    }
                    if (weekEnd > endDate) weekEnd = endDate;
                }
                else
                {
                    int daysToAdd = settings.IncludeWeekends ? 6 : 4;
                    weekEnd = weekStart.AddDays(daysToAdd);
                    if (weekEnd > endDate) weekEnd = endDate;
                }

                weeks.Add(new[] { weekStart, weekEnd });

                if (settings.SlidesPerWeek == 1)
                {
                    currentDate = weekEnd.AddDays(1);
                }
                else
                {
                    currentDate = weekStart.AddDays(settings.IncludeWeekends ? 7 : 5);
                }
            }

            return weeks;
        }
    }

    public static class ColorHelper
    {
        public static object ConvertFromString(string colorString)
        {
            if (string.IsNullOrEmpty(colorString))
                return null;

            try
            {
                colorString = colorString.TrimStart('#');

                if (colorString.Length == 6)
                {
                    byte r = byte.Parse(colorString.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                    byte g = byte.Parse(colorString.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                    byte b = byte.Parse(colorString.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                    return Color.FromRgb(r, g, b);
                }
                else if (colorString.Length == 8)
                {
                    byte a = byte.Parse(colorString.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                    byte r = byte.Parse(colorString.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                    byte g = byte.Parse(colorString.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                    byte b = byte.Parse(colorString.Substring(6, 2), System.Globalization.NumberStyles.HexNumber);
                    return Color.FromArgb(a, r, g, b);
                }
            }
            catch { }

            return null;
        }
    }
}