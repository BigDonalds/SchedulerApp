using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using static SchedulerApp.MainWindow;
using A = DocumentFormat.OpenXml.Drawing;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using P = DocumentFormat.OpenXml.Presentation;
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
        public List<BitmapImage> SlidePreviews { get; set; } = new List<BitmapImage>();
        public int CurrentPreviewSlide { get; set; } = 0;
    }

    public class ExportSettings
    {
        public double FontSize { get; set; } = 14;
        public double CellPadding { get; set; } = 8;
        public double CellMargin { get; set; } = 3;
        public bool ShowGridLines { get; set; } = true;
        public bool ShowCellBackground { get; set; } = true;
        public string HeaderColor { get; set; } = "#4F46E5";
        public string CellColor { get; set; } = "#F8FAFC";
        public string TextColor { get; set; } = "#1F2937";
        public string NameCellColor { get; set; } = "#E0F2FE";
        public string TimeCellColor { get; set; } = "#FEF3C7";
        public string DaysRowColor { get; set; } = "#DBEAFE";
        public string TemplateName { get; set; } = "Default";
        public string BackgroundImagePath { get; set; }
        public double CellOpacity { get; set; } = 0.9;
        public double HeaderFontSize { get; set; } = 16;
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
    }

    public class TemplateInfo
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string ColorHex { get; set; }
        public string Icon { get; set; }
        public string BackgroundImagePath { get; set; }
        public bool IsCustom { get; set; }
    }

    public class ExportService
    {
        private static List<ExportSchedule> _exportSchedules = new List<ExportSchedule>();
        private static List<TemplateInfo> _templates = new List<TemplateInfo>();
        private static Dispatcher _dispatcher;

        static ExportService()
        {
            InitializeTemplates();
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
                    Name = "Default",
                    DisplayName = "Default",
                    Description = "Clean and professional look",
                    ColorHex = "#4F46E5",
                    Icon = "🏢",
                    BackgroundImagePath = null
                },
                new TemplateInfo
                {
                    Name = "Professional",
                    DisplayName = "Professional",
                    Description = "Corporate style for business",
                    ColorHex = "#111827",
                    Icon = "👔",
                    BackgroundImagePath = null
                },
                new TemplateInfo
                {
                    Name = "Modern",
                    DisplayName = "Modern",
                    Description = "Contemporary design",
                    ColorHex = "#3B82F6",
                    Icon = "🎨",
                    BackgroundImagePath = null
                },
                new TemplateInfo
                {
                    Name = "Simple",
                    DisplayName = "Simple",
                    Description = "Minimalist approach",
                    ColorHex = "#6B7280",
                    Icon = "📄",
                    BackgroundImagePath = null
                },
                new TemplateInfo
                {
                    Name = "Warm",
                    DisplayName = "Warm",
                    Description = "Friendly and inviting",
                    ColorHex = "#F59E0B",
                    Icon = "☀️",
                    BackgroundImagePath = null
                },
                new TemplateInfo
                {
                    Name = "Cool",
                    DisplayName = "Cool",
                    Description = "Calm and relaxed",
                    ColorHex = "#10B981",
                    Icon = "🌿",
                    BackgroundImagePath = null
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
                return _exportSchedules.Remove(schedule);
            }
            return false;
        }

        public static bool RenameExportSchedule(string id, string newName)
        {
            var schedule = _exportSchedules.FirstOrDefault(s => s.Id == id);
            if (schedule != null && !string.IsNullOrWhiteSpace(newName))
            {
                schedule.Name = newName;
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
                    Icon = "⭐",
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
            }
        }

        public static async Task<string> ExportToPowerPointAsync(ExportSchedule exportSchedule, IProgress<string> progress = null)
        {
            return await Task.Run(() =>
            {
                string filePath = null;
                try
                {
                    progress?.Report("Preparing export...");

                    var schedule = exportSchedule.SourceSchedule;
                    var settings = exportSchedule.Settings;

                    string outputDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SchedulerExports");
                    Directory.CreateDirectory(outputDir);

                    string fileName = $"{exportSchedule.Name.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";
                    filePath = System.IO.Path.Combine(outputDir, fileName);

                    progress?.Report("Creating presentation...");

                    using (var presentation = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
                    {
                        var presentationPart = presentation.AddPresentationPart();
                        presentationPart.Presentation = new Presentation(
                            new SlideIdList(),
                            new SlideSize() { Cx = 12192000, Cy = 6858000 },
                            new NotesSize() { Cx = 6858000, Cy = 9144000 },
                            new DefaultTextStyle()
                        );

                        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                        slideMasterPart.SlideMaster = CreateSlideMaster();
                        slideMasterPart.SlideMaster.Save();

                        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
                        slideLayoutPart.SlideLayout = CreateSlideLayout();
                        slideLayoutPart.SlideLayout.Save();

                        presentationPart.Presentation.Append(
                            new SlideMasterIdList(
                                new SlideMasterId() { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }
                            )
                        );

                        var weeks = GetScheduleWeeks(schedule, settings);
                        uint slideId = 256;

                        for (int weekIndex = 0; weekIndex < weeks.Count; weekIndex++)
                        {
                            var week = weeks[weekIndex];
                            progress?.Report($"Creating slide {weekIndex + 1} of {weeks.Count}...");

                            var slidePart = presentationPart.AddNewPart<SlidePart>($"slideId{weekIndex + 1}");
                            var slide = CreateSlide();
                            slidePart.Slide = slide;
                            slidePart.AddPart(slideLayoutPart);
                            slide.Save();

                            GenerateSlideContent(slidePart, schedule, settings, week);
                            slidePart.Slide.Save();

                            var slideIdItem = new SlideId
                            {
                                Id = slideId++,
                                RelationshipId = presentationPart.GetIdOfPart(slidePart)
                            };
                            presentationPart.Presentation.SlideIdList.Append(slideIdItem);
                        }

                        presentationPart.Presentation.Save();
                    }

                    exportSchedule.OutputPath = filePath;
                    progress?.Report("Export completed!");

                    return filePath;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to export to PowerPoint: {ex.Message}", ex);
                }
            });
        }

        private static List<DateTime[]> GetScheduleWeeks(Schedule schedule, ExportSettings settings)
        {
            var weeks = new List<DateTime[]>();
            var currentDate = schedule.StartDate;
            var endDate = schedule.EndDate;

            while (currentDate <= endDate)
            {
                var weekStart = currentDate;
                DateTime weekEnd;

                if (settings.SlidesPerWeek == 1)
                {
                    weekEnd = weekStart.AddDays(6);
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

        private static SlideMaster CreateSlideMaster()
        {
            var slideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree()),
                new ColorMapOverride(new MasterColorMapping()));
            return slideMaster;
        }

        private static SlideLayout CreateSlideLayout()
        {
            var slideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree()),
                new ColorMapOverride(new MasterColorMapping()));
            return slideLayout;
        }

        private static Slide CreateSlide()
        {
            var slide = new Slide();
            var commonSlideData = new CommonSlideData();
            var shapeTree = new ShapeTree();

            var nonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1U, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties());

            shapeTree.Append(nonVisualGroupShapeProperties);
            shapeTree.Append(new GroupShapeProperties(new A.TransformGroup()));
            commonSlideData.Append(shapeTree);

            slide.Append(commonSlideData);
            slide.Append(new ColorMapOverride(new MasterColorMapping()));
            slide.Append(new SlideLayoutId() { Id = 2147483649U, RelationshipId = "rId1" });

            return slide;
        }

        private static void GenerateSlideContent(SlidePart slidePart, Schedule schedule, ExportSettings settings, DateTime[] weekRange)
        {
            long slideWidth = 12192000L;
            long slideHeight = 6858000L;

            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            var elementsToKeep = new List<OpenXmlElement>();
            foreach (var element in shapeTree.Elements())
            {
                if (element is P.NonVisualGroupShapeProperties ||
                    element is GroupShapeProperties)
                {
                    elementsToKeep.Add(element);
                }
            }

            shapeTree.RemoveAllChildren();
            foreach (var element in elementsToKeep)
            {
                shapeTree.Append(element);
            }

            _shapeIdCounter = 10000;

            AddBackgroundColor(shapeTree, slideWidth, slideHeight, settings.BackgroundColor);

            var template = _templates.FirstOrDefault(t => t.Name == settings.TemplateName);
            if (template != null && !string.IsNullOrEmpty(template.BackgroundImagePath) && File.Exists(template.BackgroundImagePath))
            {
                AddBackgroundImage(slidePart, shapeTree, template.BackgroundImagePath, slideWidth, slideHeight);
            }
            else if (!string.IsNullOrEmpty(settings.BackgroundImagePath) && File.Exists(settings.BackgroundImagePath))
            {
                AddBackgroundImage(slidePart, shapeTree, settings.BackgroundImagePath, slideWidth, slideHeight);
            }

            AddSlideTitle(shapeTree, schedule.Name, weekRange, slideWidth, settings);
            CreateScheduleGrid(shapeTree, schedule, settings, weekRange, slideWidth, slideHeight);
        }

        private static void AddSlideTitle(ShapeTree shapeTree, string scheduleName, DateTime[] weekRange, long slideWidth, ExportSettings settings)
        {
            var titleShape = new P.Shape();

            var nonVisualProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1000U, Name = "Title" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties());

            var shapeProperties = new P.ShapeProperties();
            var transform2D = new A.Transform2D();
            var offset = new A.Offset() { X = 500000L, Y = 200000L };
            var extents = new A.Extents() { Cx = slideWidth - 1000000L, Cy = 400000L };

            transform2D.Append(offset);
            transform2D.Append(extents);
            shapeProperties.Append(transform2D);
            shapeProperties.Append(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

            var textBody = new P.TextBody();
            var bodyProperties = new A.BodyProperties()
            {
                Anchor = A.TextAnchoringTypeValues.Center,
                Wrap = A.TextWrappingValues.Square
            };

            var listStyle = new A.ListStyle();
            var paragraph = new A.Paragraph();

            var paragraphProperties = new A.ParagraphProperties()
            {
                Alignment = A.TextAlignmentTypeValues.Center
            };

            var runProperties = new A.RunProperties()
            {
                FontSize = 3200,
                Bold = true
            };

            runProperties.Append(new A.SolidFill(
                new A.RgbColorModelHex()
                {
                    Val = settings.TextColor.Replace("#", "")
                }));

            var titleText = $"{scheduleName}\nWeek of {weekRange[0]:MMM dd} - {weekRange[1]:MMM dd, yyyy}";

            var textRun = new A.Run();
            textRun.Append(runProperties);
            textRun.Append(new A.Text(titleText));

            paragraph.Append(paragraphProperties);
            paragraph.Append(textRun);

            textBody.Append(bodyProperties);
            textBody.Append(listStyle);
            textBody.Append(paragraph);

            titleShape.Append(nonVisualProperties);
            titleShape.Append(shapeProperties);
            titleShape.Append(textBody);

            shapeTree.Append(titleShape);
        }

        private static void AddBackgroundColor(ShapeTree shapeTree, long slideWidth, long slideHeight, string colorHex)
        {
            var shape = new P.Shape();

            var nonVisualProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 2U, Name = "BackgroundColor" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties());

            var shapeProperties = new P.ShapeProperties();
            var solidFill = new A.SolidFill();
            solidFill.Append(new A.RgbColorModelHex() { Val = colorHex.Replace("#", "") });

            var transform2D = new A.Transform2D();
            var offset = new A.Offset() { X = 0L, Y = 0L };
            var extents = new A.Extents() { Cx = slideWidth, Cy = slideHeight };

            transform2D.Append(offset);
            transform2D.Append(extents);
            shapeProperties.Append(transform2D);
            shapeProperties.Append(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });
            shapeProperties.Append(solidFill);

            shape.Append(nonVisualProperties);
            shape.Append(shapeProperties);

            shapeTree.Append(shape);
        }

        private static void AddBackgroundImage(SlidePart slidePart, ShapeTree shapeTree, string imagePath, long slideWidth, long slideHeight)
        {
            try
            {
                var imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                using (var stream = new FileStream(imagePath, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                var picture = new P.Picture();
                var nonVisualPicProperties = new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties() { Id = 3U, Name = "Background" },
                    new P.NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties());

                var blipFill = new P.BlipFill();
                var blip = new A.Blip() { Embed = slidePart.GetIdOfPart(imagePart) };
                var stretch = new A.Stretch();
                stretch.Append(new A.FillRectangle());

                blipFill.Append(blip);
                blipFill.Append(stretch);

                var shapeProperties = new P.ShapeProperties();
                var transform2D = new A.Transform2D();
                var offset = new A.Offset() { X = 0L, Y = 0L };
                var extents = new A.Extents() { Cx = slideWidth, Cy = slideHeight };

                transform2D.Append(offset);
                transform2D.Append(extents);
                shapeProperties.Append(transform2D);
                shapeProperties.Append(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

                picture.Append(nonVisualPicProperties);
                picture.Append(blipFill);
                picture.Append(shapeProperties);

                shapeTree.Append(picture);
            }
            catch (Exception ex)
            {
            }
        }

        private static void CreateScheduleGrid(ShapeTree shapeTree, Schedule schedule, ExportSettings settings, DateTime[] weekRange, long slideWidth, long slideHeight)
        {
            int days = CalculateDaysInWeek(weekRange[0], weekRange[1], settings.IncludeWeekends);
            int intervals = schedule.ShiftIntervals;

            long margin = 500000L;
            long topMargin = 1200000L;

            long cellSpacing = (long)(settings.CellSpacing * 36000);
            long gridWidth = slideWidth - (2 * margin);
            long gridHeight = slideHeight - (topMargin + margin);

            long availableWidth = gridWidth - (days * cellSpacing);
            long availableHeight = gridHeight - (intervals * cellSpacing);

            long cellWidth = (long)(availableWidth / (days + 1) * settings.CellWidthScale);
            long cellHeight = (long)(availableHeight / (intervals + 1) * settings.CellHeightScale);

            long totalGridWidth = (days + 1) * cellWidth + days * cellSpacing;
            long gridStartX = margin + (gridWidth - totalGridWidth) / 2;

            long totalGridHeight = (intervals + 1) * cellHeight + intervals * cellSpacing;
            long gridStartY = topMargin + (gridHeight - totalGridHeight) / 2;

            long timeHeaderX = gridStartX;
            long timeHeaderY = gridStartY;
            AddTextBox(shapeTree, "Time", timeHeaderX, timeHeaderY, cellWidth, cellHeight,
                settings, isHeader: true, isTimeCell: true, shapeType: "Time Header");

            for (int interval = 0; interval < intervals; interval++)
            {
                double startTime = schedule.OpeningHour + (interval * schedule.ShiftLengthHours);
                double endTime = Math.Min(startTime + schedule.ShiftLengthHours, schedule.ClosingHour);

                string timeLabel = $"{FormatTimeFromHour(startTime)}\nto\n{FormatTimeFromHour(endTime)}";

                long x = gridStartX;
                long y = gridStartY + ((interval + 1) * (cellHeight + cellSpacing));

                AddTextBox(shapeTree, timeLabel, x, y, cellWidth, cellHeight,
                    settings, isHeader: true, isTimeCell: true, shapeType: $"Time Cell {interval}");
            }

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

                long x = gridStartX + ((dayCount + 1) * (cellWidth + cellSpacing));
                long y = gridStartY;

                AddTextBox(shapeTree, dayLabel, x, y, cellWidth, cellHeight,
                    settings, isHeader: true, shapeType: $"Day Header {dayCount}");

                currentDate = currentDate.AddDays(1);
                dayCount++;
            }

            currentDate = weekRange[0];
            dayCount = 0;
            int totalCells = 0;
            int cellsWithAssignments = 0;

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
                    totalCells++;

                    string cellId = GetCellId(schedule, currentDate, interval);
                    List<string> names = schedule.CellAssignments.ContainsKey(cellId)
                        ? schedule.CellAssignments[cellId]
                        : new List<string>();

                    if (names.Count > 0)
                        cellsWithAssignments++;

                    string cellText = names.Count > 0 ? string.Join("\n", names.Take(3)) : "";

                    long x = gridStartX + ((dayCount + 1) * (cellWidth + cellSpacing));
                    long y = gridStartY + ((interval + 1) * (cellHeight + cellSpacing));

                    string cellColor = settings.UseColorCoding
                        ? GetCellColor(names.Count, schedule.PeoplePerShift)
                        : settings.NameCellColor;

                    AddTextBox(shapeTree, cellText, x, y, cellWidth, cellHeight,
                        settings, isHeader: false, shapeType: $"Schedule Cell D{dayCount} I{interval}",
                        assignmentCount: names.Count);

                    if (names.Count > 3)
                    {
                        AddAssignmentBadge(shapeTree, x + cellWidth - 100000L, y + 50000L, names.Count - 3);
                    }
                }

                currentDate = currentDate.AddDays(1);
                dayCount++;
            }
        }

        private static void AddAssignmentBadge(ShapeTree shapeTree, long x, long y, int count)
        {
            var shape = new P.Shape();

            var nonVisualProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = GetNextShapeId(), Name = "Badge" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties());

            var shapeProperties = new P.ShapeProperties();

            var solidFill = new A.SolidFill();
            solidFill.Append(new A.RgbColorModelHex() { Val = "EF4444" });

            var outline = new A.Outline(
                new A.SolidFill(new A.RgbColorModelHex() { Val = "FFFFFF" }))
            {
                Width = 5000
            };

            var transform2D = new A.Transform2D();
            var offset = new A.Offset() { X = x, Y = y };
            var extents = new A.Extents() { Cx = 80000L, Cy = 80000L };

            transform2D.Append(offset);
            transform2D.Append(extents);
            shapeProperties.Append(transform2D);
            shapeProperties.Append(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Ellipse });
            shapeProperties.Append(solidFill);
            shapeProperties.Append(outline);

            var textBody = new P.TextBody();
            var bodyProperties = new A.BodyProperties()
            {
                Anchor = A.TextAnchoringTypeValues.Center
            };

            var paragraph = new A.Paragraph(new A.ParagraphProperties()
            {
                Alignment = A.TextAlignmentTypeValues.Center
            });

            var runProperties = new A.RunProperties()
            {
                FontSize = 700,
                Bold = true
            };

            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex() { Val = "FFFFFF" }));

            var textRun = new A.Run(runProperties, new A.Text($"+{count}"));

            paragraph.Append(textRun);
            textBody.Append(bodyProperties);
            textBody.Append(new A.ListStyle());
            textBody.Append(paragraph);

            shape.Append(nonVisualProperties);
            shape.Append(shapeProperties);
            shape.Append(textBody);

            shapeTree.Append(shape);
        }

        private static string GetCellId(Schedule schedule, DateTime date, int interval)
        {
            var daysSinceStart = (date - schedule.StartDate).Days;
            return $"cell_{daysSinceStart}_{interval}";
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

        private static string GetCellColor(int assignedCount, int requiredCount)
        {
            if (assignedCount == 0)
            {
                return "#FEE2E2";
            }
            else if (assignedCount < requiredCount)
            {
                return "#FEF3C7";
            }
            else
            {
                return "#D1FAE5";
            }
        }

        private static void AddTextBox(ShapeTree shapeTree, string text, long x, long y,
            long width, long height, ExportSettings settings, bool isHeader = false,
            bool isTimeCell = false, string shapeType = "TextBox",
            int assignmentCount = 0)
        {
            var shape = new P.Shape();

            uint shapeId = GetNextShapeId();
            string shapeName = $"{shapeType}_{shapeId}";

            var nonVisualProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = shapeId, Name = shapeName },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties());

            var shapeProperties = new P.ShapeProperties();

            var solidFill = new A.SolidFill();
            string colorHex = null;

            if (isHeader)
            {
                if (isTimeCell)
                {
                    colorHex = settings.TimeCellColor.Replace("#", "");
                }
                else
                {
                    colorHex = settings.DaysRowColor.Replace("#", "");
                }
            }
            else if (settings.ShowCellBackground)
            {
                colorHex = settings.NameCellColor.Replace("#", "");
            }

            if (colorHex != null)
            {
                var rgbColor = new A.RgbColorModelHex() { Val = colorHex };
                var alpha = new A.Alpha() { Val = (int)(settings.CellOpacity * 100000) };
                rgbColor.Append(alpha);
                solidFill.Append(rgbColor);
            }
            else
            {
                solidFill.Append(new A.NoFill());
            }

            shapeProperties.Append(solidFill);

            var outline = new A.Outline(
                new A.SolidFill(new A.RgbColorModelHex() { Val = "D1D5DB" }))
            {
                Width = isHeader ? 200 : 100
            };
            shapeProperties.Append(outline);

            var transform2D = new A.Transform2D();
            var offset = new A.Offset() { X = x, Y = y };
            var extents = new A.Extents() { Cx = width, Cy = height };

            transform2D.Append(offset);
            transform2D.Append(extents);
            shapeProperties.Append(transform2D);
            shapeProperties.Append(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

            var textBody = new P.TextBody();
            var bodyProperties = new A.BodyProperties()
            {
                Anchor = A.TextAnchoringTypeValues.Center,
                Wrap = A.TextWrappingValues.Square,
                AnchorCenter = true
            };

            var listStyle = new A.ListStyle();
            var paragraph = new A.Paragraph();

            var paragraphProperties = new A.ParagraphProperties()
            {
                Alignment = A.TextAlignmentTypeValues.Center
            };

            var runProperties = new A.RunProperties()
            {
                FontSize = (int)((isHeader ? settings.HeaderFontSize : settings.FontSize) * 100),
                Bold = isHeader
            };

            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex()
            {
                Val = settings.TextColor.Replace("#", "")
            }));

            var textRun = new A.Run();
            textRun.Append(runProperties);

            if (!string.IsNullOrEmpty(text))
            {
                textRun.Append(new A.Text(text));
            }
            else
            {
                textRun.Append(new A.Text(""));
            }

            paragraph.Append(paragraphProperties);
            paragraph.Append(textRun);

            textBody.Append(bodyProperties);
            textBody.Append(listStyle);
            textBody.Append(paragraph);

            shape.Append(nonVisualProperties);
            shape.Append(shapeProperties);
            shape.Append(textBody);

            shapeTree.Append(shape);
        }

        private static uint _shapeIdCounter = 10000;
        private static uint GetNextShapeId() => _shapeIdCounter++;

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
                    var backgroundColor = ColorConverter.ConvertFromString(settings.BackgroundColor);
                    if (backgroundColor is Color bgColor)
                    {
                        var backgroundBrush = new SolidColorBrush(bgColor);
                        drawingContext.DrawRectangle(backgroundBrush, null, new System.Windows.Rect(0, 0, 800, 450));
                    }

                    var titleText = new FormattedText(
                        $"Week of {weekRange[0]:MMM dd} - {weekRange[1]:MMM dd, yyyy}",
                        System.Globalization.CultureInfo.CurrentCulture,
                        System.Windows.FlowDirection.LeftToRight,
                        new Typeface("Arial"),
                        settings.FontSize,
                        new SolidColorBrush((Color)ColorConverter.ConvertFromString(settings.TextColor)),
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

                    var timeHeaderColor = ColorConverter.ConvertFromString(settings.TimeCellColor);
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
                        settings.HeaderFontSize * 0.7,
                        new SolidColorBrush((Color)ColorConverter.ConvertFromString(settings.TextColor)),
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

                        var daysRowColor = ColorConverter.ConvertFromString(settings.DaysRowColor);
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
                            settings.HeaderFontSize * 0.6,
                            new SolidColorBrush((Color)ColorConverter.ConvertFromString(settings.TextColor)),
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

                        var timeCellColor = ColorConverter.ConvertFromString(settings.TimeCellColor);
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
                            settings.HeaderFontSize * 0.5,
                            new SolidColorBrush((Color)ColorConverter.ConvertFromString(settings.TextColor)),
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
                            if (settings.UseColorCoding)
                            {
                                var cellColorHex = GetCellColor(names.Count, schedule.PeoplePerShift);
                                var cellColor = ColorConverter.ConvertFromString(cellColorHex);
                                if (cellColor is Color cColor)
                                {
                                    cellBrush = new SolidColorBrush(cColor);
                                }
                                else
                                {
                                    cellBrush = names.Count > 0
                                        ? new SolidColorBrush(Color.FromRgb(220, 252, 231))
                                        : new SolidColorBrush(Color.FromRgb(254, 226, 226));
                                }
                            }
                            else
                            {
                                var nameCellColor = ColorConverter.ConvertFromString(settings.NameCellColor);
                                if (nameCellColor is Color ncColor)
                                {
                                    cellBrush = new SolidColorBrush(ncColor);
                                }
                                else
                                {
                                    cellBrush = new SolidColorBrush(Colors.White);
                                }
                            }

                            drawingContext.DrawRectangle(cellBrush, new Pen(Brushes.Gray, 0.5),
                                new System.Windows.Rect(x, y, cellWidth, cellHeight));

                            if (names.Count > 0)
                            {
                                var nameText = new FormattedText(
                                    string.Join("\n", names.Take(2)),
                                    System.Globalization.CultureInfo.CurrentCulture,
                                    System.Windows.FlowDirection.LeftToRight,
                                    new Typeface("Arial"),
                                    settings.FontSize * 0.5,
                                    new SolidColorBrush((Color)ColorConverter.ConvertFromString(settings.TextColor)),
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
            catch (Exception ex)
            {
                Debug.WriteLine($"Error generating preview: {ex.Message}");
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
    }
}