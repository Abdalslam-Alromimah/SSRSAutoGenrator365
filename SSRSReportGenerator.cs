using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.IO;

namespace SSRSReportGenerator
{
    public class ColumnDefinition
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string DataField { get; set; }
        public string Width { get; set; }
        public int DisplayOrder { get; set; } 
        public bool IsVisible { get; set; } = true; 
        public bool IsGrouped { get; set; }
        public string Format { get; set; }
        public string FontFamily { get; set; }
    }

    public class GroupDefinition
    {
        public string GroupName { get; set; }
        public string ColumnName { get; set; }
        public string DisplayField { get; set; }
        public int DisplayOrder { get; set; } 

        public List<string> NestedColumns { get; set; } = new List<string>();
    }


    public class ReportGenerator
    {
        private readonly XNamespace ns = "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition";
        private readonly XNamespace rd = "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner";
        private readonly XNamespace df = "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily";
        private readonly XNamespace am = "http://schemas.microsoft.com/sqlserver/reporting/authoringmetadata";

        private readonly List<ColumnDefinition> _columns;
        private readonly List<GroupDefinition> _groups;
        private readonly string _dataSetName;

        public ReportGenerator(List<ColumnDefinition> columns, List<GroupDefinition> groups, string dataSetName)
        {
            _columns = columns;
            _groups = groups;
            _dataSetName = dataSetName;
        }

        public XDocument GenerateReport()
        {
            var doc = new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XElement(ns + "Report",
                    new XAttribute(XNamespace.Xmlns + "df", df),
                    new XAttribute(XNamespace.Xmlns + "rd", rd),
                    new XAttribute(XNamespace.Xmlns + "am", am),
                    new XAttribute("MustUnderstand", "df"),
                    CreateAuthoringMetadata(),
                    new XElement(df + "DefaultFontFamily", "Segoe UI"),
                    new XElement(ns + "AutoRefresh", "0"),
                    CreateDataSources(),
                    CreateDataSets(),
                    CreateReportSections(),
                    new XElement(ns + "ReportParametersLayout",
                        new XElement(ns + "GridLayoutDefinition",
                            new XElement(ns + "NumberOfColumns", "4"),
                            new XElement(ns + "NumberOfRows", "2")
                        )
                    ),
                    new XElement(rd + "ReportUnitType", "Inch"),
                    new XElement(rd + "ReportID", "30199e74-3fe5-461f-b6b6-f68d97e172df")
                )
            );
            return doc;
        }

        private XElement CreateAuthoringMetadata()
        {
            return new XElement(am + "AuthoringMetadata",
                new XElement(am + "CreatedBy",
                    new XElement(am + "Name", "xxxx"),
                    new XElement(am + "Version", "15.0.20283.0")
                ),
                new XElement(am + "UpdatedBy",
                    new XElement(am + "Name", "xxx"),
                    new XElement(am + "Version", "15.0.20283.0")
                ),
                new XElement(am + "LastModifiedTimestamp", "2025-01-30T19:18:16.1397756Z")
            );
        }

        private XElement CreateDataSources()
        {
            return new XElement(ns + "DataSources",
                new XElement(ns + "DataSource",
                    new XAttribute("Name", "DataSGetEmployeeReport"),
                    new XElement(ns + "ConnectionProperties",
                        new XElement(ns + "DataProvider", "SQL"),
                        new XElement(ns + "ConnectString", "Data Source=xx.xxx.xxx.221;Initial Catalog=zxxxxx"),
                        new XElement(ns + "Prompt", "Specify a user name and password for data source DataSource1:")
                    ),
                    new XElement(rd + "SecurityType", "DataBase"),
                    new XElement(rd + "DataSourceID", "da61d9fe-c960-4375-9c49-2992e9f9d4fd")
                )
            );
        }

        private XElement CreateDataSets()
        {
            return new XElement(ns + "DataSets",
                new XElement(ns + "DataSet",
                    new XAttribute("Name", _dataSetName),
                    new XElement(ns + "Query",
                        new XElement(ns + "DataSourceName", "DataSGetEmployeeReport"),
                        new XElement(ns + "CommandText", GetQueryText()),
                        new XElement(rd + "UseGenericDesigner", "true")
                    ),
                    CreateDataSetFields()
                )
            );
        }

        private XElement CreateDataSetFields()
        {
            var fields = new XElement(ns + "Fields");
            foreach (var col in _columns)
            {
                fields.Add(new XElement(ns + "Field",
                    new XAttribute("Name", col.DataField),
                    new XElement(ns + "DataField", col.DataField),
                    new XElement(rd + "TypeName", GetFieldType(col.DataField))
                ));
            }
            return fields;
        }

        private XElement CreateReportSections()
        {
            return new XElement(ns + "ReportSections",
                new XElement(ns + "ReportSection",
                    new XElement(ns + "Body",
                        new XElement(ns + "ReportItems",
                            CreateTitleTextbox(),
                            CreateMainTablix()
                        ),
                        new XElement(ns + "Height", "2.25in"),
                        new XElement(ns + "Style",
                            new XElement(ns + "Border",
                                new XElement(ns + "Style", "None")
                            )
                        )
                    ),
                    new XElement(ns + "Width", "7.08556in"),
                    CreatePageSection()
                )
            );
        }

        private XElement CreateTitleTextbox()
        {
            return new XElement(ns + "Textbox",
                new XAttribute("Name", "ReportTitle"),
                new XElement(ns + "CanGrow", "true"),
                new XElement(ns + "KeepTogether", "true"),
                new XElement(ns + "Paragraphs",
                    new XElement(ns + "Paragraph",
                        new XElement(ns + "TextRuns",
                            new XElement(ns + "TextRun",
                                new XElement(ns + "Value", "Employee Salary by Date"),
                                new XElement(ns + "Style",
                                    new XElement(ns + "FontFamily", "Segoe UI Light"),
                                    new XElement(ns + "FontSize", "14pt")
                                )
                            )
                        )
                    )
                ),
                new XElement(rd + "WatermarkTextbox", "Title"),
                new XElement(rd + "DefaultName", "ReportTitle"),
                new XElement(ns + "Top", "0.175in"),
                new XElement(ns + "Left", "0.625in"),
                new XElement(ns + "Height", "0.325in"),
                new XElement(ns + "Width", "5.5in"),
                CreateCommonStyle()
            );
        }

        private XElement CreateMainTablix()
        {
            return new XElement(ns + "Tablix",
                new XAttribute("Name", "Tablix1"),
                CreateTablixBody(),
                CreateTablixColumnHierarchy(ns),
                CreateTablixRowHierarchy(ns),
                new XElement(ns + "DataSetName", _dataSetName),
                new XElement(ns + "Top", "0.91194in"),
                new XElement(ns + "Left", "0.37722in"),
                new XElement(ns + "Height", "0.99167in"),
                new XElement(ns + "Width", "6.45834in"),
                new XElement(ns + "ZIndex", "1"),
                new XElement(ns + "Style",
                    new XElement(ns + "Border",
                        new XElement(ns + "Style", "None")
                    )
                )
            );
        }

        private XElement CreateTablixBody()
        {
            return new XElement(ns + "TablixBody",
                CreateTablixColumns(),
                CreateTablixRows()
            );
        }

        private XElement CreateTablixRowHierarchy(XNamespace ns)
        {
            return new XElement(ns + "TablixRowHierarchy",
                new XElement(ns + "TablixMembers",
                    new XElement(ns + "TablixMember",
                        new XElement(ns + "KeepWithGroup", "After")
                    ),
                    new XElement(ns + "TablixMember",
                        new XElement(ns + "Group", new XAttribute("Name", "Details"))
                    )
                )
            );
        }
        private XElement CreateTablixColumns()
        {
            var tablixColumns = new XElement(ns + "TablixColumns");

            // Get all visible columns in display order
            var orderedColumns = _columns
                .Where(c => !c.IsGrouped && c.IsVisible)
                .OrderBy(c => c.DisplayOrder);

            // Add regular columns
            foreach (var col in orderedColumns)
            {
                tablixColumns.Add(new XElement(ns + "TablixColumn",
                    new XElement(ns + "Width", col.Width)
                ));
            }

            // Add group columns in order
            foreach (var group in _groups.OrderBy(g => g.DisplayOrder))
            {
                foreach (var nested in group.NestedColumns)
                {
                    var nestedCol = _columns.First(c => c.DataField == nested);
                    if (nestedCol.IsVisible)
                    {
                        tablixColumns.Add(new XElement(ns + "TablixColumn",
                            new XElement(ns + "Width", nestedCol.Width)
                        ));
                    }
                }
            }

            return tablixColumns;
        }

        private XElement CreateTablixColumnHierarchy(XNamespace ns)
        {
            var members = new XElement(ns + "TablixMembers");

            // Static columns
            foreach (var col in _columns.Where(c => !c.IsGrouped))
            {
                members.Add(new XElement(ns + "TablixMember"));
            }

            // Group columns with nested members
            foreach (var group in _groups)
            {
                var groupMember = new XElement(ns + "TablixMember",
                    new XElement(ns + "Group",
                        new XAttribute("Name", group.GroupName),
                        new XElement(ns + "GroupExpressions",
                            new XElement(ns + "GroupExpression", $"=Fields!{group.ColumnName}.Value")
                        )
                    ),
                    new XElement(ns + "SortExpressions",
                        new XElement(ns + "SortExpression",
                            new XElement(ns + "Value", $"=Fields!{group.ColumnName}.Value")
                        )
                    ),
                    new XElement(ns + "TablixMembers",
                        group.NestedColumns.Select(nested =>
                            new XElement(ns + "TablixMember")
                        )
                    )
                );
                members.Add(groupMember);
            }

            return new XElement(ns + "TablixColumnHierarchy", members);
        }
        private XElement CreateTablixRows()
        {
            return new XElement(ns + "TablixRows",
                CreateHeaderRow(ns),
                CreateDetailsRow(ns)
            );
        }

        private XElement CreateHeaderRow(XNamespace ns)
        {
            var cells = new XElement(ns + "TablixCells");

            // Add headers for visible static columns in display order
            foreach (var col in _columns.Where(c => !c.IsGrouped && c.IsVisible).OrderBy(c => c.DisplayOrder))
            {
                cells.Add(CreateHeaderCell(col));
            }

            // Add group headers in display order
            foreach (var group in _groups.OrderBy(g => g.DisplayOrder))
            {
                cells.Add(new XElement(ns + "TablixCell",
                    new XElement(ns + "CellContents",
                        CreateGroupHeaderTextbox(ns, group.ColumnName)
                    )
                ));
            }

            return new XElement(ns + "TablixRow",
                new XElement(ns + "Height", "0.37084in"),
                cells
            );
        }

        private XElement CreateGroupDataCell(GroupDefinition group, string fieldName)
        {
            string uniqueName = $"Data_{group.GroupName}_{fieldName}";
            return new XElement(ns + "TablixCell",
                new XElement(ns + "CellContents",
                    CreateTextbox($"=Fields!{fieldName}.Value", uniqueName, "Arial") // Add fontFamily here
                )
            );
        }
        private XElement CreateHeaderCell(ColumnDefinition col)
        {
            string uniqueName = $"Header_{col.Name}";
            return new XElement(ns + "TablixCell",
                new XElement(ns + "CellContents",
                    CreateTextbox(col.DisplayName, uniqueName, col.FontFamily ?? "Arial") // Use Arial if FontFamily is null
                )
            );
        }
        private XElement CreateGroupHeaderTextbox(XNamespace ns, string columnName, int colspan = 1)
        {
            string uniqueName = $"Header_{columnName}";
            return new XElement(ns + "Textbox",
                new XAttribute("Name", uniqueName),
                new XElement(ns + "CanGrow", "true"),
                new XElement(ns + "KeepTogether", "true"),
                new XElement(ns + "Paragraphs",
                    new XElement(ns + "Paragraph",
                        new XElement(ns + "TextRuns",
                            new XElement(ns + "TextRun",
                        // Fields! expression instead of static column name
                        new XElement(ns + "Value", $"=Fields!{columnName}.Value"),
                        new XElement(ns + "Style",
                                    new XElement(ns + "FontFamily", "Arial"),
                                    new XElement(ns + "FontSize", "10pt"),
                                    new XElement(ns + "FontWeight", "Bold")
                                )
                            )
                        )
                    )
                ),
                new XElement(rd + "DefaultName", uniqueName),
                CreateCommonStyle()
            );
        }

        private XElement CreateDetailsRow(XNamespace ns)
        {
            var cells = new XElement(ns + "TablixCells");

            // Add data for visible static columns in display order
            foreach (var col in _columns.Where(c => !c.IsGrouped && c.IsVisible).OrderBy(c => c.DisplayOrder))
            {
                cells.Add(CreateDataCell(col));
            }

            // Add group data in display order
            foreach (var group in _groups.OrderBy(g => g.DisplayOrder))
            {
                cells.Add(new XElement(ns + "TablixCell",
                    new XElement(ns + "CellContents",
                        CreateTextbox(
                            $"=Fields!{group.DisplayField ?? group.ColumnName}.Value",
                            $"Group_{group.GroupName}",
                            "Arial",
                            null
                        )
                    )
                ));

                foreach (var nested in group.NestedColumns)
                {
                    var nestedCol = _columns.First(c => c.DataField == nested);
                    if (nestedCol.IsVisible)
                    {
                        cells.Add(CreateDataCell(nestedCol));
                    }
                }
            }

            return new XElement(ns + "TablixRow",
                new XElement(ns + "Height", "0.37084in"),
                cells
            );
        }
    
       
        private XElement CreateTextbox(string value, string name, string fontFamily, string format = null)
        {
            var styleElements = new List<XElement>();
            if (!string.IsNullOrEmpty(fontFamily))
            {
                styleElements.Add(new XElement(ns + "FontFamily", fontFamily));
            }
            if (!string.IsNullOrEmpty(format))
            {
                styleElements.Add(new XElement(ns + "Format", format));
            }

            var textRun = new XElement(ns + "TextRun",
                new XElement(ns + "Value", value)
            );

            if (styleElements.Count > 0)
            {
                textRun.Add(new XElement(ns + "Style", styleElements));
            }

            var textbox = new XElement(ns + "Textbox",
                new XAttribute("Name", name),
                new XElement(ns + "CanGrow", "true"),
                new XElement(ns + "KeepTogether", "true"),
                new XElement(ns + "Paragraphs",
                    new XElement(ns + "Paragraph",
                        new XElement(ns + "TextRuns", textRun)
                    )
                ),
                new XElement(rd + "DefaultName", name),
                CreateCommonStyle()
            );

            return textbox;
        }
        private XElement CreateDataCell(ColumnDefinition col)
        {
            return new XElement(ns + "TablixCell",
                new XElement(ns + "CellContents",
                    CreateTextbox($"=Fields!{col.DataField}.Value", $"Data_{col.Name}", col.FontFamily, col.Format)
                )
            );
        }


        private XElement CreateCommonStyle()
        {
            return new XElement(ns + "Style",
                new XElement(ns + "Border",
                    new XElement(ns + "Color", "LightGrey"),
                    new XElement(ns + "Style", "Solid")
                ),
                new XElement(ns + "PaddingLeft", "2pt"),
                new XElement(ns + "PaddingRight", "2pt"),
                new XElement(ns + "PaddingTop", "2pt"),
                new XElement(ns + "PaddingBottom", "2pt")
            );
        }

        private XElement CreatePageSection()
        {
            return new XElement(ns + "Page",
                new XElement(ns + "PageFooter",
                    new XElement(ns + "Height", "0.45in"),
                    new XElement(ns + "PrintOnFirstPage", "true"),
                    new XElement(ns + "PrintOnLastPage", "true"),
                    new XElement(ns + "Style",
                        new XElement(ns + "Border",
                            new XElement(ns + "Style", "None")
                        )
                    )
                ),
                new XElement(ns + "LeftMargin", "1in"),
                new XElement(ns + "RightMargin", "1in"),
                new XElement(ns + "TopMargin", "1in"),
                new XElement(ns + "BottomMargin", "1in")
            );
        }

        private string GetFieldType(string fieldName)
        {
            return fieldName switch
            {
                "Salary" => "System.Decimal",
                "Date" => "System.DateTime",
                _ => "System.String"
            };
        }

        private string GetQueryText()
        {
            return @"SELECT [LastName],[Department],[Position],[Salary],[HireDate],[FirstName] FROM EmployeeData365";
        }
    }


    class Program
    {
        static void Main(string[] args)
        {
            // User 1
            void GenerateUser1Report()
            {
                var columns = new List<ColumnDefinition>
                {
                    new ColumnDefinition { Name = "FirstName", DisplayName = "First Name", DataField = "FirstName", Width = "1.15278in", DisplayOrder = 1 },
                    new ColumnDefinition { Name = "LastName", DisplayName = "Last Name", DataField = "LastName", Width = "1.15278in", DisplayOrder = 2 },
                    new ColumnDefinition { Name = "Department", DisplayName = "Department", DataField = "Department", Width = "1.15278in", DisplayOrder = 3 },
                    new ColumnDefinition { Name = "Position", DisplayName = "Position", DataField = "Position", Width = "1in", FontFamily = "Arial", DisplayOrder = 4 },
                    new ColumnDefinition { Name = "Salary", DisplayName = "Salary", DataField = "Salary", Width = "1in", FontFamily = "Arial", Format = "C2", DisplayOrder = 5 },
                    new ColumnDefinition { Name = "HireDate", DisplayName = "Hire Date", DataField = "HireDate", Width = "1in", FontFamily = "Arial", IsGrouped = true, DisplayOrder = 6 }
                };

                var groups = new List<GroupDefinition>
                {
                    new GroupDefinition
                    {
                        GroupName = "DateGroup",
                        ColumnName = "HireDate",
                        DisplayField = "Salary",
                        NestedColumns = { "Salary" },
                        DisplayOrder = 1
                    }
                };

                GenerateReport(columns, groups, "User1Report.rdl");
            }

            // for User 2
            void GenerateUser2Report()
            {
                var columns = new List<ColumnDefinition>
                {
                    new ColumnDefinition { Name = "FirstName", DisplayName = "First Name", DataField = "FirstName", Width = "1.15278in", DisplayOrder = 2 },
                    new ColumnDefinition { Name = "LastName", DisplayName = "Last Name", DataField = "LastName", Width = "1.15278in", DisplayOrder = 1 },
                    new ColumnDefinition { Name = "Department", DisplayName = "Department", DataField = "Department", Width = "1.15278in", DisplayOrder = 3 },
                    new ColumnDefinition { Name = "Position", DisplayName = "Position", DataField = "Position", Width = "1in", FontFamily = "Arial", DisplayOrder = 4 },
                    new ColumnDefinition { Name = "Salary", DisplayName = "Salary", DataField = "Salary", Width = "1in", FontFamily = "Arial", Format = "C2", DisplayOrder = 5 },
                    new ColumnDefinition { Name = "HireDate", DisplayName = "Hire Date", DataField = "HireDate", Width = "1in", FontFamily = "Arial", IsGrouped = true, DisplayOrder = 6 }
                };

                var groups = new List<GroupDefinition>
                {
                    new GroupDefinition
                    {
                        GroupName = "DateGroup",
                        ColumnName = "HireDate",
                        DisplayField = "Salary",
                        NestedColumns = { "Salary" },
                        DisplayOrder = 1
                    }
                };

                GenerateReport(columns, groups, "User2Report.rdl");
            }

            // User 3 - with HireDate before Salary
            void GenerateUser3Report()
            {
                var columns = new List<ColumnDefinition>
                {
                    new ColumnDefinition { Name = "LastName", DisplayName = "Last Name", DataField = "LastName", Width = "1.15278in", DisplayOrder = 1 },
                    new ColumnDefinition { Name = "FirstName", DisplayName = "First Name", DataField = "FirstName", Width = "1.15278in", DisplayOrder = 2 },
                    new ColumnDefinition { Name = "Department", DisplayName = "Department", DataField = "Department", Width = "1.15278in", DisplayOrder = 3 },
                    new ColumnDefinition { Name = "Position", DisplayName = "Position", DataField = "Position", Width = "1in", FontFamily = "Arial", DisplayOrder = 4 },
                    new ColumnDefinition { Name = "HireDate", DisplayName = "Hire Date", DataField = "HireDate", Width = "1in", FontFamily = "Arial", IsGrouped = true, DisplayOrder = 5 },
                    new ColumnDefinition { Name = "Salary", DisplayName = "Salary", DataField = "Salary", Width = "1in", FontFamily = "Arial", Format = "C2", DisplayOrder = 6 }
                };

                var groups = new List<GroupDefinition>
                {
                    new GroupDefinition
                    {
                        GroupName = "DateGroup",
                        ColumnName = "HireDate",
                        DisplayField = "HireDate",
                        NestedColumns = { "Salary" },
                        DisplayOrder = 1
                    }
                };

                GenerateReport(columns, groups, "User3Report.rdl");
            }

            //  User 4 - fewer columns
            void GenerateUser4Report()
            {
                var columns = new List<ColumnDefinition>
                {
                    new ColumnDefinition { Name = "LastName", DisplayName = "Last Name", DataField = "LastName", Width = "1.15278in", DisplayOrder = 1 },
                    new ColumnDefinition { Name = "FirstName", DisplayName = "First Name", DataField = "FirstName", Width = "1.15278in", DisplayOrder = 2 },
                    new ColumnDefinition { Name = "Department", DisplayName = "Department", DataField = "Department", Width = "1.15278in", DisplayOrder = 3 },
                    new ColumnDefinition { Name = "HireDate", DisplayName = "Hire Date", DataField = "HireDate", Width = "1in", FontFamily = "Arial", IsGrouped = true, DisplayOrder = 4 },
                    new ColumnDefinition { Name = "Salary", DisplayName = "Salary", DataField = "Salary", Width = "1in", FontFamily = "Arial", Format = "C2", DisplayOrder = 5 },
                    new ColumnDefinition { Name = "Position", DisplayName = "Position", DataField = "Position", Width = "1in", FontFamily = "Arial", IsVisible = false }
                };

                var groups = new List<GroupDefinition>
                {
                    new GroupDefinition
                    {
                        GroupName = "DateGroup",
                        ColumnName = "HireDate",
                        DisplayField = "HireDate",
                        NestedColumns = { "Salary" },
                        DisplayOrder = 1
                    }
                };

                GenerateReport(columns, groups, "User4Report.rdl");
            }

            static void GenerateReport(List<ColumnDefinition> columns, List<GroupDefinition> groups, string fileName)
            {
                var generator = new ReportGenerator(columns, groups, "DataSetGetEmployeeReport");
                XDocument report = generator.GenerateReport();

                string filePath = Path.Combine(@"C:\Reports", fileName);
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                report.Save(filePath);

                Console.WriteLine($"Report generated successfully at: {filePath}");
            }

            // Generate all reports
            GenerateUser1Report();
            GenerateUser2Report();
            GenerateUser3Report();
            GenerateUser4Report();
        }
    }

}