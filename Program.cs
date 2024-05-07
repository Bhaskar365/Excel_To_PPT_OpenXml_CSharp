
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using System;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;


using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

using OfficeOpenXml;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

string[,] excelData;
string excelPath;
string filepath;

CreatePresentation();
void CreatePresentation()
{
    excelPath = "C:\\Testing\\Template_Creation\\Test_Excel_To_PPT\\NewFolder\\sample.xlsx";
    filepath = "C:\\Testing\\Template_Creation\\Test_Excel_To_PPT\\NewFolder\\sample.pptx";

    excelData = ReadExcelData(excelPath);

    // Create a presentation at a specified file path. The presentation document type is pptx, by default.
    PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
    PresentationPart presentationPart = presentationDoc.AddPresentationPart();

    presentationPart.Presentation = new Presentation();

    CreatePresentationParts(presentationPart);

    //Dispose the presentation handle
    presentationDoc.Dispose();
}

void CreatePresentationParts(PresentationPart presentationPart)
{
    SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
    SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
    SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
    NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
    DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

    presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

    SlidePart slidePart1;
    SlideLayoutPart slideLayoutPart1;
    SlideMasterPart slideMasterPart1;
    ThemePart themePart1;

    slidePart1 = CreateSlidePart(presentationPart);

    slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
    slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
    themePart1 = CreateTheme(slideMasterPart1);

    ChartPart chartPart1 = slidePart1.AddNewPart<ChartPart>("rId3");
    GenerateChartPart2Content(chartPart1);

    excelData = ReadExcelData(excelPath);

    // EmbedExcelDataIntoChartPart(chartPart1, excelDataInString);

    EmbeddedPackagePart embeddedPackagePart1 = chartPart1.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");
    GenerateEmbeddedPackagePart1Content(embeddedPackagePart1);

    ChartColorStylePart chartColorStylePart1 = chartPart1.AddNewPart<ChartColorStylePart>("rId2");
    GenerateChartColorStylePart1Content(chartColorStylePart1);

    ChartStylePart chartStylePart1 = chartPart1.AddNewPart<ChartStylePart>("rId1");
    //GenerateChartStylePart1Content(chartStylePart1);

    slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
    presentationPart.AddPart(slideMasterPart1, "rId1");
    presentationPart.AddPart(themePart1, "rId5");
}

string[,] ReadExcelData(string excelFilePath)
{
    try
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; // Ensure EPPlus runs without license issues

        string[,] data;

        using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            int columnCount = worksheet.Dimension.Columns;


            // Create an array to hold the Excel data
            data = new string[rowCount, columnCount];
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    data[row - 1, col - 1] = worksheet.Cells[row, col].Text;
                    Console.WriteLine($"Cell[{row},{col}]: {data[row - 1, col - 1]}");
                }
            }
        }
        return data;
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error reading Excel data: " + ex.Message);
        return null;
    }
}

void GenerateChartPart2Content(ChartPart chartPart)
{
    // Create a new instance of ChartPart content
    C.ChartSpace chartSpace = new C.ChartSpace();
    chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
    chartPart.ChartSpace = chartSpace;

    // Create the chart
    C.Chart chart = new C.Chart();
    C.PlotArea plotArea = new C.PlotArea();

    // Create a bar chart
    BarChart barChart = new BarChart();
    barChart.Append(new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) });

    for (int col = 1; col < excelData.GetLength(1); col++) // Skip the first column which contains categories
    {
        BarChartSeries barChartSeries = new BarChartSeries();

        // Set series data (categories)
        CategoryAxisData categoryAxisData = new CategoryAxisData();
        StringReference stringReference = new StringReference();
        C.Formula formula = new C.Formula();
        formula.Text = $"Sheet1!${Convert.ToChar(65 + col)}$1:${Convert.ToChar(65 + col)}{excelData.GetLength(0) + 1}"; // Assuming data starts from A1
        stringReference.Append(formula);
        categoryAxisData.Append(stringReference);
        barChartSeries.Append(categoryAxisData);

        // Set series values
        D.Charts.Values values = new D.Charts.Values();
        NumberReference numberReference = new NumberReference();
        C.Formula formula2 = new C.Formula();
        formula2.Text = $"Sheet1!${Convert.ToChar(65 + col)}$2:${Convert.ToChar(65 + col)}{excelData.GetLength(0) + 1}"; // Assuming data starts from A2
        numberReference.Append(formula2);
        values.Append(numberReference);
        barChartSeries.Append(values);

        barChart.Append(barChartSeries);
    }

    plotArea.Append(barChart);
    chart.Append(plotArea);
    chartSpace.Append(chart);
}

//void GenerateChartStylePart1Content(ChartStylePart chartStylePart1)
//{
//    Cs.ChartStyle chartStyle1 = new Cs.ChartStyle() { Id = (UInt32Value)297U };
//    chartStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
//    chartStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

//    Cs.AxisTitle axisTitle1 = new Cs.AxisTitle();
//    Cs.LineReference lineReference1 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference1 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference1 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference1 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor25.Append(luminanceModulation15);
//    schemeColor25.Append(luminanceOffset11);

//    fontReference1.Append(schemeColor25);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType1 = new Cs.TextCharacterPropertiesType() { FontSize = 1330, Kerning = 1200 };

//    axisTitle1.Append(lineReference1);
//    axisTitle1.Append(fillReference1);
//    axisTitle1.Append(effectReference1);
//    axisTitle1.Append(fontReference1);
//    axisTitle1.Append(textCharacterPropertiesType1);

//    Cs.CategoryAxis categoryAxis2 = new Cs.CategoryAxis();
//    Cs.LineReference lineReference2 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference2 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference2 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference2 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation16 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset12 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor26.Append(luminanceModulation16);
//    schemeColor26.Append(luminanceOffset12);

//    fontReference2.Append(schemeColor26);

//    Cs.ShapeProperties shapeProperties3 = new Cs.ShapeProperties();

//    A.Outline outline11 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill19 = new A.SolidFill();

//    A.SchemeColor schemeColor27 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation17 = new A.LuminanceModulation() { Val = 15000 };
//    A.LuminanceOffset luminanceOffset13 = new A.LuminanceOffset() { Val = 85000 };

//    schemeColor27.Append(luminanceModulation17);
//    schemeColor27.Append(luminanceOffset13);

//    solidFill19.Append(schemeColor27);
//    A.Round round3 = new A.Round();

//    outline11.Append(solidFill19);
//    outline11.Append(round3);

//    shapeProperties3.Append(outline11);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType2 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    categoryAxis2.Append(lineReference2);
//    categoryAxis2.Append(fillReference2);
//    categoryAxis2.Append(effectReference2);
//    categoryAxis2.Append(fontReference2);
//    categoryAxis2.Append(shapeProperties3);
//    categoryAxis2.Append(textCharacterPropertiesType2);

//    Cs.ChartArea chartArea1 = new Cs.ChartArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
//    Cs.LineReference lineReference3 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference3 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference3 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference3 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor28 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference3.Append(schemeColor28);

//    Cs.ShapeProperties shapeProperties4 = new Cs.ShapeProperties();

//    A.SolidFill solidFill20 = new A.SolidFill();
//    A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

//    solidFill20.Append(schemeColor29);

//    A.Outline outline12 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill21 = new A.SolidFill();

//    A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation18 = new A.LuminanceModulation() { Val = 15000 };
//    A.LuminanceOffset luminanceOffset14 = new A.LuminanceOffset() { Val = 85000 };

//    schemeColor30.Append(luminanceModulation18);
//    schemeColor30.Append(luminanceOffset14);

//    solidFill21.Append(schemeColor30);
//    A.Round round4 = new A.Round();

//    outline12.Append(solidFill21);
//    outline12.Append(round4);

//    shapeProperties4.Append(solidFill20);
//    shapeProperties4.Append(outline12);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType3 = new Cs.TextCharacterPropertiesType() { FontSize = 1330, Kerning = 1200 };

//    chartArea1.Append(lineReference3);
//    chartArea1.Append(fillReference3);
//    chartArea1.Append(effectReference3);
//    chartArea1.Append(fontReference3);
//    chartArea1.Append(shapeProperties4);
//    chartArea1.Append(textCharacterPropertiesType3);

//    Cs.DataLabel dataLabel1 = new Cs.DataLabel();
//    Cs.LineReference lineReference4 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference4 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference4 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference4 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 75000 };
//    A.LuminanceOffset luminanceOffset15 = new A.LuminanceOffset() { Val = 25000 };

//    schemeColor31.Append(luminanceModulation19);
//    schemeColor31.Append(luminanceOffset15);

//    fontReference4.Append(schemeColor31);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType4 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    dataLabel1.Append(lineReference4);
//    dataLabel1.Append(fillReference4);
//    dataLabel1.Append(effectReference4);
//    dataLabel1.Append(fontReference4);
//    dataLabel1.Append(textCharacterPropertiesType4);

//    Cs.DataLabelCallout dataLabelCallout1 = new Cs.DataLabelCallout();
//    Cs.LineReference lineReference5 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference5 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference5 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference5 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor32 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
//    A.LuminanceModulation luminanceModulation20 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset16 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor32.Append(luminanceModulation20);
//    schemeColor32.Append(luminanceOffset16);

//    fontReference5.Append(schemeColor32);

//    Cs.ShapeProperties shapeProperties5 = new Cs.ShapeProperties();

//    A.SolidFill solidFill22 = new A.SolidFill();
//    A.SchemeColor schemeColor33 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

//    solidFill22.Append(schemeColor33);

//    A.Outline outline13 = new A.Outline();

//    A.SolidFill solidFill23 = new A.SolidFill();

//    A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
//    A.LuminanceModulation luminanceModulation21 = new A.LuminanceModulation() { Val = 25000 };
//    A.LuminanceOffset luminanceOffset17 = new A.LuminanceOffset() { Val = 75000 };

//    schemeColor34.Append(luminanceModulation21);
//    schemeColor34.Append(luminanceOffset17);

//    solidFill23.Append(schemeColor34);

//    outline13.Append(solidFill23);

//    shapeProperties5.Append(solidFill22);
//    shapeProperties5.Append(outline13);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType5 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    Cs.TextBodyProperties textBodyProperties1 = new Cs.TextBodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 36576, TopInset = 18288, RightInset = 36576, BottomInset = 18288, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
//    A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

//    textBodyProperties1.Append(shapeAutoFit1);

//    dataLabelCallout1.Append(lineReference5);
//    dataLabelCallout1.Append(fillReference5);
//    dataLabelCallout1.Append(effectReference5);
//    dataLabelCallout1.Append(fontReference5);
//    dataLabelCallout1.Append(shapeProperties5);
//    dataLabelCallout1.Append(textCharacterPropertiesType5);
//    dataLabelCallout1.Append(textBodyProperties1);

//    Cs.DataPoint dataPoint1 = new Cs.DataPoint();
//    Cs.LineReference lineReference6 = new Cs.LineReference() { Index = (UInt32Value)0U };

//    Cs.FillReference fillReference6 = new Cs.FillReference() { Index = (UInt32Value)1U };
//    Cs.StyleColor styleColor1 = new Cs.StyleColor() { Val = "auto" };

//    fillReference6.Append(styleColor1);
//    Cs.EffectReference effectReference6 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference6 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference6.Append(schemeColor35);

//    dataPoint1.Append(lineReference6);
//    dataPoint1.Append(fillReference6);
//    dataPoint1.Append(effectReference6);
//    dataPoint1.Append(fontReference6);

//    Cs.DataPoint3D dataPoint3D1 = new Cs.DataPoint3D();
//    Cs.LineReference lineReference7 = new Cs.LineReference() { Index = (UInt32Value)0U };

//    Cs.FillReference fillReference7 = new Cs.FillReference() { Index = (UInt32Value)1U };
//    Cs.StyleColor styleColor2 = new Cs.StyleColor() { Val = "auto" };

//    fillReference7.Append(styleColor2);
//    Cs.EffectReference effectReference7 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference7 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference7.Append(schemeColor36);

//    dataPoint3D1.Append(lineReference7);
//    dataPoint3D1.Append(fillReference7);
//    dataPoint3D1.Append(effectReference7);
//    dataPoint3D1.Append(fontReference7);

//    Cs.DataPointLine dataPointLine1 = new Cs.DataPointLine();

//    Cs.LineReference lineReference8 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.StyleColor styleColor3 = new Cs.StyleColor() { Val = "auto" };

//    lineReference8.Append(styleColor3);
//    Cs.FillReference fillReference8 = new Cs.FillReference() { Index = (UInt32Value)1U };
//    Cs.EffectReference effectReference8 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference8 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference8.Append(schemeColor37);

//    Cs.ShapeProperties shapeProperties6 = new Cs.ShapeProperties();

//    A.Outline outline14 = new A.Outline() { Width = 28575, CapType = A.LineCapValues.Round };

//    A.SolidFill solidFill24 = new A.SolidFill();
//    A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

//    solidFill24.Append(schemeColor38);
//    A.Round round5 = new A.Round();

//    outline14.Append(solidFill24);
//    outline14.Append(round5);

//    shapeProperties6.Append(outline14);

//    dataPointLine1.Append(lineReference8);
//    dataPointLine1.Append(fillReference8);
//    dataPointLine1.Append(effectReference8);
//    dataPointLine1.Append(fontReference8);
//    dataPointLine1.Append(shapeProperties6);

//    Cs.DataPointMarker dataPointMarker1 = new Cs.DataPointMarker();

//    Cs.LineReference lineReference9 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.StyleColor styleColor4 = new Cs.StyleColor() { Val = "auto" };

//    lineReference9.Append(styleColor4);

//    Cs.FillReference fillReference9 = new Cs.FillReference() { Index = (UInt32Value)1U };
//    Cs.StyleColor styleColor5 = new Cs.StyleColor() { Val = "auto" };

//    fillReference9.Append(styleColor5);
//    Cs.EffectReference effectReference9 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference9 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference9.Append(schemeColor39);

//    Cs.ShapeProperties shapeProperties7 = new Cs.ShapeProperties();

//    A.Outline outline15 = new A.Outline() { Width = 9525 };

//    A.SolidFill solidFill25 = new A.SolidFill();
//    A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

//    solidFill25.Append(schemeColor40);

//    outline15.Append(solidFill25);

//    shapeProperties7.Append(outline15);

//    dataPointMarker1.Append(lineReference9);
//    dataPointMarker1.Append(fillReference9);
//    dataPointMarker1.Append(effectReference9);
//    dataPointMarker1.Append(fontReference9);
//    dataPointMarker1.Append(shapeProperties7);
//    Cs.MarkerLayoutProperties markerLayoutProperties1 = new Cs.MarkerLayoutProperties() { Symbol = Cs.MarkerStyle.Circle, Size = 5 };

//    Cs.DataPointWireframe dataPointWireframe1 = new Cs.DataPointWireframe();

//    Cs.LineReference lineReference10 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.StyleColor styleColor6 = new Cs.StyleColor() { Val = "auto" };

//    lineReference10.Append(styleColor6);
//    Cs.FillReference fillReference10 = new Cs.FillReference() { Index = (UInt32Value)1U };
//    Cs.EffectReference effectReference10 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference10 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference10.Append(schemeColor41);

//    Cs.ShapeProperties shapeProperties8 = new Cs.ShapeProperties();

//    A.Outline outline16 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Round };

//    A.SolidFill solidFill26 = new A.SolidFill();
//    A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

//    solidFill26.Append(schemeColor42);
//    A.Round round6 = new A.Round();

//    outline16.Append(solidFill26);
//    outline16.Append(round6);

//    shapeProperties8.Append(outline16);

//    dataPointWireframe1.Append(lineReference10);
//    dataPointWireframe1.Append(fillReference10);
//    dataPointWireframe1.Append(effectReference10);
//    dataPointWireframe1.Append(fontReference10);
//    dataPointWireframe1.Append(shapeProperties8);

//    Cs.DataTableStyle dataTableStyle1 = new Cs.DataTableStyle();
//    Cs.LineReference lineReference11 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference11 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference11 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference11 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation22 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset18 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor43.Append(luminanceModulation22);
//    schemeColor43.Append(luminanceOffset18);

//    fontReference11.Append(schemeColor43);

//    Cs.ShapeProperties shapeProperties9 = new Cs.ShapeProperties();
//    A.NoFill noFill15 = new A.NoFill();

//    A.Outline outline17 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill27 = new A.SolidFill();

//    A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation23 = new A.LuminanceModulation() { Val = 15000 };
//    A.LuminanceOffset luminanceOffset19 = new A.LuminanceOffset() { Val = 85000 };

//    schemeColor44.Append(luminanceModulation23);
//    schemeColor44.Append(luminanceOffset19);

//    solidFill27.Append(schemeColor44);
//    A.Round round7 = new A.Round();

//    outline17.Append(solidFill27);
//    outline17.Append(round7);

//    shapeProperties9.Append(noFill15);
//    shapeProperties9.Append(outline17);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType6 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    dataTableStyle1.Append(lineReference11);
//    dataTableStyle1.Append(fillReference11);
//    dataTableStyle1.Append(effectReference11);
//    dataTableStyle1.Append(fontReference11);
//    dataTableStyle1.Append(shapeProperties9);
//    dataTableStyle1.Append(textCharacterPropertiesType6);

//    Cs.DownBar downBar1 = new Cs.DownBar();
//    Cs.LineReference lineReference12 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference12 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference12 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference12 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference12.Append(schemeColor45);

//    Cs.ShapeProperties shapeProperties10 = new Cs.ShapeProperties();

//    A.SolidFill solidFill28 = new A.SolidFill();

//    A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
//    A.LuminanceModulation luminanceModulation24 = new A.LuminanceModulation() { Val = 75000 };
//    A.LuminanceOffset luminanceOffset20 = new A.LuminanceOffset() { Val = 25000 };

//    schemeColor46.Append(luminanceModulation24);
//    schemeColor46.Append(luminanceOffset20);

//    solidFill28.Append(schemeColor46);

//    A.Outline outline18 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill29 = new A.SolidFill();

//    A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation25 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset21 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor47.Append(luminanceModulation25);
//    schemeColor47.Append(luminanceOffset21);

//    solidFill29.Append(schemeColor47);
//    A.Round round8 = new A.Round();

//    outline18.Append(solidFill29);
//    outline18.Append(round8);

//    shapeProperties10.Append(solidFill28);
//    shapeProperties10.Append(outline18);

//    downBar1.Append(lineReference12);
//    downBar1.Append(fillReference12);
//    downBar1.Append(effectReference12);
//    downBar1.Append(fontReference12);
//    downBar1.Append(shapeProperties10);

//    Cs.DropLine dropLine1 = new Cs.DropLine();
//    Cs.LineReference lineReference13 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference13 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference13 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference13 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference13.Append(schemeColor48);

//    Cs.ShapeProperties shapeProperties11 = new Cs.ShapeProperties();

//    A.Outline outline19 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill30 = new A.SolidFill();

//    A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation26 = new A.LuminanceModulation() { Val = 35000 };
//    A.LuminanceOffset luminanceOffset22 = new A.LuminanceOffset() { Val = 65000 };

//    schemeColor49.Append(luminanceModulation26);
//    schemeColor49.Append(luminanceOffset22);

//    solidFill30.Append(schemeColor49);
//    A.Round round9 = new A.Round();

//    outline19.Append(solidFill30);
//    outline19.Append(round9);

//    shapeProperties11.Append(outline19);

//    dropLine1.Append(lineReference13);
//    dropLine1.Append(fillReference13);
//    dropLine1.Append(effectReference13);
//    dropLine1.Append(fontReference13);
//    dropLine1.Append(shapeProperties11);

//    Cs.ErrorBar errorBar1 = new Cs.ErrorBar();
//    Cs.LineReference lineReference14 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference14 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference14 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference14 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference14.Append(schemeColor50);

//    Cs.ShapeProperties shapeProperties12 = new Cs.ShapeProperties();

//    A.Outline outline20 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill31 = new A.SolidFill();

//    A.SchemeColor schemeColor51 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation27 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset23 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor51.Append(luminanceModulation27);
//    schemeColor51.Append(luminanceOffset23);

//    solidFill31.Append(schemeColor51);
//    A.Round round10 = new A.Round();

//    outline20.Append(solidFill31);
//    outline20.Append(round10);

//    shapeProperties12.Append(outline20);

//    errorBar1.Append(lineReference14);
//    errorBar1.Append(fillReference14);
//    errorBar1.Append(effectReference14);
//    errorBar1.Append(fontReference14);
//    errorBar1.Append(shapeProperties12);

//    Cs.Floor floor1 = new Cs.Floor();
//    Cs.LineReference lineReference15 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference15 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference15 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference15 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor52 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference15.Append(schemeColor52);

//    Cs.ShapeProperties shapeProperties13 = new Cs.ShapeProperties();
//    A.NoFill noFill16 = new A.NoFill();

//    A.Outline outline21 = new A.Outline();
//    A.NoFill noFill17 = new A.NoFill();

//    outline21.Append(noFill17);

//    shapeProperties13.Append(noFill16);
//    shapeProperties13.Append(outline21);

//    floor1.Append(lineReference15);
//    floor1.Append(fillReference15);
//    floor1.Append(effectReference15);
//    floor1.Append(fontReference15);
//    floor1.Append(shapeProperties13);

//    Cs.GridlineMajor gridlineMajor1 = new Cs.GridlineMajor();
//    Cs.LineReference lineReference16 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference16 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference16 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference16 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference16.Append(schemeColor53);

//    Cs.ShapeProperties shapeProperties14 = new Cs.ShapeProperties();

//    A.Outline outline22 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill32 = new A.SolidFill();

//    A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation28 = new A.LuminanceModulation() { Val = 15000 };
//    A.LuminanceOffset luminanceOffset24 = new A.LuminanceOffset() { Val = 85000 };

//    schemeColor54.Append(luminanceModulation28);
//    schemeColor54.Append(luminanceOffset24);

//    solidFill32.Append(schemeColor54);
//    A.Round round11 = new A.Round();

//    outline22.Append(solidFill32);
//    outline22.Append(round11);

//    shapeProperties14.Append(outline22);

//    gridlineMajor1.Append(lineReference16);
//    gridlineMajor1.Append(fillReference16);
//    gridlineMajor1.Append(effectReference16);
//    gridlineMajor1.Append(fontReference16);
//    gridlineMajor1.Append(shapeProperties14);

//    Cs.GridlineMinor gridlineMinor1 = new Cs.GridlineMinor();
//    Cs.LineReference lineReference17 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference17 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference17 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference17 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference17.Append(schemeColor55);

//    Cs.ShapeProperties shapeProperties15 = new Cs.ShapeProperties();

//    A.Outline outline23 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill33 = new A.SolidFill();

//    A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation29 = new A.LuminanceModulation() { Val = 5000 };
//    A.LuminanceOffset luminanceOffset25 = new A.LuminanceOffset() { Val = 95000 };

//    schemeColor56.Append(luminanceModulation29);
//    schemeColor56.Append(luminanceOffset25);

//    solidFill33.Append(schemeColor56);
//    A.Round round12 = new A.Round();

//    outline23.Append(solidFill33);
//    outline23.Append(round12);

//    shapeProperties15.Append(outline23);

//    gridlineMinor1.Append(lineReference17);
//    gridlineMinor1.Append(fillReference17);
//    gridlineMinor1.Append(effectReference17);
//    gridlineMinor1.Append(fontReference17);
//    gridlineMinor1.Append(shapeProperties15);

//    Cs.HiLoLine hiLoLine1 = new Cs.HiLoLine();
//    Cs.LineReference lineReference18 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference18 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference18 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference18 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference18.Append(schemeColor57);

//    Cs.ShapeProperties shapeProperties16 = new Cs.ShapeProperties();

//    A.Outline outline24 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill34 = new A.SolidFill();

//    A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation30 = new A.LuminanceModulation() { Val = 50000 };
//    A.LuminanceOffset luminanceOffset26 = new A.LuminanceOffset() { Val = 50000 };

//    schemeColor58.Append(luminanceModulation30);
//    schemeColor58.Append(luminanceOffset26);

//    solidFill34.Append(schemeColor58);
//    A.Round round13 = new A.Round();

//    outline24.Append(solidFill34);
//    outline24.Append(round13);

//    shapeProperties16.Append(outline24);

//    hiLoLine1.Append(lineReference18);
//    hiLoLine1.Append(fillReference18);
//    hiLoLine1.Append(effectReference18);
//    hiLoLine1.Append(fontReference18);
//    hiLoLine1.Append(shapeProperties16);

//    Cs.LeaderLine leaderLine1 = new Cs.LeaderLine();
//    Cs.LineReference lineReference19 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference19 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference19 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference19 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference19.Append(schemeColor59);

//    Cs.ShapeProperties shapeProperties17 = new Cs.ShapeProperties();

//    A.Outline outline25 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill35 = new A.SolidFill();

//    A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation31 = new A.LuminanceModulation() { Val = 35000 };
//    A.LuminanceOffset luminanceOffset27 = new A.LuminanceOffset() { Val = 65000 };

//    schemeColor60.Append(luminanceModulation31);
//    schemeColor60.Append(luminanceOffset27);

//    solidFill35.Append(schemeColor60);
//    A.Round round14 = new A.Round();

//    outline25.Append(solidFill35);
//    outline25.Append(round14);

//    shapeProperties17.Append(outline25);

//    leaderLine1.Append(lineReference19);
//    leaderLine1.Append(fillReference19);
//    leaderLine1.Append(effectReference19);
//    leaderLine1.Append(fontReference19);
//    leaderLine1.Append(shapeProperties17);

//    Cs.LegendStyle legendStyle1 = new Cs.LegendStyle();
//    Cs.LineReference lineReference20 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference20 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference20 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference20 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation32 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset28 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor61.Append(luminanceModulation32);
//    schemeColor61.Append(luminanceOffset28);

//    fontReference20.Append(schemeColor61);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType7 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    legendStyle1.Append(lineReference20);
//    legendStyle1.Append(fillReference20);
//    legendStyle1.Append(effectReference20);
//    legendStyle1.Append(fontReference20);
//    legendStyle1.Append(textCharacterPropertiesType7);

//    Cs.PlotArea plotArea2 = new Cs.PlotArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
//    Cs.LineReference lineReference21 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference21 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference21 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference21 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference21.Append(schemeColor62);

//    plotArea2.Append(lineReference21);
//    plotArea2.Append(fillReference21);
//    plotArea2.Append(effectReference21);
//    plotArea2.Append(fontReference21);

//    Cs.PlotArea3D plotArea3D1 = new Cs.PlotArea3D() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
//    Cs.LineReference lineReference22 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference22 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference22 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference22 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference22.Append(schemeColor63);

//    plotArea3D1.Append(lineReference22);
//    plotArea3D1.Append(fillReference22);
//    plotArea3D1.Append(effectReference22);
//    plotArea3D1.Append(fontReference22);

//    Cs.SeriesAxis seriesAxis1 = new Cs.SeriesAxis();
//    Cs.LineReference lineReference23 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference23 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference23 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference23 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation33 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset29 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor64.Append(luminanceModulation33);
//    schemeColor64.Append(luminanceOffset29);

//    fontReference23.Append(schemeColor64);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType8 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    seriesAxis1.Append(lineReference23);
//    seriesAxis1.Append(fillReference23);
//    seriesAxis1.Append(effectReference23);
//    seriesAxis1.Append(fontReference23);
//    seriesAxis1.Append(textCharacterPropertiesType8);

//    Cs.SeriesLine seriesLine1 = new Cs.SeriesLine();
//    Cs.LineReference lineReference24 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference24 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference24 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference24 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference24.Append(schemeColor65);

//    Cs.ShapeProperties shapeProperties18 = new Cs.ShapeProperties();

//    A.Outline outline26 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill36 = new A.SolidFill();

//    A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation34 = new A.LuminanceModulation() { Val = 35000 };
//    A.LuminanceOffset luminanceOffset30 = new A.LuminanceOffset() { Val = 65000 };

//    schemeColor66.Append(luminanceModulation34);
//    schemeColor66.Append(luminanceOffset30);

//    solidFill36.Append(schemeColor66);
//    A.Round round15 = new A.Round();

//    outline26.Append(solidFill36);
//    outline26.Append(round15);

//    shapeProperties18.Append(outline26);

//    seriesLine1.Append(lineReference24);
//    seriesLine1.Append(fillReference24);
//    seriesLine1.Append(effectReference24);
//    seriesLine1.Append(fontReference24);
//    seriesLine1.Append(shapeProperties18);

//    Cs.TitleStyle titleStyle1 = new Cs.TitleStyle();
//    Cs.LineReference lineReference25 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference25 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference25 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference25 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation35 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset31 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor67.Append(luminanceModulation35);
//    schemeColor67.Append(luminanceOffset31);

//    fontReference25.Append(schemeColor67);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType9 = new Cs.TextCharacterPropertiesType() { FontSize = 1862, Bold = false, Kerning = 1200, Spacing = 0, Baseline = 0 };

//    titleStyle1.Append(lineReference25);
//    titleStyle1.Append(fillReference25);
//    titleStyle1.Append(effectReference25);
//    titleStyle1.Append(fontReference25);
//    titleStyle1.Append(textCharacterPropertiesType9);

//    Cs.TrendlineStyle trendlineStyle1 = new Cs.TrendlineStyle();

//    Cs.LineReference lineReference26 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.StyleColor styleColor7 = new Cs.StyleColor() { Val = "auto" };

//    lineReference26.Append(styleColor7);
//    Cs.FillReference fillReference26 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference26 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference26 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference26.Append(schemeColor68);

//    Cs.ShapeProperties shapeProperties19 = new Cs.ShapeProperties();

//    A.Outline outline27 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

//    A.SolidFill solidFill37 = new A.SolidFill();
//    A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

//    solidFill37.Append(schemeColor69);
//    A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.SystemDot };

//    outline27.Append(solidFill37);
//    outline27.Append(presetDash1);

//    shapeProperties19.Append(outline27);

//    trendlineStyle1.Append(lineReference26);
//    trendlineStyle1.Append(fillReference26);
//    trendlineStyle1.Append(effectReference26);
//    trendlineStyle1.Append(fontReference26);
//    trendlineStyle1.Append(shapeProperties19);

//    Cs.TrendlineLabel trendlineLabel1 = new Cs.TrendlineLabel();
//    Cs.LineReference lineReference27 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference27 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference27 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference27 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation36 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset32 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor70.Append(luminanceModulation36);
//    schemeColor70.Append(luminanceOffset32);

//    fontReference27.Append(schemeColor70);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType10 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    trendlineLabel1.Append(lineReference27);
//    trendlineLabel1.Append(fillReference27);
//    trendlineLabel1.Append(effectReference27);
//    trendlineLabel1.Append(fontReference27);
//    trendlineLabel1.Append(textCharacterPropertiesType10);

//    Cs.UpBar upBar1 = new Cs.UpBar();
//    Cs.LineReference lineReference28 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference28 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference28 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference28 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference28.Append(schemeColor71);

//    Cs.ShapeProperties shapeProperties20 = new Cs.ShapeProperties();

//    A.SolidFill solidFill38 = new A.SolidFill();
//    A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

//    solidFill38.Append(schemeColor72);

//    A.Outline outline28 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

//    A.SolidFill solidFill39 = new A.SolidFill();

//    A.SchemeColor schemeColor73 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation37 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset33 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor73.Append(luminanceModulation37);
//    schemeColor73.Append(luminanceOffset33);

//    solidFill39.Append(schemeColor73);
//    A.Round round16 = new A.Round();

//    outline28.Append(solidFill39);
//    outline28.Append(round16);

//    shapeProperties20.Append(solidFill38);
//    shapeProperties20.Append(outline28);

//    upBar1.Append(lineReference28);
//    upBar1.Append(fillReference28);
//    upBar1.Append(effectReference28);
//    upBar1.Append(fontReference28);
//    upBar1.Append(shapeProperties20);

//    Cs.ValueAxis valueAxis2 = new Cs.ValueAxis();
//    Cs.LineReference lineReference29 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference29 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference29 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference29 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

//    A.SchemeColor schemeColor74 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
//    A.LuminanceModulation luminanceModulation38 = new A.LuminanceModulation() { Val = 65000 };
//    A.LuminanceOffset luminanceOffset34 = new A.LuminanceOffset() { Val = 35000 };

//    schemeColor74.Append(luminanceModulation38);
//    schemeColor74.Append(luminanceOffset34);

//    fontReference29.Append(schemeColor74);
//    Cs.TextCharacterPropertiesType textCharacterPropertiesType11 = new Cs.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };

//    valueAxis2.Append(lineReference29);
//    valueAxis2.Append(fillReference29);
//    valueAxis2.Append(effectReference29);
//    valueAxis2.Append(fontReference29);
//    valueAxis2.Append(textCharacterPropertiesType11);

//    Cs.Wall wall1 = new Cs.Wall();
//    Cs.LineReference lineReference30 = new Cs.LineReference() { Index = (UInt32Value)0U };
//    Cs.FillReference fillReference30 = new Cs.FillReference() { Index = (UInt32Value)0U };
//    Cs.EffectReference effectReference30 = new Cs.EffectReference() { Index = (UInt32Value)0U };

//    Cs.FontReference fontReference30 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
//    A.SchemeColor schemeColor75 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

//    fontReference30.Append(schemeColor75);

//    Cs.ShapeProperties shapeProperties21 = new Cs.ShapeProperties();
//    A.NoFill noFill18 = new A.NoFill();

//    A.Outline outline29 = new A.Outline();
//    A.NoFill noFill19 = new A.NoFill();

//    outline29.Append(noFill19);

//    shapeProperties21.Append(noFill18);
//    shapeProperties21.Append(outline29);

//    wall1.Append(lineReference30);
//    wall1.Append(fillReference30);
//    wall1.Append(effectReference30);
//    wall1.Append(fontReference30);
//    wall1.Append(shapeProperties21);

//    chartStyle1.Append(axisTitle1);
//    chartStyle1.Append(categoryAxis2);
//    chartStyle1.Append(chartArea1);
//    chartStyle1.Append(dataLabel1);
//    chartStyle1.Append(dataLabelCallout1);
//    chartStyle1.Append(dataPoint1);
//    chartStyle1.Append(dataPoint3D1);
//    chartStyle1.Append(dataPointLine1);
//    chartStyle1.Append(dataPointMarker1);
//    chartStyle1.Append(markerLayoutProperties1);
//    chartStyle1.Append(dataPointWireframe1);
//    chartStyle1.Append(dataTableStyle1);
//    chartStyle1.Append(downBar1);
//    chartStyle1.Append(dropLine1);
//    chartStyle1.Append(errorBar1);
//    chartStyle1.Append(floor1);
//    chartStyle1.Append(gridlineMajor1);
//    chartStyle1.Append(gridlineMinor1);
//    chartStyle1.Append(hiLoLine1);
//    chartStyle1.Append(leaderLine1);
//    chartStyle1.Append(legendStyle1);
//    chartStyle1.Append(plotArea2);
//    chartStyle1.Append(plotArea3D1);
//    chartStyle1.Append(seriesAxis1);
//    chartStyle1.Append(seriesLine1);
//    chartStyle1.Append(titleStyle1);
//    chartStyle1.Append(trendlineStyle1);
//    chartStyle1.Append(trendlineLabel1);
//    chartStyle1.Append(upBar1);
//    chartStyle1.Append(valueAxis2);
//    chartStyle1.Append(wall1);

//    chartStylePart1.ChartStyle = chartStyle1;
//}

void GenerateChartColorStylePart1Content(ChartColorStylePart chartColorStylePart1)
{
    Cs.ColorStyle colorStyle1 = new Cs.ColorStyle() { Method = "cycle", Id = (UInt32Value)10U };
    colorStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
    colorStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
    A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
    A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };
    A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };
    A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 };
    A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent5 };
    A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent6 };
    Cs.ColorStyleVariation colorStyleVariation1 = new Cs.ColorStyleVariation();

    Cs.ColorStyleVariation colorStyleVariation2 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 60000 };

    colorStyleVariation2.Append(luminanceModulation7);

    Cs.ColorStyleVariation colorStyleVariation3 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 80000 };
    A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 20000 };

    colorStyleVariation3.Append(luminanceModulation8);
    colorStyleVariation3.Append(luminanceOffset7);

    Cs.ColorStyleVariation colorStyleVariation4 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 80000 };

    colorStyleVariation4.Append(luminanceModulation9);

    Cs.ColorStyleVariation colorStyleVariation5 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 60000 };
    A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset() { Val = 40000 };

    colorStyleVariation5.Append(luminanceModulation10);
    colorStyleVariation5.Append(luminanceOffset8);

    Cs.ColorStyleVariation colorStyleVariation6 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 50000 };

    colorStyleVariation6.Append(luminanceModulation11);

    Cs.ColorStyleVariation colorStyleVariation7 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 70000 };
    A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset() { Val = 30000 };

    colorStyleVariation7.Append(luminanceModulation12);
    colorStyleVariation7.Append(luminanceOffset9);

    Cs.ColorStyleVariation colorStyleVariation8 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 70000 };

    colorStyleVariation8.Append(luminanceModulation13);

    Cs.ColorStyleVariation colorStyleVariation9 = new Cs.ColorStyleVariation();
    A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 50000 };
    A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset() { Val = 50000 };

    colorStyleVariation9.Append(luminanceModulation14);
    colorStyleVariation9.Append(luminanceOffset10);

    colorStyle1.Append(schemeColor19);
    colorStyle1.Append(schemeColor20);
    colorStyle1.Append(schemeColor21);
    colorStyle1.Append(schemeColor22);
    colorStyle1.Append(schemeColor23);
    colorStyle1.Append(schemeColor24);
    colorStyle1.Append(colorStyleVariation1);
    colorStyle1.Append(colorStyleVariation2);
    colorStyle1.Append(colorStyleVariation3);
    colorStyle1.Append(colorStyleVariation4);
    colorStyle1.Append(colorStyleVariation5);
    colorStyle1.Append(colorStyleVariation6);
    colorStyle1.Append(colorStyleVariation7);
    colorStyle1.Append(colorStyleVariation8);
    colorStyle1.Append(colorStyleVariation9);

    chartColorStylePart1.ColorStyle = colorStyle1;
}

string ConvertToBase64(string[,] arrayData)
{
    // Convert the 2D array to a 1D array
    string[] flatArray = FlattenArray(arrayData);

    // Convert the 1D array to a byte array
    byte[] byteArray;
    var bf = new BinaryFormatter();
    using (var ms = new MemoryStream())
    {
        bf.Serialize(ms, flatArray);
        byteArray = ms.ToArray();
    }

    // Convert the byte array to a base64 string
    string base64String = Convert.ToBase64String(byteArray);

    return base64String;
}

string[] FlattenArray(string[,] arrayData)
{
    int rows = arrayData.GetLength(0);
    int cols = arrayData.GetLength(1);
    string[] flatArray = new string[rows * cols];
    int index = 0;

    for (int i = 0; i < rows; i++)
    {
        for (int j = 0; j < cols; j++)
        {
            flatArray[index++] = arrayData[i, j];
        }
    }

    return flatArray;
}

void GenerateEmbeddedPackagePart1Content(EmbeddedPackagePart embeddedPackagePart1)
{
    // Convert the excelData to a base64 string
    string base64ExcelData = ConvertToBase64(excelData);

    // Convert the base64 string to a stream
    System.IO.Stream data = new System.IO.MemoryStream(System.Convert.FromBase64String(base64ExcelData));

    // Feed the data into the embedded package part
    embeddedPackagePart1.FeedData(data);
    data.Close();
}
SlidePart CreateSlidePart(PresentationPart presentationPart)
{
    SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
    slidePart1.Slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new BodyProperties(),
                            new ListStyle(),
                            new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
            new C.ColorMapOverride(new MasterColorMapping()));
    return slidePart1;
}

SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
{
    SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
    SlideLayout slideLayout = new SlideLayout(
    new CommonSlideData(new ShapeTree(
      new P.NonVisualGroupShapeProperties(
      new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
      new P.NonVisualGroupShapeDrawingProperties(),
      new ApplicationNonVisualDrawingProperties()),
      new GroupShapeProperties(new TransformGroup()),
      new P.Shape(
      new P.NonVisualShapeProperties(
        new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
        new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
        new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
      new P.ShapeProperties(),
      new P.TextBody(
        new BodyProperties(),
        new ListStyle(),
        new Paragraph(new EndParagraphRunProperties()))))),
    new C.ColorMapOverride(new MasterColorMapping()));
    slideLayoutPart1.SlideLayout = slideLayout;
    return slideLayoutPart1;
}

SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
{
    SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
    SlideMaster slideMaster = new SlideMaster(
    new CommonSlideData(new ShapeTree(
      new P.NonVisualGroupShapeProperties(
      new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
      new P.NonVisualGroupShapeDrawingProperties(),
      new ApplicationNonVisualDrawingProperties()),
      new GroupShapeProperties(new TransformGroup()),
      new P.Shape(
      new P.NonVisualShapeProperties(
        new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
        new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
      new P.ShapeProperties(),
      new P.TextBody(
        new BodyProperties(),
        new ListStyle(),
        new Paragraph())))),
    new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
    new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
    new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
    slideMasterPart1.SlideMaster = slideMaster;

    return slideMasterPart1;
}

ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
{
    ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
    D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

    D.ThemeElements themeElements1 = new D.ThemeElements(
    new D.ColorScheme(
      new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
      new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
      new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
      new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
      new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
      new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
      new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
      new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
      new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
      new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
      new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
      new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
    { Name = "Office" },
      new D.FontScheme(
      new D.MajorFont(
      new D.LatinFont() { Typeface = "Calibri" },
      new D.EastAsianFont() { Typeface = "" },
      new D.ComplexScriptFont() { Typeface = "" }),
      new D.MinorFont(
      new D.LatinFont() { Typeface = "Calibri" },
      new D.EastAsianFont() { Typeface = "" },
      new D.ComplexScriptFont() { Typeface = "" }))
      { Name = "Office" },
      new D.FormatScheme(
      new D.FillStyleList(
      new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
      new D.GradientFill(
        new D.GradientStopList(
        new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
          new D.SaturationModulation() { Val = 300000 })
        { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
         new D.SaturationModulation() { Val = 300000 })
        { Val = D.SchemeColorValues.PhColor })
        { Position = 35000 },
        new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
         new D.SaturationModulation() { Val = 350000 })
        { Val = D.SchemeColorValues.PhColor })
        { Position = 100000 }
        ),
        new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
      new D.NoFill(),
      new D.PatternFill(),
      new D.GroupFill()),
      new D.LineStyleList(
      new D.Outline(
        new D.SolidFill(
        new D.SchemeColor(
          new D.Shade() { Val = 95000 },
          new D.SaturationModulation() { Val = 105000 })
        { Val = D.SchemeColorValues.PhColor }),
        new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
      {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
      },
      new D.Outline(
        new D.SolidFill(
        new D.SchemeColor(
          new D.Shade() { Val = 95000 },
          new D.SaturationModulation() { Val = 105000 })
        { Val = D.SchemeColorValues.PhColor }),
        new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
      {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
      },
      new D.Outline(
        new D.SolidFill(
        new D.SchemeColor(
          new D.Shade() { Val = 95000 },
          new D.SaturationModulation() { Val = 105000 })
        { Val = D.SchemeColorValues.PhColor }),
        new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
      {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
      }),
      new D.EffectStyleList(
      new D.EffectStyle(
        new D.EffectList(
        new D.OuterShadow(
          new D.RgbColorModelHex(
          new D.Alpha() { Val = 38000 })
          { Val = "000000" })
        { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
      new D.EffectStyle(
        new D.EffectList(
        new D.OuterShadow(
          new D.RgbColorModelHex(
          new D.Alpha() { Val = 38000 })
          { Val = "000000" })
        { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
      new D.EffectStyle(
        new D.EffectList(
        new D.OuterShadow(
          new D.RgbColorModelHex(
          new D.Alpha() { Val = 38000 })
          { Val = "000000" })
        { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
      new D.BackgroundFillStyleList(
      new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
      new D.GradientFill(
        new D.GradientStopList(
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 }),
        new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
      new D.GradientFill(
        new D.GradientStopList(
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 },
        new D.GradientStop(
          new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
        { Position = 0 }),
        new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
      { Name = "Office" });

    theme1.Append(themeElements1);
    theme1.Append(new D.ObjectDefaults());
    theme1.Append(new D.ExtraColorSchemeList());

    themePart1.Theme = theme1;
    return themePart1;
}

