require 'rspec'
require 'ooxml_parser'

describe 'My behaviour' do
  describe 'Chart' do
    it 'Insert Chart Window | Check Set Display X Axis' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartDisplayXAxis.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.picture.chart.axises.first.title).not_to be_nil
    end

    it 'Insert Chart Window | Check Set Not Display X Axis' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartNotDisplayXAxis.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.picture.chart.axises.first.title).to be_nil
    end

    it 'Insert Chart Window | Check Legend Left Position without Overlay' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartLegendLeftPosition.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.picture.chart.legend.position).to eq(:left)
      expect(drawing.picture.chart.legend.overlay).to eq(false)
    end

    it 'Insert Chart Window | Check Legend Left Position with Overlay' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartLegendLeftPositionWithOverlay.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.picture.chart.legend.overlay).to eq(true)
    end

    describe 'Axises' do
      it 'Chart - Axises - ChartAxisDisplayNone.xlsx' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/Axises/ChartAxisDisplayNone.xlsx')
        drawing = xlsx.worksheets[0].drawings[0]
        expect(drawing.picture.chart.axises[1].display).to eq(false)
      end

      it 'Chart - Axises - ChartAxisDisplayOverlay.xlsx' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/Axises/ChartAxisDisplayOverlay.xlsx')
        drawing = xlsx.worksheets[0].drawings[0]
        expect(drawing.picture.chart.axises[1].display).to eq(true)
      end

      it 'Chart - Axises - ChartAxisPosition.xlsx' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/Axises/ChartAxisPosition.xlsx')
        drawing = xlsx.worksheets[0].drawings[0]
        expect(drawing.picture.chart.axises[1].position).to eq(:left)
      end
    end

    describe 'Legend position' do
      it 'Chart - Legend - ChartLegendLeftOverlay.xlsx' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/Legend/ChartLegendLeftOverlay.xlsx')
        drawing = xlsx.worksheets[0].drawings[0]
        expect(drawing.picture.chart.legend.position_with_overlay).to eq(:left_overlay)
      end

      it 'Chart - Legend - ChartLegendRightOverlay.xlsx' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/Legend/ChartLegendRightOverlay.xlsx')
        drawing = xlsx.worksheets[0].drawings[0]
        expect(drawing.picture.chart.legend.position_with_overlay).to eq(:right_overlay)
      end

      it 'Chart - Legend - ChartLegendRight.xlsx' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/Legend/ChartLegendRight.xlsx')
        drawing = xlsx.worksheets[0].drawings[0]
        expect(drawing.picture.chart.legend.position_with_overlay).to eq(:right)
      end
    end

    describe 'Labels' do
      it 'Chart - Labels - ChartDisplayLabelsPosition.xlsx' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/Labels/ChartDisplayLabelsPosition.xlsx')
        expect(xlsx.worksheets[0].drawings[1].picture.chart.display_labels.position).to eq(:top)
      end
    end
  end

  it 'Добавление новой автофигуры к существующей группе автофигур' do
    # spec/spreadsheet_editor/several_users/collaborative/smoke/shape_collaborative_spec.rb
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeGrouping.xlsx')
    expect(xlsx.worksheets[0].drawings[0].grouping.nil?).not_to eq(true)
    expect(xlsx.worksheets[0].drawings[0].grouping.shapes.length).to eq(1)
    expect(xlsx.worksheets[0].drawings[0].grouping.grouping.shapes.length).to eq(2)
  end

  it 'Shape Stroke Size None' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeStrokeWidthNone.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.shape.properties.line.stroke_size).to eq(0)
  end

  it 'Shape Stroke Size 6px' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeStrokeWidth6px.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.shape.properties.line.stroke_size).to eq(6)
  end

  it 'Chart Without Title' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartWithoutTitle.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.picture.chart.title.elements.first).to be_nil
  end

  it 'Chart With Title' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartWithTitle.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.picture.chart.title.elements.first.characters.first.text).to eq('Custom Title')
  end

  it 'Chart With Empty Title' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartWithEmptyTitle.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.picture.chart.title.elements.first.characters.first).to be_nil
  end

  it 'Incorrect image resource' do
    expect { OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/IncorrectImageResource.xlsx') }.to raise_error(LoadError)
  end

  it 'Cell Text Not Empty' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextNotEmpty.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].text).not_to be_empty
  end

  it 'Cell Bold Style' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/BoldStyle.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.font_style.bold).to be_truthy
  end

  it 'Cell Fill Color' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellFillColor.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(197, 224, 180))
  end

  it 'Font Color 255-255-255' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/FontColor255-255-255.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.color).to eq(OoxmlParser::Color.new(255, 255, 255))
  end

  it 'Font Color Theme 237-125-49' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/FontColor237-125-49.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.color).to eq(OoxmlParser::Color.new(237, 125, 49))
  end

  it 'Font Color Standart 255-0-0' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/FontColorStandart255-0-0.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.color).to eq(OoxmlParser::Color.new(255, 0, 0))
  end

  it 'ColorBackgroundNone.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ColorBackgroundNone.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(nil, nil, nil))
  end

  it 'Border Color' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/BorderColor.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.borders.left.color).to eq(OoxmlParser::Color.new(84, 130, 53))
  end

  it 'Shape Line 2' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeLine2.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.properties.preset_geometry).to eq(:line)
  end

  it 'ABS_emb' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ABS_emb.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.text).to eq('Data')
  end

  it 'LOOKUP_emb' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/LOOKUP_emb.xlsx')
    expect(xlsx.worksheets[1].rows[5].cells.first.text).to be_empty
  end

  it 'FormulaDoc' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/FormulaDoc.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('a')
  end

  it 'ShapeStrokeColor1' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeStrokeColor_1.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.properties.line.color_scheme.color).to eq(OoxmlParser::Color.new(248, 203, 173))
  end

  it 'ShapeStrokeColor2' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeStrokeColor_2.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.properties.line.color_scheme.color).to eq(OoxmlParser::Color.new(192, 0, 0))
  end

  it 'ChartTitle' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartTitle.xlsx')
    expect(xlsx.worksheets[0].drawings[0].picture.chart.title.elements.first.characters[0].text).to eq('CustomChartTitle')
  end

  it 'ShapesSpacing' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeSpacing.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.paragraphs[0].properties.spacing).to eq(OoxmlParser::Spacing.new(4, 3, 5, :exact))
  end

  it 'ShapesRuns' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeSpacing.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.paragraphs[0].runs).not_to be_empty
  end

  it 'UndoRedoImageIncorrectLink.xlsx' do
    expect { OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/UndoRedoImageIncorrectLink.xlsx') }.to raise_error(LoadError)
  end

  it 'ShapeWithLink.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeWithLink.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.runs.first.properties.hyperlink).not_to be_nil
  end

  it 'ShapeWithLink.xlsx Link To' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ShapeWithLink.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.runs.first.properties.hyperlink).not_to be_nil
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.runs.first.properties.hyperlink.url).to eq('http://www.yandex.ru/')
  end

  it 'FormatAsTable.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/FormatAsTable.xlsx')
    expect(xlsx.worksheets.first.table_parts.first).to be_a(OoxmlParser::XlsxTable)
  end

  it 'Comments.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Comments.xlsx')
    expect(xlsx.worksheets.first.comments.comments.first.characters.first.text).to eq('Hello')
  end

  it 'FreezePanes.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/FreezePanes.xlsx')
    expect(xlsx.worksheets.first.sheet_views.first.pane.top_left_cell).to eq(OoxmlParser::Coordinates.new(3, 'B'))
  end

  it 'AutoFilter.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/AutoFilter.xlsx')
    expect(xlsx.worksheets.first.autofilter.first.column).to eq('A')
    expect(xlsx.worksheets.first.autofilter.first.row).to eq(1)
  end

  it 'ChartGrouping.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ChartGrouping.xlsx')
    expect(xlsx.worksheets[0].drawings[0].grouping.nil?).not_to be_nil
    expect(xlsx.worksheets[0].drawings[0].grouping.pictures.length).to eq(2)
  end

  it 'TextStrikeoutInCell.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextStrikeoutInCell.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.font_style.strike).to eq(:single)
  end

  it 'CellFormat' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellFormat.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].style.numerical_format).to eq('$#,##0.00')
  end

  it 'ImageSize' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ImageSize.xlsx')
    expect(xlsx.worksheets.first.drawings).not_to be_empty
    image_path = xlsx.worksheets.first.drawings.first.picture.path_to_image
    expect(File.size?(image_path)).to be > 0
  end

  it 'CellTextAlignTop.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellTextAlignTop.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].style.alignment.vertical).to eq(:top)
  end

  it 'ColumnStyle.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/ColumnStyle.xlsx')
    expect(xlsx.worksheets.first.columns.first.style.fill_color.color).to eq(OoxmlParser::Color.new(91, 155, 213))
  end

  it 'HyperlinkInternal.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/HyperlinkInternal.xlsx')
    expect(xlsx.worksheets[0].hyperlinks[0].link).to eq(OoxmlParser::Coordinates.new(10, 'F'))
    expect(xlsx.worksheets[0].rows[0].cells.first.text).to eq('yandex')
    expect(xlsx.worksheets[0].hyperlinks[0].tooltip).to eq('go to www.yandex.ru')
  end

  it 'Text Art | FillWithoutOutline.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/FillWithoutOutline.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.color).to eq(OoxmlParser::Color.new(79, 129, 189))
  end

  it 'Text Art | FillWithoutOutline.xlsx - type check' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/FillWithoutOutline.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.type).to eq(:solid)
  end

  it 'Text Art | FillWithoutOutline.xlsx - nil size' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/FillWithoutOutline.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.width).to eq(0)
  end

  it 'Text Art | OutlineWithoutFill.xlsx - color' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/OutlineWithoutFill.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.color_scheme.color).to eq(OoxmlParser::Color.new(149, 55, 53))
  end

  it 'Text Art | OutlineWithoutFill.xlsx - size' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/OutlineWithoutFill.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.width).to eq(1.24)
  end

  it 'Text Art | OutlineWithoutFill6px.xlsx - big size' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/OutlineWithoutFill6px.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.outline.width).to eq(6)
  end

  it 'Text Art | GradientFill.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/GradientFill.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.type).to eq(:gradient)
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.color.gradient_stops.first.color.converted_color).to eq(OoxmlParser::Color.new(179, 162, 199))
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.color.gradient_stops.first.position).to eq(0)
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.color.gradient_stops[1].color.converted_color).to eq(OoxmlParser::Color.new(250, 192, 144))
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.paragraphs.first.runs.first.properties.font_color.color.gradient_stops[1].position).to eq(100)
  end

  it 'Text Art | TransformArchUp.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/TransformArchUp.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.properties.preset_text_warp.preset).to eq(:textArchUp)
  end

  it 'Text Art | TransformTextCircle.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/TextArt/TransformTextCircle.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.properties.preset_text_warp.preset).to eq(:textCircle)
  end

  describe 'CellStyle' do
    it 'CellTextNotStartWithQuote' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellStyle/CellTextNotStartWithQuote.xlsx')
      expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('abc')
      expect(xlsx.worksheets[0].rows[0].cells[0].raw_text).to eq('abc')
    end

    it 'CellTextStartWithQuote' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellStyle/CellTextStartWithQuote.xlsx')
      expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq("'abc")
      expect(xlsx.worksheets[0].rows[0].cells[0].raw_text).to eq('abc')
    end

    it 'underline_single' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellStyle/underline_single.xlsx')
      expect(xlsx.worksheets[0].rows[1].cells[31].style.font.font_style.underlined).to eq(OoxmlParser::Underline.new(:single))
    end

    describe 'fill' do
      it 'cell_style_fill_color_0.5' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellStyle/fill/cell_style_fill_color_tint_0.5.xlsx')
        expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(127, 127, 127))
      end

      it 'cell_style_fill_color_0.35' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellStyle/fill/cell_style_fill_color_tint_0.35.xlsx')
        expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(89, 89, 89))
      end

      it 'cell_style_fill_color_0_theme_0' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellStyle/fill/cell_style_fill_color_tint_0_theme_0.xlsx')
        expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(255, 255, 255))
      end

      it 'cell_style_fill_color_0_theme_1' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/CellStyle/fill/cell_style_fill_color_tint_0_theme_1.xlsx')
        expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(0, 0, 0))
      end
    end
  end

  describe 'editor_specific_documents' do
    describe 'libreoffice' do
      it 'simple_text' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/editor_specific_documents/libreoffice/simple_text.xlsx')
        expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('This is a test')
      end
    end

    describe 'ms_office_2013' do
      it 'simple_text' do
        xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/editor_specific_documents/ms_office_2013/simple_text.xlsx')
        expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('This is a test')
      end
    end
  end

  describe 'formulas' do
    it 't_formula' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/formulas/t_formula.xlsx')
      expect(xlsx.worksheets[0].rows[0].cells[0].formula).to eq('T(1)')
    end

    it 'all_formulas_values' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/formulas/all_formulas_values.xlsx')
      expect(xlsx.all_formula_values.length).to eq(12)
    end
  end
end
