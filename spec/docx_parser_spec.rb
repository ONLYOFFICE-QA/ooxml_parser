require 'rspec'
require 'ooxml_parser'

describe 'My behaviour' do
  it 'Check Diagram offset should !=0' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DiagramOffsetNonZero.docx')
    expect(docx.elements.first.character_style_array.first.drawing.properties.horizontal_position.offset).not_to be_zero
    expect(docx.elements.first.character_style_array.first.drawing.properties.vertical_position.offset).not_to be_zero
  end

  it 'Check Diagram offset should ==0' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DiagramOffsetZero.docx')
    expect(docx.elements.first.character_style_array.first.drawing.properties.horizontal_position.offset).to be_zero
    expect(docx.elements.first.character_style_array.first.drawing.properties.vertical_position.offset).to be_zero
  end

  it 'Apply paragraph style for several paragraphs' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ParagraphStyleForSeveralParagraphs.docx')
    docx.elements.last.character_style_array.delete_at(-1) # remove last character style array that always added
    expect(docx.elements.first).to eq(docx.elements.last)
  end

  it 'Header: insert chart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/HeaderInsertChart.docx')
    elements = docx.page_properties.notes.first.elements
    expect(elements.first.character_style_array.first.drawing.graphic.data.type).to eq(:column)
    expect(elements.first.character_style_array.first.drawing.graphic.data.grouping).to eq(:clustered)
  end

  it 'HeaderFooterCount' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/HeaderFooterCount.docx')
    expect(docx.notes.length).to eq(4)
  end

  it 'Font color 1' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FontColor.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(197, 216, 240))
  end

  it 'Font color 2' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FontColor02.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(220, 216, 195))
  end

  it 'Font color 3' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FontColor03.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(118, 145, 60))
  end

  it 'Font color 4' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FontColor04.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(54, 97, 145))
  end

  it 'Set Gradient Fill for Shape in canvas into paragraph' do
    # spec/document_editor/one_user/smoke/content/shape_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeGradientColor.docx')
    expect(docx.elements.first.character_style_array[0]
             .alternate_content.office2010_content.graphic.data.properties.fill_color.type).to eq(:gradient)
    expect(docx.elements.first.character_style_array[0]
             .alternate_content.office2010_content.graphic.data.properties.fill_color.value.gradient_stops[0].color).to eq(OoxmlParser::Color.new(235, 241, 221))
    expect(docx.elements.first.character_style_array[0]
             .alternate_content.office2010_content.graphic.data.properties.fill_color.value
             .gradient_stops[0].position - 10).to be < 3
    expect(docx.elements.first.character_style_array[0]
             .alternate_content.office2010_content.graphic.data.properties.fill_color.value
             .gradient_stops[1].position - 70).to be < 3
  end

  it 'Add Shape to Group of Shapes in canvas into paragraph with data' do
    # spec/document_editor/one_user/smoke/content/shape_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeGroup.docx')
    alternate = docx.elements.first.character_style_array.first.alternate_content
    expect(alternate.office2007_content.data.elements.length).to eq(2)
    expect(alternate.office2007_content.data.elements[1].type).to eq(:shape)
    expect(alternate.office2007_content.data.elements.first.elements.length).to eq(2)
    expect(alternate.office2007_content.data.elements.first.elements[0].type).to eq(:shape)
    expect(alternate.office2007_content.data.elements.first.elements[1].type).to eq(:shape)

    expect(alternate.office2010_content.graphic.data.elements.length).to eq(2)
    expect(alternate.office2010_content.graphic.data.elements[1].type).to eq(:shape)
    expect(alternate.office2010_content.graphic.data.elements.first.elements.length).to eq(2)
    expect(alternate.office2010_content.graphic.data.elements.first.elements[0].type).to eq(:shape)
    expect(alternate.office2010_content.graphic.data.elements.first.elements[1].type).to eq(:shape)
  end

  it 'Parse default spacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DefaultSpacing.docx')
    expect(docx.elements.first.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.15, :auto))
  end

  it 'Parse Table spacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableCustomSize.docx')
    expect(docx.elements[1].table_properties.table_width).to eq(12.698412698412698)
  end

  it 'Check shape with text' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeWithText.docx')
    expect(docx.elements.first.nonempty_runs.first.alternate_content.office2007_content.data.text_box.first.character_style_array.first.text).to eq('Lorem ipsum dolor sit amet, consectetur adipiscing elit.'\
                                                                                                                                                    ' Integer consequat faucibus eros, sed mattis tortor consectetur '\
                                                                                                                                                    'cursus. Mauris non eros odio. Curabitur velit metus, placerat sit '\
                                                                                                                                                    'amet tempus cursus, pulvinar sed enim. Vivamus odio arcu, volutpat '\
                                                                                                                                                    'gravida imperdiet vitae, mollis eget augue. Sed ultricies viverra '\
                                                                                                                                                    'convallis. Fusce pharetra mi eget')
  end

  it 'ShapeWithTextColor' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeWithTextColor.docx')
    expect(docx.elements.first.nonempty_runs.first.alternate_content
               .office2010_content.graphic.data.text_body.elements.first
               .nonempty_runs.first.font_color).to eq(OoxmlParser::Color.new(150, 62, 3))
  end

  it 'Check shape line' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeLine.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2007_content.type).to eq(:shape)
    expect(alternate_content.office2010_content.graphic.data
             .properties.preset_geometry).to eq(:line)
  end

  it 'ShapeLineCustomGeometry.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeLineCustomGeometry.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2007_content.type).to eq(:shape)
    expect(alternate_content.office2010_content.graphic.data
             .properties.preset_geometry).to eq(:custom)
  end

  it 'Check ParagraphLineBreak' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ParagraphLineBreak.docx')
    expect(docx.elements.first.character_style_array.first.text).to eq("Simple Test Text\rSimple Test Text")
  end

  it 'Check Double Strikethrought.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DoubleStrikethrought.docx')
    expect(docx.elements.first.character_style_array.first.font_style.strike).to eq(:double)
  end

  it 'Check Hyperlink' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/Hyperlink.docx')
    expect(docx.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
  end

  it 'Check Hyperlink with tooltip' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/HyperlinkTestTooltip.docx')
    expect(docx.notes.first.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
    expect(docx.notes.first.elements.first.nonempty_runs.first.link.tooltip).to eq('go to www.yandex.ru')
  end

  it 'Check Docx With Chart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DocxWithChart.docx')
    elements = docx.elements[1].rows.first.cells[0].elements
    expect(elements.first.nonempty_runs.first.drawing.graphic.data.type).to eq(:line)
    expect(elements.first.nonempty_runs.first.drawing.graphic.data.grouping).to eq(:standard)
  end

  it 'Check Docx With Several Chart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DocxWithSeveralChart.docx')
    expect(docx.elements[0].nonempty_runs.first.drawing.graphic.data.type).to eq(:column)
    expect(docx.elements[0].nonempty_runs.first.drawing.graphic.data.grouping).to eq(:clustered)
    expect(docx.elements[1].nonempty_runs.first.drawing.graphic.data.type).to eq(:column)
    expect(docx.elements[1].nonempty_runs.first.drawing.graphic.data.grouping).to eq(:stacked)
  end

  it 'Different Pie Chats' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/PieDrawingType.docx')
    expect(docx.elements[0].nonempty_runs.first.drawing.graphic.data.type).to eq(:pie)
    expect(docx.elements[1].nonempty_runs.first.drawing.graphic.data.type).to eq(:doughnut)
  end

  it 'Two charts in same paragraph' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TwoDrawingsInSameParagraph.docx')
    elements = docx.elements.first.nonempty_runs
    expect(elements.first.drawings[0].graphic.data.type).to eq(:doughnut)
    expect(elements.first.drawings[1].graphic.data.type).to eq(:pie)
  end

  it 'Check Page Break Before' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/PageBreakBefore.docx')
    expect(docx.elements.first.page_break).to eq(true)
  end

  it 'Check Keep Lines Together' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/KeepLinesTogether.docx')
    expect(docx.elements.first.keep_lines).to eq(true)
  end

  it 'Check Keep Next True' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/KeepNext.docx')
    expect(docx.elements.first.keep_next).to eq(true)
  end

  it 'Check Keep Next False' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/KeepLinesTogether.docx')
    expect(docx.elements.first.keep_next).to eq(false)
  end

  it 'Check Parsing docx without statistic' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/NoWordsStatistic.docx')
    expect(docx.document_properties.pages).to be_nil
    expect(docx.document_properties.words).to be_nil
  end

  it 'Check Parsing Table Style Shade None' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableStyleShadeNone.docx')
    expect(docx.elements[1].table_properties.table_style.first_column.cell_properties.shd).to eq(:none)
  end

  it 'Check Parsing Table Style Banding2 without color' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableStyleColumnsNone.docx')
    expect(docx.elements[1].table_properties.table_style.banding_2_horizontal.cell_properties.shd).to eq(:none)
  end

  it 'Check Parsing Table Style Banding2 with color' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableStyleBanding2.docx')
    expect(docx.elements[1].table_properties.table_style.banding_2_horizontal.cell_properties.shd.to_s).to eq('RGB (217, 217, 217)')
  end

  it 'NumberingTest.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/NumberingTest.docx')
    expect(docx.elements.first.numbering).to be_an_instance_of OoxmlParser::Numbering
    expect(docx.elements.first.numbering.numbering_properties.ilvls.first.level_text).to eq('o')
    expect(docx.elements.first.numbering.numbering_properties.ilvls.first.num_format).to eq('bullet')
  end

  it 'DropCapSmoke.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DropCapSmoke.docx')
    expect(docx.elements[0].borders.top.space.round(1)).to eq(1.2)
    expect(docx.elements[0].borders.right.space.round(1)).to eq(0.8)
    expect(docx.elements[0].borders.bottom.space.round(1)).to eq(2.3)
    expect(docx.elements[0].borders.left.space.round(1)).to eq(0.4)
  end

  it 'Check ImageWithInccorectLinkInRes' do
    expect { OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ImageWithInccorectLinkInRes.docx') }.to raise_error(LoadError)
  end

  it 'ShapeInsertAllFillPicture.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeInsertAllFillPicture.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2010_content.graphic.data.properties.fill_color.type).to eq(:picture)
  end

  it 'CommentObject.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/CommentObject.docx')
    expect(docx.element_by_description(location: :comment, type: :paragraph)[0].character_style_array[0].text).to eq('Is it true?')
  end

  it 'TableMerge.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableMerge.docx')
    docx.elements[1].rows[0].cells[0].elements.each do |element|
      expect(element.nonempty_runs.first.text).not_to be_empty
    end
  end

  it 'TableIndent.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableIndent.docx')
    expect(docx.element_by_description[1].table_properties
               .table_indent.round(1)).to eq(1.5)
  end

  it 'ShapePatternFill.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapePatternFill.docx')
    expect(docx.element_by_description(location: :canvas, type: :paragraph).first.nonempty_runs.first
             .alternate_content.office2010_content.graphic.data.properties.fill_color.type).to eq(:pattern)
    expect(docx.element_by_description(location: :canvas, type: :paragraph).first.nonempty_runs.first
             .alternate_content.office2010_content.graphic.data.properties.fill_color.value
             .preset).to eq(:dkUpDiag)
    expect(docx.element_by_description(location: :canvas, type: :paragraph).first.nonempty_runs.first
             .alternate_content.office2010_content.graphic.data.properties
             .fill_color.value.foreground_color).to eq(OoxmlParser::Color.new(237, 125, 49))
    expect(docx.element_by_description(location: :canvas, type: :paragraph).first.nonempty_runs.first
             .alternate_content.office2010_content.graphic.data.properties
             .fill_color.value.background_color).to eq(OoxmlParser::Color.new(237, 125, 49))
  end

  it 'TableCellColor.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableCellColor.docx')
    docx.element_by_description(location: :canvas, type: :paragraph)[1].rows.each do |each_row|
      each_row.cells.each do |current_cell|
        expect(current_cell.cell_properties.shd).to eq(OoxmlParser::Color.new(217, 226, 242))
      end
    end
  end

  it 'TableCustomCellMargin.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableCustomCellMargin.docx')
    margins_in_doc = docx.element_by_description(location: :canvas, type: :paragraph)[1].rows[0].cells[0].cell_properties.table_cell_margin
    expect(OoxmlParser::TableMargins.new(false, 0.5, 3.19, 6.0, 3.19)).to eq(margins_in_doc)
  end

  it 'Formulas.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/Formulas.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasSymbol.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasSymbol.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulaFraction.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulaFraction.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasIntegral.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasIntegral.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasBrackets.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasBrackets.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasFunction.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasFunction.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasAccent.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasAccent.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasGroupChar.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasGroupChar.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasLimit.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasLimit.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasBar1.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasBar1.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormulasBar2.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormulasBar2.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'FormualsMatrix.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FormualsMatrix.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'CopyTableFromCSE.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/CopyTableFromCSE.docx')
    expect(docx.elements.first.rows.first.cells.first.elements.first.character_style_array.first.text).to eq('Simple Test Text')
  end

  it 'Check Error for Empty zip docx' do
    expect { OoxmlParser::DocxParser.parse_docx('spec/docx_examples/EmptyZipDocx.docx') }.to raise_error(LoadError)
  end

  it 'Footer' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/Footer.docx')
    expect(docx.notes.first.type).to eq('footer1')
    expect(docx.notes.first.elements.first.character_style_array.first.text).to eq('Simple Test Text')
  end

  it 'LinkXpathError' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/LinkXpathError.docx')
    expect(docx.elements.first.nonempty_runs.first.text).to eq('yandex')
    expect(docx.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
    expect(docx.elements.first.nonempty_runs.first.link.tooltip).to eq('go to www.yandex.ru')
  end

  it 'ParagrphEquals' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ParagraphEquals.docx')
    expect(docx.elements[0]).to eq(docx.elements[2])
  end

  it 'TableInTable' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableInTable.docx')
    expect(docx.elements[1].rows[0].cells[0].elements[1].is_a?(OoxmlParser::Table)).to eq(true)
  end

  it 'TableLook' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableLook.docx')
    table_look = docx.elements[1].table_properties.table_look
    expect(table_look.first_column).to eq(false)
    expect(table_look.first_row).to eq(false)
    expect(table_look.last_column).to eq(false)
    expect(table_look.last_row).to eq(false)
    expect(table_look.no_horizontal_banding).to eq(true)
    expect(table_look.no_vertical_banding).to eq(true)
  end

  it 'TablePropertiesPosition' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TablePropertiesPosition.docx')
    expect(docx.elements[1].table_properties.table_properties.position_x).to eq(nil)
  end

  it 'TablePropertiesPositionX' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TablePropertiesPositionX.docx')
    expect(docx.elements[1].table_properties.table_properties.position_x).to eq(2.20)
  end

  it 'TableCellMargin' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableCellMargin.docx')
    margins_in_doc = docx.element_by_description(location: :canvas, type: :paragraph)[1].table_properties.table_cell_margin
    expect(OoxmlParser::TableMargins.new(true, 1.0, 0.19, 2.0, 0.19)).to eq(margins_in_doc)
  end

  it 'TablePropertiesJC' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TablePropertiesJC.docx')
    expect(docx.element_by_description(location: :canvas, type: :paragraph)[1]
               .table_properties.jc).to eq(:center)
  end

  it 'TablePropertiesJCLeft' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TablePropertiesJCLeft.docx')
    expect(docx.element_by_description[1].table_properties.jc).to eq(:left)
  end

  it 'TableShadeColor' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableShadeColor.docx')
    expect(docx.element_by_description(location: :canvas, type: :paragraph)[1]
               .table_properties.shd).to eq(OoxmlParser::Color.new(30, 79, 121))
  end

  it 'TableCellSpacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableCellSpacing.docx')
    docx.element_by_description(location: :canvas, type: :paragraph)[1].rows.each do |current_row|
      expect(current_row.cells_spacing.round(1)).to eq(1.5)
    end
  end

  it 'HeaderFooterSection' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/HeaderFooterSection.docx')
    expect(docx.elements.first.sector_properties.notes.first.elements.first
               .character_style_array.first.text).to eq('Lorem ipsum dolor sit amet, consectetur '\
                              'adipiscing elit. Integer consequat faucibus eros, sed mattis tortor...')
    expect(docx.notes.first.elements.first.character_style_array.first.text).to eq('Simple Test Text')
  end

  it 'SectionIndentHangingNilValue' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/SectionIndentHangingNilValue.docx')
    expect(docx.elements.first.sector_properties.notes.first
               .elements.first.ind).to eq(OoxmlParser::Indents.new(1.02, 2.02, 3.03))
  end

  it 'TableBorders' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableBorders.docx')
    expect(docx.element_by_description(location: :canvas, type: :table)[0].borders.right.val).to eq(:none)
    expect(docx.element_by_description(location: :canvas, type: :table)[1].borders.left.val).to eq(:none)
  end

  it 'TableBorderVisual' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableBorderVisual.docx')
    expect(docx.element_by_description(location: :canvas, type: :table)[0].borders.border_visual_type).to eq(:none)
  end

  it 'LinePropertiesDefault' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/LinePropertiesDefault.docx')
    expect(docx.elements[0].nonempty_runs.first.alternate_content.office2010_content
               .graphic.data.properties.line).to be_nil
  end

  it 'ShapeParagraphSpacing_1' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeParagraphSpacing_1.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(0.0, 0.35, 1.0, :auto))
  end

  it 'ShapeParagraphSpacing_2' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeParagraphSpacing_2.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(10.0, 0.35, 1.0, :auto))
  end

  it 'ShapeParagrpaphSpacing_3' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeParagrpaphSpacing_3.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.5, :auto))
  end

  it 'ShapeDefaultSpacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ShapeDefaultSpacing.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.15, :auto))
  end

  it 'TableLineSpacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TableLineSpacing.docx')
    expect(docx.element_by_description(location: :canvas, type: :table)[0].spacing.line).to eq(1.0)
  end

  it 'ParagraphStyleSecondParagraph' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/ParagraphStyleSecondParagraph.docx')
    expect(docx.element_by_description[1].character_style_array[0].font_style.italic).to eq(false)
  end

  it 'FontSizeFloat' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FontSizeFloat.docx')
    expect(docx.elements[0].character_style_array[1].size).to eq(53.5)
  end

  it 'FontStyleCopiedParagraph' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/FontStyleCopiedParagraph.docx')
    expect(docx.elements[0].nonempty_runs.first.font_style).to eq(OoxmlParser::FontStyle.new(true, true, nil, :none))
  end

  it 'CompareTwoParagraphs' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/CompareTwoParagraphs.docx')
    expect(docx.elements[0]).to eq(docx.elements[1])
  end

  it 'DropCapFontStyle - Strike off' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DropCapFontStyleStrikeOff.docx')
    expect(docx.element_by_description[0].character_style_array[0].font_style).to eq(OoxmlParser::FontStyle.new(false, false))
  end

  it 'DropCapFontStyle - Strike on' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/DropCapFontStyleStrikeOn.docx')
    expect(docx.element_by_description[0].character_style_array[0].font_style).to eq(OoxmlParser::FontStyle.new(false, false, nil, :single))
  end

  it 'TextArt | TextBordersWidth1pt.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/TextBordersWidth1pt.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.width).to eq(1)
  end

  it 'TextArt | TextBordersWidth2.25pt.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/TextBordersWidth2.25pt.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.width).to eq(2.25)
  end

  it 'TextArt | TextOutlineColorStandard.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/TextOutlineColorStandard.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(OoxmlParser::Color.new(255, 0, 0))
  end

  it 'TextArt | TextOutlineColorScheme.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/TextOutlineColorScheme.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(OoxmlParser::Color.new(244, 177, 131))
  end

  it 'TextArt | TransformCircle.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/TransformCircle.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.body_properties.preset_text_warp.preset).to eq(:textCircle)
  end

  it 'TextArt | Transform_arc.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/TransformArc.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.body_properties.preset_text_warp.preset).to eq(:textArchUp)
  end

  it 'TextArt | Opacity50.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/Opacity50.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color.alpha_channel).to eq(50)
  end

  it 'TextArt | Opacity0.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/Opacity0.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color.alpha_channel).to eq(0)
  end

  it 'TextArt | GradientText.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/GradientText.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.type).to eq(:gradient)
  end

  it 'TextArt | GradientText.docx - outline check' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/GradientText.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(:none)
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.type).to eq(:none)
  end

  # File with created with 'no fill' fill
  it 'TextArt | TextWithOutFillColor.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/TextWithOutFillColor.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color).to eq(:none)
  end

  # Take different
  it 'TextArt | ColoreFillIsNouthing.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/ColoreFillIsNouthing.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color).to eq(:none)
  end

  it 'TextArt | WithoutOutlineColor.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/WithoutOutlineColor.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(:none)
  end

  it 'TextArt | WithoutOutlineWidth.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/WithoutOutlineColor.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.width).to eq(0)
  end

  it 'TextArt | WithoutOutline.docx - color outline check' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/TextArt/WithoutOutlineColor.docx')
    expect(docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
               .data.text_body.elements.first.character_style_array.first.text_outline.color_scheme.color).to eq(:none)
  end

  describe 'document style' do
    describe 'page_properties' do
      describe 'page_numbering' do
        it 'page_num_type_empty_format' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_style/page_properties/page_numbering/page_num_type_empty_format.docx')
          expect(docx.page_properties.num_type).to be_nil
        end
      end
    end

    it 'New Paragraph Document Style' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_style/NewParagraphStyles.docx')
      expect(docx.document_styles.last.name).to eq('NewParagraphStyle')
    end

    it 'Paragraph Document Visible Style' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_style/style_visibility.docx')
      expect(docx.document_style_by_name('Heading 8')).to be_visible
      expect(docx.document_style_by_name('Footer')).not_to be_visible
    end

    it 'New Paragraph Document Style: method "style_exist?"' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_style/style_exists.docx')
      expect(docx.style_exist?('NewParagraphStyle')).to be_truthy
    end

    it 'New Paragraph Document Style: method "style_exist?", negative' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_style/style_exists.docx')
      expect(docx.style_exist?('ThisStyleIsNotExist')).to be_falsey
    end
  end

  describe 'editor_specific_documents' do
    describe 'libreoffice' do
      it 'simple_text' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/editor_specific_documents/libreoffice/simple_text.docx')
        expect(docx.elements.first.character_style_array.first.text).to eq('This is a test')
      end
    end

    describe 'ms_office_2013' do
      it 'simple_text' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/editor_specific_documents/ms_office_2013/simple_text.docx')
        expect(docx.elements.first.character_style_array.first.text).to eq('This is a test')
      end
    end
  end

  describe 'compare' do
    it 'Compare two similar docx - set 1' do
      docx1 = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/compare/set_1/first.docx')
      docx2 = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/compare/set_1/second.docx')
      expect(docx1).to eq(docx2)
    end

    it 'Compare two similar docx - set 2' do
      docx1 = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/compare/set_2/first.docx')
      docx2 = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/compare/set_2/second.docx')
      expect(docx1).to eq(docx2)
    end
  end

  describe 'document' do
    it 'document_background' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document/document_background.docx')
      expect(docx.background.color1).to eq(OoxmlParser::Color.new(255, 255, 255))
    end

    it 'document_background_fill' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document/document_background_fill.docx')
      expect(File.exist?(docx.background.image)).to be_truthy
    end
  end

  describe 'document properties' do
    it 'Page Count' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_properties/page_count.docx')
      expect(docx.document_properties.pages).to eq(2)
    end

    it 'Word Count' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_properties/word_count.docx')
      expect(docx.document_properties.words).to eq(1)
    end

    it 'no_app_xml_file' do
      expect do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_properties/no_app_xml_file.docx')
        expect(docx.document_properties.pages).to be_nil
      end.to output(%r{no 'docProps\/app.xml'}).to_stderr
    end
  end

  describe 'elements' do
    describe 'table' do
      describe 'borders' do
        it 'borders_properties_compare' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/borders/borders_properties_compare.docx')
          docx.element_by_description(location: :canvas, type: :paragraph)[1].rows.each do |each_row|
            (2...6).each do |column_counter|
              expect(each_row.cells[column_counter].cell_properties.borders_properties).to eq(OoxmlParser::Borders.new)
            end
          end
        end

        it 'border_properties_nil' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/borders/border_properties_nil.docx')
          expect(docx.element_by_description[0].borders.top).to be_nil
          expect(docx.element_by_description[0].borders.left.sz).to eq(1)
          expect(docx.element_by_description[0].borders.right).to be_nil
          expect(docx.element_by_description[0].borders.bottom).to be_nil
        end
      end

      describe 'cell' do
        describe 'paragraph' do
          describe 'run' do
            it 'table_with_default_table_run_style' do
              docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/cell/paragraph/run/table_with_default_table_run_style.docx')
              expect(docx.elements.first.rows.first.cells.first.elements.first).not_to be_nil
            end
          end
        end

        describe 'properties' do
          it 'cell_properties_merge_restart' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/cell/properties/cell_properties_merge_restart.docx')
            expect(docx.elements[556].rows.first.cells.first.properties.merge.value).to eq(:restart)
          end

          it 'shade_nil' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/cell/properties/shade_nil.docx')
            expect(docx.elements.first.rows.first.cells.first.properties.shd).to eq(:none)
          end

          it 'table_cell_text_no_rotation' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/cell/properties/table_cell_text_no_rotation.docx')
            expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:horizontal)
          end

          it 'table_cell_text_rotated_90' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/cell/properties/table_cell_text_rotated_90.docx')
            expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:rotate_on_90)
          end

          it 'table_cell_text_rotated_270' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/cell/properties/table_cell_text_rotated_270.docx')
            expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:rotate_on_270)
          end
        end
      end

      describe 'row' do
        it 'table_row_height' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/row/table_row_height.docx')
          expect(docx.elements[1].rows[1].table_row_properties.height).to eq(1000)
        end

        it 'border_color_nil_1' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/row/border_color_nil_1.docx')
          expect(docx.elements.first.properties.table_borders.left).to be_nil
          expect(docx.elements.first.properties.table_borders.right).to be_nil
          expect(docx.elements.first.properties.table_borders.top).to be_nil
          expect(docx.elements.first.properties.table_borders.bottom).to be_nil
        end

        it 'border_color_nil_2' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/table/row/border_color_nil_2.docx')
          expect(docx.elements[29].properties.table_borders.left).to be_nil
          expect(docx.elements[29].properties.table_borders.right).to be_nil
          expect(docx.elements[29].properties.table_borders.top).to be_nil
          expect(docx.elements[29].properties.table_borders.bottom).to be_nil
        end
      end
    end
  end

  describe 'chart' do
    describe 'chart' do
      describe 'offset' do
        it 'chart_offset_w14' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/chart/offset/chart_offset_w14.docx')
          elements = docx.element_by_description(location: :canvas, type: :paragraph)
          expect(elements.first.nonempty_runs.first.drawing.properties.horizontal_position.offset).to be_zero
          expect(elements.first.nonempty_runs.first.drawing.properties.vertical_position.offset).to be_zero
        end
      end
    end
  end

  describe 'paragraph' do
    describe 'borders' do
      it 'border_size_nil' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/borders/border_size_nil.docx')
        expect(docx.document_styles[39].paragraph_properties.borders.bar).to be_nil
      end

      it 'borders_space_nil' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/borders/borders_space_nil.docx')
        expect(docx.document_styles[25].paragraph_properties.borders.bar.space).to eq(0)
      end
    end

    describe 'character' do
      describe 'pict' do
        describe 'group' do
          it 'old_docx_group_wrap_nil' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/character/pict/group/old_docx_group_wrap_nil.docx')
            expect(docx.elements.first.rows.first.cells.first.elements.first.character_style_array[2]
                       .alternate_content.office2007_content.data.properties.wrap).to be_nil
          end

          it 'old_docx_shape' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/character/pict/group/old_docx_shape.docx')
            expect(docx.elements[29].character_style_array.first.alternate_content.office2007_content.data.elements
                    .first.elements.first.elements.first.elements.first.object.image).not_to be_nil
          end
        end

        it 'pict_in_character_run' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/character/pict/pict_in_character_run.docx')
          expect(docx.elements.first.character_style_array[5].shape).not_to be_nil
        end

        it 'picture_target_null' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/character/pict/picture_target_null.docx')
          expect(docx.elements[9].rows.first.cells.first.elements.first.character_style_array[1].drawings.first.graphic.data.path_to_image.path).to be_nil
        end

        it 'picture_target_no_blip_internals' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/character/pict/picture_target_no_blip_internals.docx')
          expect(docx.elements[13].character_style_array[1].alternate_content.office2010_content.graphic.data.elements[1].object.path_to_image.path).to be_nil
        end
      end

      describe 'text' do
        it 'nonbreaking_hyphen' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/character/text/nonbreaking_hyphen.docx')
          expect(docx.elements.first.character_style_array.first.text).to eq('Test–Text')
        end

        it 'tab_inside_run' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/character/text/tab_inside_run.docx')
          expect(docx.elements[1].character_style_array.first.text).to eq("Run\t\StopRun")
        end
      end
    end

    describe 'fields' do
      it 'instruction_type' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/fields/instruction_type.docx')
        expect(docx.elements.first.nonempty_runs.first.instruction).to eq('MERGEFIELD a')
        expect(docx.element_by_description.first.page_numbering).to eq(false)
      end

      it 'field_inside_hyperlink' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/fields/field_inside_hyperlink.docx')
        expect(docx.elements.first.nonempty_runs.first.link.url).to eq('https://www.yandex.ru/')
      end

      it 'text_inside_filed' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/fields/text_inside_field.docx')
        expect(docx.elements.first.nonempty_runs.first.text).to eq('www.ya.ru')
      end

      it 'page_numbering' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/fields/page_numbering.docx')
        expect(docx.element_by_description.first.page_numbering).to eq(true)
        expect(docx.element_by_description.first.character_style_array[1].text).to eq('1')
        expect(docx.element_by_description.first.character_style_array[1].page_number).to eq(true)
      end

      describe 'page_borders' do
        it 'page borders offset' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/page_borders/page_border_offset.docx')
          expect(docx.page_properties.page_borders.offset_from).to eq(:page)
        end
      end

      describe 'page_margins' do
        it 'page margins parsing' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/page_margins/page_margins.docx')
          expect(docx.page_properties.margins).to eq(OoxmlParser::PageMargins.new(top: 2, left: 3, bottom: 2, right: 1.5, gutter: 0, header: 1.25, footer: 1.25))
        end
      end
    end

    describe 'frame_properties' do
      it 'anchor_lock' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/frame_properties/anchor_lock.docx')
        expect(docx.document_styles[17].paragraph_properties.frame_properties.anchor_lock).to be_truthy
      end
    end

    describe 'indents' do
      it 'indents_round' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/indents/indents_round.docx')
        expect(docx.element_by_description.first.ind.round(3).left_indent).to eq(1.25)
        expect(docx.element_by_description.first.ind.round(3).right_indent).to eq(0)
      end

      it 'indents_comparasion' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/indents/indents_comparasion.docx')
        expect(docx.element_by_description.first.ind.equal_with_round(OoxmlParser::Indents.new(0, 1.27, 0))).to be_truthy
      end
    end

    describe 'spacing' do
      describe 'contextual spacing' do
        it 'contextual_spacing false' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/spacing/contextual_spacing_false.docx')
          expect(docx.elements.first.contextual_spacing).to be_falsey
        end

        it 'contextual_spacing true' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/spacing/contextual_spacing_true.docx')
          expect(docx.document_style_by_name('Title').paragraph_properties.contextual_spacing).to be_truthy
        end
      end
    end

    describe 'runs' do
      describe 'drawing' do
        describe 'graphic' do
          it 'chart_color_style' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_color_style.docx')
            expect(docx.elements[0].character_style_array[0].drawing
                       .graphic.data.alternate_content.office2007_content.style_number).to eq(2)
            expect(docx.elements[0].character_style_array[0].drawing
                       .graphic.data.alternate_content.office2010_content.style_number).to eq(102)
          end

          it 'chart_fill_pattern' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_fill_pattern.docx')
            drawing = docx.element_by_description[1].nonempty_runs.first.drawing
            expect(drawing.graphic.data.type).to eq(:line)
            expect(drawing.graphic.data.grouping).to eq(:standard)
            expect(drawing.graphic.data.shape_properties.fill.type).to eq(:picture)
            expect(drawing.graphic.data.shape_properties.fill.value.path_to_media_file).not_to be_nil
          end

          it 'chart_grid_lines_for_x' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_grid_lines_for_x.docx')
            drawing = docx.elements.first.character_style_array.first.drawing
            expect(drawing.graphic.data.axises.first.major_grid_lines).to eq(false)
            expect(drawing.graphic.data.axises.first.minor_grid_lines).to eq(false)
          end

          it 'chart_in_header' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_in_header.docx')
            elements = docx.element_by_description(location: :header, type: :paragraph)
            expect(elements.first.nonempty_runs.first.drawing.graphic.data.type).to eq(:column)
            expect(elements.first.nonempty_runs.first.drawing.graphic.data.grouping).to eq(:clustered)
          end

          it 'chart_in_table' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_in_table.docx')
            expect(docx.elements[1].rows.first.cells.first.elements.first.nonempty_runs.first.drawing.graphic.type).to eq(:chart)
          end

          it 'chart_minor_gridline' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_minor_gridline.docx')
            drawing = docx.elements.first.character_style_array.first.drawing
            expect(drawing.graphic.data.axises[1].major_grid_lines).to eq(false)
            expect(drawing.graphic.data.axises[1].minor_grid_lines).to eq(true)
          end

          it 'chart_point_1' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_point_1.docx')
            expect(docx.elements.first.character_style_array.first.drawing.graphic.data.type).to eq(:point)
            expect(docx.elements.first.character_style_array.first.drawing.graphic.data.grouping).to be_nil
          end

          it 'chart_point_2' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_point_2.docx')
            expect(docx.elements.first.character_style_array.first.drawing.graphic.data.type).to eq(:point)
            expect(docx.elements.first.character_style_array.first.drawing.graphic.data.grouping).to be_nil
          end

          it 'chart_points' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_points.docx')
            drawing = docx.elements.first.character_style_array.first.drawing
            expect(drawing.graphic.data.data.size).to eq(6)
            expect(drawing.graphic.data.data.first.points.size).to eq(3)
          end

          it 'chart_title' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_title.docx')
            expect(docx.elements.first.character_style_array.first.drawing.graphic.data.title.elements.first.characters.first.text).to eq('CustomChartTitle')
          end

          it 'chart_title_with_empty_paragraph' do
            docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/drawing/graphic/chart_title_with_empty_paragraph.docx')
            expect(docx.elements.first.character_style_array.first.drawing.graphic.data.title.elements.first.characters.empty?).to be_falsey
          end
        end
      end

      describe 'font_color' do
        it 'color' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/font_color/color.docx')
          expect(docx.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(64, 64, 64))
        end

        it 'color_nil' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/font_color/color_nil.docx')
          expect(docx.elements[1].rows.first.cells.first.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(nil, nil, nil))
        end

        it 'copied_paragrph_with_color_1' do
          pending('Color problems. I cannot solve it')
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/font_color/copied_paragrph_with_color_1.docx')
          expect(docx.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(197, 216, 240))
          expect(docx.elements.last.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(146, 208, 80))
        end

        it 'copied_paragrph_with_color_2' do
          pending('Color problems. I cannot solve it')
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/font_color/copied_paragrph_with_color_2.docx')
          expect(docx.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(191, 191, 191))
          expect(docx.elements.last.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(95, 74, 121))
        end
      end

      describe 'footnote' do
        it 'footnote_reference' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/footnote/footnote_reference.docx')
          expect(docx.elements[91].character_style_array[23].footnote.id).to eq(1)
        end
      end

      it 'caps_characters_on' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/caps_characters_on.docx')
        expect(docx.elements.first.character_style_array.first.caps).to eq(:caps)
      end

      it 'run_with_link' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/runs/run_with_link.docx')
        expect(docx.elements.first.character_style_array[0].link.link).to eq('http://www.yandex.ru')
        expect(docx.elements.first.character_style_array[0].link.tooltip).to eq('go to www.yandex.ru')
      end
    end

    describe 'shape' do
      describe 'text_box' do
        it 'text_box_list' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/elements/paragraph/shape/text_box/text_box_list.docx')
          expect(docx.elements[160].character_style_array.first.shape.elements.first.character_style_array.first.text).to include('Guidelines')
        end
      end
    end
  end

  describe 'page_properties' do
    describe 'columns' do
      it 'two_columns' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/two_columns.docx')
        expect(docx.page_properties.columns.count).to eq(2)
      end

      it 'ten_columns' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/ten_columns.docx')
        expect(docx.page_properties.columns.count).to eq(10)
      end

      it 'left_column' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/left_column.docx')
        expect(docx.page_properties.columns[0].width).to eq(0.01)
      end

      it 'right_column' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/right_column.docx')
        expect(docx.page_properties.columns[0].width).to eq(0.02)
      end

      it 'several_types_of_columns' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/several_types_of_columns.docx')
        expect(docx.elements[2].sector_properties.columns.count).to eq(2)
        expect(docx.elements[5].sector_properties.columns.count).to eq(3)
      end

      describe 'equal_columns' do
        it 'equal_columns' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/equal_columns/equal_columns_undefined.docx')
          expect(docx.page_properties.columns.equal_width).to be_nil
        end

        it 'non_equal_columns' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/equal_columns/non_equal_columns.docx')
          expect(docx.page_properties.columns).not_to be_equal_width
        end

        it 'equal_columns_undefined' do
          docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/columns/equal_columns/equal_columns.docx')
          expect(docx.page_properties.columns).to be_equal_width
        end
      end
    end

    it 'page_size.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/page_properties/page_size.docx')
      expect(docx.page_properties.size.width).to eq(5)
      expect(docx.page_properties.size.height).to eq(29.7)
    end
  end
  describe 'shape' do
    describe 'lines' do
      it 'ShapeLineEnding' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/shape/lines/ShapeLineEnding.docx')
        alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
        expect(alternate_content.office2010_content.graphic.data
                   .properties.line.head_end.length).to eq(:medium)
      end

      it 'ShapeLinePoints' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/shape/lines/ShapeLinePoints.docx')
        alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
        expect(alternate_content.office2010_content.graphic.data
                   .properties.custom_geometry.paths_list.first.elements.first.points.first).to eq(OoxmlParser::OOXMLCoordinates.new(5_440, 16_475))
      end

      it 'ShapeLineObjectSize' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/shape/lines/ShapeLineObjectSize.docx')
        alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
        expect(alternate_content.office2010_content.properties.object_size).to eq(OoxmlParser::OOXMLCoordinates.new(11.033, 12.543))
      end
    end
  end
end
