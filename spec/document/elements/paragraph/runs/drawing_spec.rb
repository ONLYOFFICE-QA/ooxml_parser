require 'spec_helper'

describe 'drawing' do
  describe 'graphic' do
    it 'chart_color_style' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_color_style.docx')
      expect(docx.elements[0].character_style_array[0].drawing
                 .graphic.data.alternate_content.office2007_content.style_number).to eq(2)
      expect(docx.elements[0].character_style_array[0].drawing
                 .graphic.data.alternate_content.office2010_content.style_number).to eq(102)
    end

    it 'chart_fill_pattern' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_fill_pattern.docx')
      drawing = docx.element_by_description[1].nonempty_runs.first.drawing
      expect(drawing.graphic.data.type).to eq(:line)
      expect(drawing.graphic.data.grouping).to eq(:standard)
      expect(drawing.graphic.data.shape_properties.fill.type).to eq(:picture)
      expect(drawing.graphic.data.shape_properties.fill.value.file_reference.content.length).to be > 1000
    end

    it 'chart_grid_lines_for_x' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_grid_lines_for_x.docx')
      drawing = docx.elements.first.character_style_array.first.drawing
      expect(drawing.graphic.data.axises.first.major_grid_lines).to eq(false)
      expect(drawing.graphic.data.axises.first.minor_grid_lines).to eq(false)
    end

    it 'chart_in_header' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_in_header.docx')
      elements = docx.element_by_description(location: :header, type: :paragraph)
      expect(elements.first.nonempty_runs.first.drawing.graphic.data.type).to eq(:column)
      expect(elements.first.nonempty_runs.first.drawing.graphic.data.grouping).to eq(:clustered)
    end

    it 'chart_in_table' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_in_table.docx')
      expect(docx.elements[1].rows.first.cells.first.elements.first.nonempty_runs.first.drawing.graphic.type).to eq(:chart)
    end

    it 'chart_minor_gridline' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_minor_gridline.docx')
      drawing = docx.elements.first.character_style_array.first.drawing
      expect(drawing.graphic.data.axises[1].major_grid_lines).to eq(false)
      expect(drawing.graphic.data.axises[1].minor_grid_lines).to eq(true)
    end

    it 'chart_point_1' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_point_1.docx')
      expect(docx.elements.first.character_style_array.first.drawing.graphic.data.type).to eq(:point)
      expect(docx.elements.first.character_style_array.first.drawing.graphic.data.grouping).to be_nil
    end

    it 'chart_point_2' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_point_2.docx')
      expect(docx.elements.first.character_style_array.first.drawing.graphic.data.type).to eq(:point)
      expect(docx.elements.first.character_style_array.first.drawing.graphic.data.grouping).to be_nil
    end

    it 'chart_points' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_points.docx')
      drawing = docx.elements.first.character_style_array.first.drawing
      expect(drawing.graphic.data.data.size).to eq(6)
      expect(drawing.graphic.data.data.first.points.size).to eq(3)
    end

    it 'chart_title' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_title.docx')
      expect(docx.elements.first.character_style_array.first.drawing.graphic.data.title.elements.first.characters.first.text).to eq('CustomChartTitle')
    end

    it 'chart_title_with_empty_paragraph' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/chart_title_with_empty_paragraph.docx')
      expect(docx.elements.first.character_style_array.first.drawing.graphic.data.title.elements.first.characters.empty?).to be_falsey
    end
  end

  describe 'offset' do
    it 'chart_offset_w14' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/offset/chart_offset_w14.docx')
      elements = docx.element_by_description(location: :canvas, type: :paragraph)
      expect(elements.first.nonempty_runs.first.drawing.properties.horizontal_position.offset).to be_zero
      expect(elements.first.nonempty_runs.first.drawing.properties.vertical_position.offset).to be_zero
    end

    it 'Check Diagram offset should !=0' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/offset/chart_offset_non_zero.docx')
      expect(docx.elements.first.character_style_array.first.drawing.properties.horizontal_position.offset).not_to be_zero
      expect(docx.elements.first.character_style_array.first.drawing.properties.vertical_position.offset).not_to be_zero
    end

    it 'Check Diagram offset should ==0' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/offset/chart_offset_zero.docx')
      expect(docx.elements.first.character_style_array.first.drawing.properties.horizontal_position.offset).to be_zero
      expect(docx.elements.first.character_style_array.first.drawing.properties.vertical_position.offset).to be_zero
    end
  end

  it 'Two charts in same paragraph' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/two_drawings_in_same_paragraph.docx')
    elements = docx.elements.first.nonempty_runs
    expect(elements.first.drawings[0].graphic.data.type).to eq(:doughnut)
    expect(elements.first.drawings[1].graphic.data.type).to eq(:pie)
  end

  it 'Add Shape to Group of Shapes in canvas into paragraph with data' do
    # spec/document_editor/one_user/smoke/content/shape_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/shape_grouping.docx')
    alternate = docx.elements.first.character_style_array.first.alternate_content
    expect(alternate.office2007_content.data.elements.length).to eq(2)
    expect(alternate.office2007_content.data.elements[1].type).to eq(:shape)
    expect(alternate.office2007_content.data.elements.first.elements.length).to eq(2)
    expect(alternate.office2007_content.data.elements.first.elements[0].type).to eq(:shape)
    expect(alternate.office2007_content.data.elements.first.elements[1].type).to eq(:shape)

    expect(alternate.office2010_content.graphic.data.elements.length).to eq(2)
    expect(alternate.office2010_content.graphic.data.elements[1]).to be_a(OoxmlParser::DocxShape)
    expect(alternate.office2010_content.graphic.data.elements.first.elements.length).to eq(2)
    expect(alternate.office2010_content.graphic.data.elements.first.elements[0]).to be_a(OoxmlParser::DocxShape)
    expect(alternate.office2010_content.graphic.data.elements.first.elements[1]).to be_a(OoxmlParser::DocxShape)
  end
end
