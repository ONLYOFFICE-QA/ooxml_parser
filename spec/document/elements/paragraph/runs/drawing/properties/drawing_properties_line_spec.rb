require 'spec_helper'

describe OoxmlParser::DocxShapeLine do
  it 'LinePropertiesDefault' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/line_properties_default.docx')
    expect(docx.elements[0].nonempty_runs.first.alternate_content.office2010_content
               .graphic.data.properties.line).to be_invisible
  end

  it 'ShapeLineEnding' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/shape_line_ending.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2010_content.graphic.data
               .properties.line.head_end.length).to eq(:medium)
  end

  it 'shape_line_custom_geometry.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/shape_line_custom_geometry.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2007_content.type).to eq(:shape)
    expect(alternate_content.office2010_content.graphic.data
               .properties.preset_geometry.name).to eq(:custom)
  end

  it 'Check shape line' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/shape_line_ending.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2007_content.type).to eq(:shape)
    expect(alternate_content.office2010_content.graphic.data
               .properties.preset_geometry.name).to eq(:line)
  end

  it 'ShapeLinePoints' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/shape_line_points.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2010_content.graphic.data
               .properties.custom_geometry.paths_list.first.elements.first.points.first).to eq(OoxmlParser::OOXMLCoordinates.new(5_440, 16_475))
  end

  it 'ShapeLineObjectSize' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/shape_line_object_size.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2010_content.properties.object_size).to eq(OoxmlParser::OOXMLCoordinates.new(OoxmlParser::OoxmlSize.new(11.03, :centimeter),
                                                                                                                OoxmlParser::OoxmlSize.new(12.54, :centimeter)))
  end

  it 'adjust_values' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/adjust_values.docx')
    expect(docx.notes.first.elements.first.nonempty_runs.first.alternate_content.office2010_content.graphic.data.properties.preset_geometry.adjust_values_list[0].formula).to eq('val 16667')
  end
end
