require 'spec_helper'

describe 'Drawing Properties Color' do
  it 'Set Gradient Fill for Shape in canvas into paragraph' do
    # spec/document_editor/one_user/smoke/content/shape_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/colors/shape_gradient_color.docx')
    expect(docx.elements.first.character_style_array[0]
               .alternate_content.office2010_content.graphic.data.properties.fill_color.type).to eq(:gradient)
    expect(docx.elements.first.character_style_array[0]
               .alternate_content.office2010_content.graphic.data.properties.fill_color.value.gradient_stops[0].color.value).to eq(OoxmlParser::Color.new(235, 241, 221))
    expect(docx.elements.first.character_style_array[0]
               .alternate_content.office2010_content.graphic.data.properties.fill_color.value
               .gradient_stops[0].position - 10).to be < 3
    expect(docx.elements.first.character_style_array[0]
               .alternate_content.office2010_content.graphic.data.properties.fill_color.value
               .gradient_stops[1].position - 70).to be < 3
  end

  it 'shape_insert_all_fill_picture.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/colors/shape_insert_all_fill_picture.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2010_content.graphic.data.properties.fill_color.type).to eq(:picture)
  end

  it 'shape_pattern_fill.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/colors/shape_pattern_fill.docx')
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

  it 'shape_gradient_color_angle.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/colors/shape_gradient_color_angle.docx')
    expect(docx.elements.first.character_style_array[1]
               .alternate_content.office2010_content.graphic.data.properties.fill_color.value.linear_gradient.angle).to eq(13.37)
  end

  it 'graphic_background_color_gradient.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/colors/graphic_background_color_gradient.docx')
    drawing = docx.element_by_description[1].nonempty_runs.first.drawing
    expect(drawing.graphic.data.shape_properties.fill.value.gradient_stops[0].color.converted_color).to eq(OoxmlParser::Color.new(237, 237, 237))
  end
end
