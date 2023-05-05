# frozen_string_literal: true

require 'spec_helper'

describe 'Gradient Stop color properties' do
  docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/drawing/graphic/text_art/gradient_text.docx')
  color_properties = docx.elements.first.character_style_array.first.alternate_content.office2010_content.graphic
                         .data.text_body.elements.first.character_style_array.first.text_fill.color_scheme.color.gradient_stops[0].color.properties

  it 'color properties have luminance_modulation' do
    expect(color_properties.luminance_modulation).to eq(0.6)
  end

  it 'color properties have luminance_offset' do
    expect(color_properties.luminance_offset).to eq(0.4)
  end

  it 'color properties have tint' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/colors/shape_gradient_color.docx')
    expect(docx.elements.first.character_style_array[0]
               .alternate_content.office2010_content.graphic
               .data.properties.fill_color.value
               .gradient_stops[0].color.properties.tint).to eq(0.2)
  end

  it 'color properties have alpha' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/alternate_content/office2010_content/graphic/series/chart_series_no_values.docx')
    expect(docx.elements[3].character_style_array[18].drawings[0]
               .graphic.data.shape_properties.fill_color.value
               .gradient_stops[0].color.properties.alpha).to eq(100)
  end
end
