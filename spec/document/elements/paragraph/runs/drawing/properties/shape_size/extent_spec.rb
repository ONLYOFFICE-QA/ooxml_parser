require 'spec_helper'

describe 'My behaviour' do
  it 'extent_x_with_default_accuracy' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/shape_size/extent/extent_x_with_default_accuracy.docx')
    elements = docx.element_by_description(location: :canvas, type: :paragraph)
    expect(elements[0].character_style_array[0].drawing.graphic.data.properties.shape_size.extent.x).to eq(OoxmlParser::OoxmlSize.new(16.5, :centimeter))
  end

  it 'extent_x_with_custom_accuracy' do
    OoxmlParser.configure do |config|
      config.accuracy = 3
    end

    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/shape_size/extent/extent_x_with_custom_accuracy.docx')
    elements = docx.element_by_description(location: :canvas, type: :paragraph)
    expect(elements[0].character_style_array[0].drawing.graphic.data.properties.shape_size.extent.x).to eq(OoxmlParser::OoxmlSize.new(16.501, :centimeter))
  end

  after do
    OoxmlParser.reset
  end
end
