require 'spec_helper'

describe 'Margin' do
  it 'margin' do
    pptx = OoxmlParser::Parser.parse('spec/document/elements/table/cell/properties/margin/margin.pptx')
    expect(pptx.slides.first.elements.last.graphic_data.first.rows.first.cells.first.properties.margins.top).to eq(OoxmlParser::OoxmlSize.new(1.0, :centimeter))
  end
end
