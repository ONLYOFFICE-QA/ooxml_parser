# frozen_string_literal: true

require 'spec_helper'

describe 'Table Properties Fill' do
  it 'table_properties_fill.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/properties/fill/table_properties_fill.pptx')
    table = pptx.slides.first.graphic_frames.first.graphic_data.first
    expect(table.properties.fill.type).to eq(:solid)
    expect(table.properties.fill.color).to eq(OoxmlParser::Color.new(0, 255, 0))
  end
end
