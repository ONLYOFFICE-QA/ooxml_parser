# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'table_look.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/properties/table_look/table_look.pptx')
    table_look = pptx.slides.first.elements.last.graphic_data[0].properties.table_look

    expect(table_look.first_column).to be(false)
    expect(table_look.first_row).to be(true)
    expect(table_look.last_column).to be(false)
    expect(table_look.last_row).to be(false)
    expect(table_look.banding_row).to be(false)
    expect(table_look.banding_column).to be(false)
  end
end
