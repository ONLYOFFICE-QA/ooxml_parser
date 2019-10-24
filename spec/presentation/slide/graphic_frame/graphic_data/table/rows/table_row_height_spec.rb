# frozen_string_literal: true

require 'spec_helper'

describe 'TableRowHeight' do
  it 'table_row_height' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/graphic_frame/graphic_data/table/rows/height/table_row_heigth.pptx')
    expect(pptx.slides.first.elements.last.graphic_data.first.rows.first.height).to eq(OoxmlParser::OoxmlSize.new(1.36, :centimeter))
  end
end
