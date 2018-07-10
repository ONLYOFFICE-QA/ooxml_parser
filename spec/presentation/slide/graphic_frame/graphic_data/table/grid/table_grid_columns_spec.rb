require 'spec_helper'

describe 'TableGridColumns' do
  it 'table_grid_columns_width' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/graphic_frame/graphic_data/table/grid/columns/table_grid_columns_width.pptx')
    expect(pptx.slides.first.elements.last.graphic_data.first.grid.columns.first.width).to eq(OoxmlParser::OoxmlSize.new(4, :centimeter))
  end
end
