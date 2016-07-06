require 'spec_helper'

describe 'My behaviour' do
  it 'table_color.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/table/properties/style/table_color.pptx')
    table = pptx.slides.first.graphic_frames.first.graphic_data.first
    expect(table.properties.style.first_row.cell_style.fill.color).to eq(OoxmlParser::Color.new(91, 155, 213))
  end
end
