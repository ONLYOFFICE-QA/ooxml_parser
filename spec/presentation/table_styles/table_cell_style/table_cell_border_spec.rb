require 'spec_helper'

describe 'Table Cell Border' do
  it 'simple_table_cell_border.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/table_styles/table_cell_style/table_cell_border/simple_table_cell_border.pptx')
    expect(pptx.table_styles.table_style_list.first.whole_table.cell_style.borders_properties.top.fill.color).to be_a(OoxmlParser::SchemeColor)
  end
end
