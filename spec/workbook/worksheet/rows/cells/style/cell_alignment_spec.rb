require 'spec_helper'

describe 'My behaviour' do
  it 'cell_style_compare_1' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_alignment/cell_align_default_horizontal.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].style.alignment.vertical).to eq(:bottom)
  end
end
