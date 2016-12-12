require 'spec_helper'

describe 'font_size' do
  it 'font_size_float' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/size/font_size_float.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.size).to eq(10.5)
  end

  it 'font_size_not_string' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/size/font_size_not_string.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.size).to eq(11)
  end
end
