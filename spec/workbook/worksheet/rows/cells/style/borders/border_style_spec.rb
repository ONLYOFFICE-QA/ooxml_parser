require 'spec_helper'

describe 'border_srtyle' do
  it 'border_style_medium_dash' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/borders/style/border_style_medium_dash.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].style.borders.top.style).to eq(:medium_dashed)
  end
end
