require 'rspec'
require 'ooxml_parser'

describe 'Chart types' do
  it 'Chart Line 3d' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/Chart/types/chart-line-3d.xlsx')
    expect(xlsx.worksheets.first.drawings.first.picture.chart.type).to eq(:line_3d)
    expect(xlsx.worksheets.first.drawings.first.picture.chart.grouping).to eq(:standard)
  end
end
