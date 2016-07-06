require 'spec_helper'

describe 'Chart types' do
  it 'Chart Line 3d' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/type/chart-line-3d.xlsx')
    expect(xlsx.worksheets.first.drawings.first.picture.chart.type).to eq(:line_3d)
    expect(xlsx.worksheets.first.drawings.first.picture.chart.grouping).to eq(:standard)
  end

  it 'Bar Clustered 3d' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/type/bar_clustered_3d.xlsx')
    expect(xlsx.worksheets.first.drawings.first.picture.chart.type).to eq(:bar_3d)
    expect(xlsx.worksheets.first.drawings.first.picture.chart.grouping).to eq(:clustered)
  end

  it 'Pie 3d' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/type/pie_3d.xlsx')
    expect(xlsx.worksheets.first.drawings.first.picture.chart.type).to eq(:pie_3d)
  end
end
