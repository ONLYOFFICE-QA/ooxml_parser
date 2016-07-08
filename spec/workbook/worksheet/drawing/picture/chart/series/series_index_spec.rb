require 'spec_helper'

describe 'My behaviour' do
  it 'index.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/index/index.xlsx')
    series = xlsx.worksheets[0].drawings[0].picture.chart.series.first
    expect(series.index.value).to eq(0)
  end
end
