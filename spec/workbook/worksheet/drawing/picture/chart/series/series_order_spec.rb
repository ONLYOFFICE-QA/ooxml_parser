require 'spec_helper'

describe 'My behaviour' do
  it 'order.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/order/order.xlsx')
    series = xlsx.worksheets[0].drawings[0].picture.chart.series.first
    expect(series.order.value).to eq(0)
  end
end
