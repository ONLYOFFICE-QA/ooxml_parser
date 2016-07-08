require 'spec_helper'

describe 'My behaviour' do
  it 'formula.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/text/string_reference/string_cache/string_cache_point_count.xlsx')
    series = xlsx.worksheets[0].drawings[0].picture.chart.series.first
    expect(series.text.string.cache.point_count.value).to eq(1)
  end
end
