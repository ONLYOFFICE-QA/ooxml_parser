require 'spec_helper'

describe 'My behaviour' do
  it 'string_cache_point_count.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/text/string_reference/string_cache/string_cache_point_count.xlsx')
    series = xlsx.worksheets[0].drawings[0].picture.chart.series.first
    expect(series.text.string.cache.point_count.value).to eq(1)
  end

  it 'point_index.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/text/string_reference/string_cache/point_index.xlsx')
    series = xlsx.worksheets[0].drawings[0].picture.chart.series.first
    expect(series.text.string.cache.points.first.index).to eq(0)
  end

  it 'point_index_text.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/text/string_reference/string_cache/point_index_text.xlsx')
    series = xlsx.worksheets[0].drawings[0].picture.chart.series.first
    expect(series.text.string.cache.points.first.text.value).to eq('Gold')
  end
end
