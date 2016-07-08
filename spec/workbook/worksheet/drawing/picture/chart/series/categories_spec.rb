require 'spec_helper'

describe 'My behaviour' do
  it 'string_cache_point_count.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/categories/categroies.xlsx')
    series = xlsx.worksheets[0].drawings[0].picture.chart.series.first
    expect(series.categories.string.cache.point_count.value).to eq(25)
    expect(series.categories.string.cache.points.first.index).to eq(0)
    expect(series.categories.string.cache.points.first.text.value).to eq('USA')
  end
end
