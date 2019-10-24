# frozen_string_literal: true

require 'spec_helper'

describe 'Series#values' do
  it 'series_values.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/values/series_values.xlsx')
    series = xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.series.first
    expect(series.values.number_reference.cache.points.first.value.value).to eq('1')
  end
end
