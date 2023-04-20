# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'categories.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/categories/categories.xlsx')
    series = xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.series.first
    expect(series.categories.reference.cache.point_count.value).to eq(25)
    expect(series.categories.reference.cache.points.first.index).to eq(0)
    expect(series.categories.reference.cache.points.first.text.value).to eq('USA')
  end

  it 'categories_number_reference.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/categories/categories_number_reference.xlsx')
    series = xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.series.first
    expect(series.categories.reference.formula).to eq('Sheet1!$B$1:$B$2')
  end
end
