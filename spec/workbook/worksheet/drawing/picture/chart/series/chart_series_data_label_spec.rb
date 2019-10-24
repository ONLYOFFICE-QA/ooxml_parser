# frozen_string_literal: true

require 'spec_helper'

describe 'data_labels' do
  it 'chart_series_have_data_label' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/data_labels/chart_series_have_data_label.xlsx')
    series = xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.series.first
    expect(series.display_labels).to be_a(OoxmlParser::DisplayLabelsProperties)
  end
end
