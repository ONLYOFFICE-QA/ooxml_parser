# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Chart - Labels - chart_labels_position.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/labels/chart_labels_position.xlsx')
    expect(xlsx.worksheets[0].drawings[1].graphic_frame.graphic_data.first.display_labels.position).to eq(:top)
  end

  it 'legend_show_category_name.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/legend/legend_show_category_name.xlsx')
    expect(xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.display_labels.show_category_name).to be_truthy
  end

  it 'legend_show_series_name.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/legend/legend_show_series_name.xlsx')
    expect(xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.display_labels.show_series_name).to be_truthy
  end
end
