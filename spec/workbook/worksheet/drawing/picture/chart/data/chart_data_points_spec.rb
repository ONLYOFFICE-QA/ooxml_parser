# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'empty_num_ref' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/data/points/empty_num_ref.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.data.first.points.length).to eq(10)
  end

  it 'chart_data_points_3d_chart' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/data/points/chart_data_points_3d_chart.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.data.first.points.first.value).to eq(200.55)
  end

  it 'chart_data_points_empty' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/data/points/chart_data_points_empty.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.data).to be_empty
  end
end
