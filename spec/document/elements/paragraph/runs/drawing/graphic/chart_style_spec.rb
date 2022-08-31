# frozen_string_literal: true

require 'spec_helper'

describe 'chart', '#style' do
  file = 'spec/document/elements/paragraph/runs/' \
         'drawing/graphic/chart/chart_with_style.docx'
  docx = OoxmlParser::Parser.parse(file)
  style = docx.elements[0].nonempty_runs.first.drawing.graphic.data.style

  it 'Chart has axis title' do
    expect(style.axis_title).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has category axis' do
    expect(style.category_axis).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has chart area' do
    expect(style.chart_area).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data label' do
    expect(style.data_label).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data label callout' do
    expect(style.data_label_callout).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data point' do
    expect(style.data_point).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data point 3d' do
    expect(style.data_point_3d).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data point line' do
    expect(style.data_point_line).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data point marker' do
    expect(style.data_point_marker).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data point wireframe' do
    expect(style.data_point_wireframe).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has data table' do
    expect(style.data_table).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has down bar' do
    expect(style.down_bar).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has drop line' do
    expect(style.drop_line).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has error bar' do
    expect(style.error_bar).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has floor' do
    expect(style.floor).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has gridline major' do
    expect(style.gridline_major).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has gridline minor' do
    expect(style.gridline_minor).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has high low line' do
    expect(style.high_low_line).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has leader line' do
    expect(style.leader_line).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has legend' do
    expect(style.legend).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has plot area' do
    expect(style.plot_area).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has plot area 3d' do
    expect(style.plot_area_3d).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has series axis' do
    expect(style.series_axis).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has series line' do
    expect(style.series_line).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has title' do
    expect(style.title).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has trend line' do
    expect(style.trend_line).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has trend line label' do
    expect(style.trend_line_label).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has up bar' do
    expect(style.up_bar).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has value axis' do
    expect(style.value_axis).to be_a(OoxmlParser::ChartStyleEntry)
  end

  it 'Chart has wall' do
    expect(style.wall).to be_a(OoxmlParser::ChartStyleEntry)
  end
end
