# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PlotArea do
  let(:plot_area) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/' \
                              'drawing/picture/chart/' \
                              'plot_area/plot_with_typese.xlsx')
  end

  it 'Check bubbleChart type' do
    expect(plot_area.worksheets[0]
                    .drawings.first
                    .graphic_frame
                    .graphic_data.first
                    .plot_area.bubble_chart).to be_a(OoxmlParser::CommonChartData)
  end

  it 'Check radarChart type' do
    expect(plot_area.worksheets[1]
                    .drawings.first
                    .graphic_frame
                    .graphic_data.first
                    .plot_area.radar_chart).to be_a(OoxmlParser::CommonChartData)
  end

  it 'Check surface3DChart type' do
    expect(plot_area.worksheets[2]
                    .drawings.first
                    .graphic_frame
                    .graphic_data.first
                    .plot_area.surface_3d_chart).to be_a(OoxmlParser::CommonChartData)
  end
end
