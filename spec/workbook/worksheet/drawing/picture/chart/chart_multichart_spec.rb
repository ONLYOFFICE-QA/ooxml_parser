require 'spec_helper'

describe 'My behaviour' do
  it 'multichart' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/multichart/multichart.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.plot_area.bar_chart).to be_a(OoxmlParser::CommonChartData)
    expect(drawing.graphic_frame.graphic_data.first.plot_area.line_chart).to be_a(OoxmlParser::CommonChartData)
  end
end
