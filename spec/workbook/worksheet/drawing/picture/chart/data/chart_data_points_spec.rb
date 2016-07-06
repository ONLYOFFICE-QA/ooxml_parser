require 'spec_helper'

describe 'My behaviour' do
  it 'empty_num_ref' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/data/points/empty_num_ref.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.picture.chart.data.first.points.length).to eq(10)
  end
end
