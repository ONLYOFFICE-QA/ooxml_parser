require 'spec_helper'

describe 'My behaviour' do
  it 'Chart - Labels - chart_labels_position.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/labels/chart_labels_position.xlsx')
    expect(xlsx.worksheets[0].drawings[1].picture.chart.display_labels.position).to eq(:top)
  end
end
