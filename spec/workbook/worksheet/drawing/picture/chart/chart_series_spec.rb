require 'spec_helper'

describe 'My behaviour' do
  it 'chart_series_not_nil' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/series/chart_series_not_nil.xlsx')
    expect(xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.series).not_to be_empty
  end
end
