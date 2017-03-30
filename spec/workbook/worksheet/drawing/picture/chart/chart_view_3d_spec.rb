require 'spec_helper'

describe 'My behaviour' do
  it 'chart_series_not_nil' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/chart_3d_view/rotation_x.xlsx')
    expect(xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.view_3d.rotation_x.value).to eq(90)
  end
end
