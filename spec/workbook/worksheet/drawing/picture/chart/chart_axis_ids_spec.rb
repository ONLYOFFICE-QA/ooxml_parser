require 'spec_helper'

describe 'Chart axis ids' do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/drawing/picture/chart/legend/legend_left_position.xlsx') }

  it 'Chart axis ids is an array' do
    expect(xlsx.worksheets.first.drawings.first.graphic_frame.graphic_data.first.axis_ids).to be_a(Array)
  end

  it 'Chart axis id has value' do
    expect(xlsx.worksheets.first.drawings.first.graphic_frame.graphic_data.first.axis_ids[0].value).to eq(1005)
  end
end
