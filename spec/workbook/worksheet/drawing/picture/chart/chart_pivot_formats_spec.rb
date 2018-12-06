require 'spec_helper'

describe 'My behaviour' do
  let(:graphic) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/drawing/picture/chart/pivot_formats/chart_with_pivot_formats.xlsx')
                       .worksheets[0].drawings.first.graphic_frame.graphic_data.first
  end

  it 'pivot formats has index' do
    expect(graphic.pivot_formats[2].index.value).to eq(2)
  end

  it 'pivot formats has data label with delete flag' do
    expect(graphic.pivot_formats[0].data_label.delete).to be_truthy
  end
end
