# frozen_string_literal: true

require 'spec_helper'

describe 'Chart axis ids' do
  it 'Chart vary color is true' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/drawing/picture/chart/vary_colors/vary_colors_true.xlsx')
    expect(xlsx.worksheets.first.drawings.first.graphic_frame.graphic_data.first.vary_colors).to be_truthy
  end

  it 'Chart vary color is false' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/drawing/picture/chart/vary_colors/vary_colors_false.xlsx')
    expect(xlsx.worksheets.first.drawings.first.graphic_frame.graphic_data.first.vary_colors).to be_falsey
  end
end
