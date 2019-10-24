# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Text Art | transform_arch_up.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/preset_text_wrap/transform_arch_up.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.properties.preset_text_warp.preset).to eq(:textArchUp)
  end

  it 'Text Art | transform_text_circle.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/preset_text_wrap/transform_text_circle.xlsx')
    expect(xlsx.worksheets[0].drawings.first.shape.text_body.properties.preset_text_warp.preset).to eq(:textCircle)
  end
end
