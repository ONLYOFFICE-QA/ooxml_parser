# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ColorScale do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions'\
                              '/conditional_formattings/conditional_formatting_rule/color_scale/color_scale.xlsx')
  end
  let(:color_scale) { xlsx.worksheets.first.conditional_formattings[0].rules[0].color_scale }

  it 'Contains values' do
    expect(color_scale.values[0]).to be_a(OoxmlParser::ConditionalFormatValueObject)
  end

  it 'Contains colors' do
    expect(color_scale.colors[0]).to eq(OoxmlParser::Color.new(248, 105, 107))
  end

  it 'Value' do
    expect(color_scale.values[1].value).to eq('50')
  end
end
